# -*- coding: utf-8 -*-
"""
================================================================================
 HBV 文献语义检索系统  ——  Streamlit实现
--------------------------------------------------------------------------------
 结构（对应页面三部分）：
   ① 数据加载：读取数据库
   ② 检索方式：关键词检索（AND 子串匹配） / 本体检索（OWL 语义扩展）
   ③ 结果展示：表格 + 分页（20 条/页）
--------------------------------------------------------------------------------
 运行：streamlit run app.py
================================================================================
"""

from __future__ import annotations

import os
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from math import ceil
from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st


# ==============================================================================
# ================================ ① 配置区 ====================================
# ==============================================================================
#
# 1. 文件路径
# 2. 列名映射
# 3. 拆分行数
# 4. 分页大小：PER_PAGE（默认 20）。
# ==============================================================================

BASE_DIR = Path(__file__).resolve().parent

EN_XLSX: Path = BASE_DIR /         # 英文文献 Excel
CN_XLSX: Path = BASE_DIR /         # 中文文献 Excel

EN_SHEET: str | int = 0        # 英文库 sheet，0 表示第一个 sheet
CN_SHEET: str | int = 0        # 中文库 sheet

SPLIT_ROW: int = 4659          # A 库 = 第 1 ~ 4659 个数据行
PER_PAGE: int = 20             # 每页条数

# pyarrow 兼容开关：st.dataframe 依赖 pyarrow，若环境里的 pyarrow 损坏
# （典型报错：ImportError: DLL load failed while importing lib），
# 代码会自动退化为 HTML 表格；也可把环境变量 LITSEARCH_NO_ARROW=1 强制走 HTML 表格。
FORCE_HTML_TABLE: bool = os.environ.get("LITSEARCH_NO_ARROW", "").strip().lower() in (
    "1", "true", "yes", "on"
)

# Database labels shown in the "Target databases" multiselect (order: A / B / C)
LIB_A = "A Database"
LIB_B = "B Database"
LIB_C = "C Database"

# 本体目录：HBPO 默认放 ontologies/HBPO.owl
ONTOLOGY_DIR: Path = BASE_DIR / "ontologies"
UPLOAD_DIR: Path = ONTOLOGY_DIR / "uploaded"
HBPO_CANDIDATES: list[Path] = [
    ONTOLOGY_DIR / "HBPO.owl",
    ONTOLOGY_DIR / "HBPO(HBV_protein_ontology).owl",
    BASE_DIR / "HBPO.owl",
    BASE_DIR / "HBPO(HBV_protein_ontology).owl",
]
ALLOWED_ONTOLOGY_SUFFIX = {".owl", ".ttl", ".rdf", ".xml", ".n3"}

# 参与检索的注解属性（局部名，小写比较）：首选标签 + 各类同义词。
# 这样本体里写进 oboInOwl:hasExactSynonym / skos:altLabel 的中文译名也能被搜索到。
LABEL_PROPERTIES: tuple[str, ...] = ("label", "preflabel", "title")
SYNONYM_PROPERTIES: tuple[str, ...] = (
    "hasexactsynonym",
    "hasrelatedsynonym",
    "hasnarrowsynonym",
    "hasbroadsynonym",
    "altlabel",
    "hassynonym",
    "synonym",
)
MATCHABLE_PROPERTIES: set[str] = {p.lower() for p in LABEL_PROPERTIES + SYNONYM_PROPERTIES}

# ---- 列名映射：系统标准名 -> 候选真实列名（从上到下依次匹配）---------------------
COLUMN_CANDIDATES: dict[str, list[str]] = {
    # 通用候选（中英文都可能出现）
    "id": ["id", "文章id", "文献id", "swxwzid", "记录id", "编号"],
    "title": ["title", "标题", "篇名", "题名", "文章标题", "titleenglish", "titlechinese"],
    "keywords": ["keywords", "keyword", "关键词", "主题词", "keywordenglish", "keywordchinese",
                 "subjectwordenglish", "subjectwordchinese"],
    "journal": ["journal", "期刊", "刊名", "来源", "发布期刊",
                "parenttitleenglish", "parenttitlechinese"],
    "year": ["year", "年份", "年", "出版年"],
    "author": ["author", "作者", "责任者", "第一作者", "全部作者",
               "responsiblepersonenglish", "responsiblepersonchinese", "responsibleperson"],
    "content": ["content", "摘要", "内容", "全文", "abstract",
                "contentenglish", "contentchinese"],
}

# ==============================================================================
# ============================== ② 数据层 ======================================
# ==============================================================================


class LibraryError(RuntimeError):
    """Excel 读取 / 列名映射相关错误。"""


def norm_text(value) -> str:
    """文本归一化：NFKC（全角转半角）+ 小写 + 压缩空白。
    中文库里有大量全角字符（ＰＣＲ、ＨＢｓＡｇ），归一化后中英文都能正常命中。
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    s = unicodedata.normalize("NFKC", str(value))
    s = s.replace("\u3000", " ").replace("\xa0", " ").replace("\u200b", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()


def _clean_cell(value) -> str:
    """单元格转干净字符串：空值 -> ''，去掉制表符/换行，压缩空白。"""
    try:
        if value is None or (not isinstance(value, (list, dict)) and pd.isna(value)):
            return ""
    except (TypeError, ValueError):
        pass
    s = unicodedata.normalize("NFKC", str(value))
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", s).strip()


def _clean_year(value) -> str:
    s = _clean_cell(value)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _resolve_columns(df: pd.DataFrame, lib_name: str) -> dict[str, str]:
    """把真实列名映射到系统标准名，返回 {标准名: 真实列名}。
    id / title 缺失时直接抛错；其余列缺失只告警（用空列兜底）。
    """
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    resolved: dict[str, str] = {}
    missing_required: list[str] = []
    missing_optional: list[str] = []

    for canonical, candidates in COLUMN_CANDIDATES.items():
        hit = None
        for cand in candidates:                       # 精确匹配
            if cand in df.columns:
                hit = cand
                break
        if hit is None:                               # 忽略大小写/空格再匹配一次
            for cand in candidates:
                key = cand.strip().lower()
                if key in norm_map:
                    hit = norm_map[key]
                    break
        if hit is not None:
            resolved[canonical] = hit
        elif canonical in ("id", "title"):
            missing_required.append(canonical)
        else:
            missing_optional.append(canonical)

    if missing_required:
        raise LibraryError(
            f"{lib_name}: missing required column(s) {missing_required}.\n"
            f"Columns found in the file: {list(df.columns)}\n"
            f"Please extend the candidate lists in COLUMN_CANDIDATES at the top of app.py."
        )
    if missing_optional:
        st.warning(
            f"{lib_name}: column(s) {missing_optional} not found — these fields will be empty. "
            f"Columns found in the file: {list(df.columns)}"
        )
    return resolved


@st.cache_data(show_spinner=False)
def load_library(path_str: str, sheet: str | int, lib_name: str) -> pd.DataFrame:
    """读取一个 Excel 文件并标准化为内部结构（带缓存）。

    返回列：id / title / author / keywords / journal / year / content / _title_kw / _all
    """
    path = Path(path_str)
    if not path.exists():
        raise LibraryError(
            f"Data file not found: {path}\n"
            f"Place the file at this path or update EN_XLSX / CN_XLSX at the top of app.py."
        )

    try:
        raw = pd.read_excel(path, sheet_name=sheet, dtype=str)
    except ValueError as exc:                          # sheet 名不存在
        raise LibraryError(f"{path.name}: failed to read sheet '{sheet}' — {exc}") from exc
    except Exception as exc:                           # 文件损坏 / 被占用等
        raise LibraryError(f"{path.name}: failed to read the file — {exc}") from exc

    if raw.empty:
        raise LibraryError(f"{path.name}: the file contains no data.")

    cols = _resolve_columns(raw, lib_name)

    def col(name: str) -> pd.Series:
        if name in cols:
            return raw[cols[name]].map(_clean_cell)
        return pd.Series([""] * len(raw), index=raw.index, dtype="object")

    df = pd.DataFrame(
        {
            "id": col("id"),
            "title": col("title"),
            "author": col("author"),
            "keywords": col("keywords"),
            "journal": col("journal"),
            "content": col("content"),
        }
    )
    df["year"] = raw[cols["year"]].map(_clean_year) if "year" in cols else ""

    # 空值兜底：任何一列为空都不会导致后续匹配/展示报错
    for c in ("id", "title", "author", "keywords", "journal", "year", "content"):
        df[c] = df[c].fillna("").astype(str)

    # 预计算两套检索文本：
    #   _title_kw → 标题 + 关键词（默认检索范围）
    #   _all      → 标题 + 关键词 + 摘要（勾选 "Include article content" 时使用）
    df["_title_kw"] = (df["title"] + " " + df["keywords"]).map(norm_text)
    df["_all"] = (df["title"] + " " + df["keywords"] + " " + df["content"]).map(norm_text)

    df = df.reset_index(drop=True)
    df["library"] = lib_name
    return df


def build_libraries() -> dict[str, pd.DataFrame]:
    """读取两个 Excel 并按 SPLIT_ROW 拆成 A / B / C 三个逻辑库。"""
    en = load_library(str(EN_XLSX), EN_SHEET, "英文库")
    cn = load_library(str(CN_XLSX), CN_SHEET, "中文库")

    lib_a = en.iloc[:SPLIT_ROW].copy().reset_index(drop=True)
    lib_b = en.iloc[SPLIT_ROW:].copy().reset_index(drop=True)
    lib_c = cn.copy().reset_index(drop=True)

    lib_a["library"] = LIB_A
    lib_b["library"] = LIB_B
    lib_c["library"] = LIB_C
    return {LIB_A: lib_a, LIB_B: lib_b, LIB_C: lib_c}


def combine_libraries(libs: dict[str, pd.DataFrame], selected: Iterable[str]) -> pd.DataFrame:
    """把用户勾选的库纵向合并；index 重置避免重复索引导致 mask 错位。"""
    frames = [libs[name] for name in selected if name in libs]
    if not frames:
        return pd.DataFrame(columns=["id", "title", "author", "keywords", "journal", "year",
                                     "content", "_title_kw", "_all", "library"])
    return pd.concat(frames, ignore_index=True)


# ==============================================================================
# ============================== ③ 检索层 ======================================
# ==============================================================================

def keyword_search(df: pd.DataFrame, query: str, include_content: bool = False) -> pd.DataFrame:
    """关键词检索：空格分隔的多词按 AND 逻辑做子串匹配（不区分大小写）。
    include_content=False（默认）只匹配 title / keywords；
    include_content=True 时把摘要 content 一并纳入匹配（对应页面上的
    “Include article content (abstract)” 选项）。
    """
    terms = norm_text(query).split()
    if not terms:
        return df.iloc[0:0]

    field = "_all" if include_content else "_title_kw"
    if field not in df.columns:                        # 兼容旧缓存/外部构造的 DataFrame
        field = "_title_kw"

    mask = pd.Series(True, index=df.index)
    for term in terms:
        # norm_text 已经把文档侧也归一化，这里再走一次保证全半角一致
        mask &= df[field].str.contains(norm_text(term), regex=False, na=False)
    return df[mask]

# ------------------------------------------------------------------ 本体索引 ----
class OntologyError(RuntimeError):
    """本体加载失败。"""


@dataclass
class OntologyIndex:
    """轻量本体索引：只保留概念名、标签和三类关系，够做语义扩展即可。"""
    name: str
    path: str
    loader: str = ""
    nodes: dict[str, dict] = field(default_factory=dict)                 # key -> {name, labels}
    parents: dict[str, set] = field(default_factory=lambda: defaultdict(set))
    children: dict[str, set] = field(default_factory=lambda: defaultdict(set))
    equivalents: dict[str, set] = field(default_factory=lambda: defaultdict(set))
    same_as: dict[str, set] = field(default_factory=lambda: defaultdict(set))

    # ---------------- 构建 ----------------
    def add_node(self, key: str, name: str, labels: Iterable[str] = ()) -> None:
        if not key:
            return
        node = self.nodes.setdefault(key, {"name": name or key, "labels": []})
        for lb in labels:
            lb = str(lb).strip()
            if lb and lb not in node["labels"]:
                node["labels"].append(lb)

    def add_edge(self, child: str, parent: str) -> None:
        if not child or not parent or child == parent:
            return
        self.children[parent].add(child)
        self.parents[child].add(parent)

    def add_equivalent(self, a: str, b: str) -> None:
        if not a or not b or a == b:
            return
        self.equivalents[a].add(b)
        self.equivalents[b].add(a)

    def add_same_as(self, a: str, b: str) -> None:
        if not a or not b or a == b:
            return
        self.same_as[a].add(b)
        self.same_as[b].add(a)

    # ---------------- 名称 / 变体 ----------------
    def display(self, key: str) -> str:
        return self.nodes.get(key, {}).get("name", key)

    def surface_forms(self, key: str) -> tuple[set[str], set[str]]:
        """返回。主干变体用于更宽松的文献匹配。"""
        node = self.nodes.get(key)
        if not node:
            return set(), set()

        strict: set[str] = set()
        for raw in [node["name"], *node["labels"]]:
            n = norm_text(raw)
            if not n:
                continue
            strict.add(n)
            strict.add(n.replace("_", " "))
            strict.add(n.replace("-", " "))
            strict.add(n.replace("_", " ").replace("-", " "))

        trunk: set[str] = set()
        base = norm_text(node["name"])
        # envelope_of_HBV -> envelope ; pre-genome_of_HBV -> pre genome
        stem = re.sub(r"_(of|in|to)_hbv$", "", base)
        if stem and stem != base:
            stem = stem.replace("_", " ").replace("-", " ")
            if len(stem) >= 4:
                trunk.add(stem)
        return strict, trunk

    # ---------------- 概念映射 ----------------
    def match(self, query: str, limit: int = 12) -> list[str]:
        """把用户输入映射到本体概念：先精确匹配，再退回包含匹配。"""
        q = norm_text(query)
        if not q:
            return []

        exact: list[str] = []
        partial: list[str] = []
        for key in self.nodes:
            strict, _ = self.surface_forms(key)
            if q in strict:
                exact.append(key)
                continue
            for form in strict:
                if len(form) >= 3 and (q in form or form in q):
                    partial.append(key)
                    break

        if exact:
            return exact[:limit]
        # 按名称长度排序，优先更具体的概念
        partial.sort(key=lambda k: (len(self.display(k)), self.display(k)))
        return partial[:limit]

    def expand(self, keys: Iterable[str], depth: int = 3, include_parents: bool = False) -> list[str]:
        """沿 subClassOf / equivalentClass / sameAs 做语义扩展。"""
        keys = [k for k in dict.fromkeys(keys) if k]
        included: list[str] = list(keys)
        seen: set[str] = set(keys)

        def walk(start: list[str], downward: bool) -> None:
            frontier = list(start)
            for _ in range(max(1, int(depth))):
                nxt: list[str] = []
                for k in frontier:
                    links = set(self.equivalents.get(k, set())) | set(self.same_as.get(k, set()))
                    links |= set(self.children.get(k, set()) if downward else self.parents.get(k, set()))
                    for nb in sorted(links):
                        if nb not in seen:
                            seen.add(nb)
                            nxt.append(nb)
                            included.append(nb)
                if not nxt:
                    break
                frontier = nxt

        walk(keys, downward=True)
        if include_parents:
            walk(keys, downward=False)
        return included


# ---------------------------------------------------------------- 本体加载器 ----
def _local_name(iri: str) -> str:
    iri = (iri or "").strip()
    if not iri:
        return ""
    if "#" in iri:
        return iri.rsplit("#", 1)[1]
    return iri.rstrip("/").rsplit("/", 1)[-1]


def _load_with_owlready2(path: Path) -> OntologyIndex:
    """主加载器：owlready2 支持 RDF/XML、OWL/XML、Turtle、NTriples。"""
    from owlready2 import ThingClass, World  # 延迟导入，缺失时给友好提示

    index = OntologyIndex(name=path.stem, path=str(path), loader="owlready2")
    cwd = os.getcwd()
    # owlready2 在 Windows 下对 file:// URI 支持不佳，切换到文件目录后按文件名加载
    os.chdir(str(path.parent))
    try:
        world = World()
        onto = world.get_ontology(path.name).load()
    finally:
        os.chdir(cwd)

    entities: dict[object, str] = {}
    for cls in onto.classes():
        key = cls.name
        if not key:
            continue
        # rdfs:label(s) 作为首选术语，同义词（oboInOwl / skos）一并纳入匹配词表
        labels = [str(lb) for lb in (getattr(cls, "label", []) or [])]
        synonyms: list[str] = []
        for prop in ("prefLabel", "hasExactSynonym", "hasRelatedSynonym", "hasNarrowSynonym",
                     "hasBroadSynonym", "altLabel", "hasSynonym"):
            for lb in (getattr(cls, prop, None) or []):
                synonyms.append(str(lb))
        name = labels[0] if labels else key
        index.add_node(key, name, [*labels, *synonyms, key])
        entities[cls] = key

    for cls in onto.classes():
        key = entities.get(cls)
        if not key:
            continue
        for parent in getattr(cls, "is_a", []) or []:
            if isinstance(parent, ThingClass) and parent in entities:
                index.add_edge(key, entities[parent])
        for equiv in getattr(cls, "equivalent_to", []) or []:
            if isinstance(equiv, ThingClass) and equiv in entities:
                index.add_equivalent(key, entities[equiv])
        # 个别版本把 owl:sameAs 暴露在 identical_to 上（主要针对个体）
        for same in getattr(cls, "identical_to", []) or []:
            if isinstance(same, ThingClass) and same in entities:
                index.add_same_as(key, entities[same])

    # owl:sameAs 兜底：直接在 RDF 图上扫描一次（失败则忽略）
    try:
        from rdflib import URIRef

        graph = getattr(world, "graph", None)
        registry = getattr(world, "_entities", {})
        same_pred = URIRef("http://www.w3.org/2002/07/owl#sameAs")
        if graph is not None:
            for s, _, o in graph.triples((None, same_pred, None)):
                a, b = registry.get(s), registry.get(o)
                ka = entities.get(a) if a is not None else None
                kb = entities.get(b) if b is not None else None
                if ka and kb:
                    index.add_same_as(ka, kb)
    except Exception:
        pass

    if not index.nodes:
        raise OntologyError("owlready2 parsed the file but found no owl:Class (concept).")
    return index


def _load_with_owlxml(path: Path) -> OntologyIndex:
    """兜底加载器 1：纯标准库解析 OWL/XML（Protege 导出的 .owl 常见此格式）。"""
    import xml.etree.ElementTree as ET

    OWL = "{http://www.w3.org/2002/07/owl#}"
    RDF = "{http://www.w3.org/1999/02/22-rdf-syntax-ns#}"

    try:
        root = ET.parse(str(path)).getroot()
    except Exception as exc:
        raise OntologyError(f"XML parsing failed: {exc}") from exc

    index = OntologyIndex(name=path.stem, path=str(path), loader="owlxml(stdlib)")

    def iri_of(elem) -> str:
        for attr in ("IRI", f"{RDF}about", f"{RDF}resource", f"{RDF}ID"):
            val = elem.get(attr)
            if val:
                return _local_name(val)
        return ""

    # 1) 类声明
    for decl in root.findall(f"{OWL}Declaration"):
        cls = decl.find(f"{OWL}Class")
        if cls is not None:
            key = iri_of(cls)
            index.add_node(key, key)

    # 2) 子类关系（OWL/XML 中第一个 Class 是子类，第二个是父类）
    for axiom in root.findall(f"{OWL}SubClassOf"):
        kids = [iri_of(c) for c in axiom.findall(f"{OWL}Class")]
        kids = [k for k in kids if k]
        if len(kids) >= 2:
            index.add_node(kids[0], kids[0])
            index.add_node(kids[1], kids[1])
            index.add_edge(kids[0], kids[1])
        elif len(kids) == 1:
            index.add_node(kids[0], kids[0])

    # 3) 等价类 / sameAs
    for tag, adder in ((f"{OWL}EquivalentClasses", index.add_equivalent),
                       (f"{OWL}SameIndividual", index.add_same_as)):
        for axiom in root.findall(tag):
            members = [iri_of(c) for c in axiom]
            members = [m for m in members if m]
            for other in members[1:]:
                adder(members[0], other)

    # 4) rdfs:label / oboInOwl:hasExactSynonym / skos:altLabel 等注解
    label_keys = {p.lower() for p in LABEL_PROPERTIES}
    annos: dict[str, dict[str, list[str]]] = {}
    for ann in root.findall(f"{OWL}AnnotationAssertion"):
        prop = ann.find(f"{OWL}AnnotationProperty")
        subject = ann.find(f"{OWL}IRI")
        value = ann.find(f"{OWL}Literal")
        if prop is None or subject is None or value is None or not value.text:
            continue
        pname = iri_of(prop).lower()
        if pname not in MATCHABLE_PROPERTIES:
            continue
        key = _local_name(subject.text or "")
        if not key:
            continue
        bucket = annos.setdefault(key, {"label": [], "syn": []})
        bucket["label" if pname in label_keys else "syn"].append(value.text)

    for key, bucket in annos.items():
        labels, synonyms = bucket["label"], bucket["syn"]
        if key in index.nodes:
            index.nodes[key]["name"] = labels[0] if labels else key
        index.add_node(key, labels[0] if labels else key, [*labels, *synonyms, key])

    if not index.nodes:
        raise OntologyError("No owl:Class found in this OWL/XML file.")
    return index


def _load_with_rdflib(path: Path) -> OntologyIndex:
    """兜底加载器 2：rdflib 通用 RDF 解析（RDF/XML、Turtle、N3…）。"""
    try:
        from rdflib import Graph, Namespace, RDF, RDFS, OWL
    except ImportError as exc:
        raise OntologyError("rdflib is not installed, cannot parse this format.") from exc

    graph = Graph()
    suffix = path.suffix.lower()
    fmt = {".ttl": "turtle", ".n3": "n3", ".rdf": "xml", ".xml": "xml", ".owl": "xml"}.get(suffix, "xml")
    try:
        graph.parse(str(path), format=fmt)
    except Exception as exc:
        raise OntologyError(f"rdflib parsing failed (format={fmt}): {exc}") from exc

    index = OntologyIndex(name=path.stem, path=str(path), loader=f"rdflib({fmt})")

    def key_of(term) -> str:
        return _local_name(str(term))

    label_keys = {p.lower() for p in LABEL_PROPERTIES}
    for subj in graph.subjects(RDF.type, OWL.Class):
        key = key_of(subj)
        labels: list[str] = []
        synonyms: list[str] = []
        for pred in graph.predicates(subj, None):
            pname = key_of(pred).lower()
            if pname not in MATCHABLE_PROPERTIES:
                continue
            values = [str(o) for o in graph.objects(subj, pred)]
            (labels if pname in label_keys else synonyms).extend(values)
        index.add_node(key, labels[0] if labels else key, [*labels, *synonyms, key])

    for subj, _, obj in graph.triples((None, RDFS.subClassOf, None)):
        if (obj, RDF.type, OWL.Class) in graph or isinstance(obj, Namespace):
            index.add_edge(key_of(subj), key_of(obj))
    for subj, _, obj in graph.triples((None, OWL.equivalentClass, None)):
        index.add_equivalent(key_of(subj), key_of(obj))
    for subj, _, obj in graph.triples((None, OWL.sameAs, None)):
        index.add_same_as(key_of(subj), key_of(obj))

    if not index.nodes:
        raise OntologyError("No owl:Class found in the RDF graph.")
    return index


def load_ontology(path_str: str) -> OntologyIndex:
    """依次尝试 owlready2 → OWL/XML(标准库) → rdflib，全部失败抛 OntologyError。"""
    path = Path(path_str)
    if not path.exists():
        raise OntologyError(f"Ontology file not found: {path}")
    if path.suffix.lower() not in ALLOWED_ONTOLOGY_SUFFIX:
        raise OntologyError(f"Unsupported ontology format: {path.suffix} (allowed: .owl/.ttl/.rdf/.xml/.n3)")

    errors: list[str] = []
    for loader in (_load_with_owlready2, _load_with_owlxml, _load_with_rdflib):
        try:
            return loader(path)
        except Exception as exc:                       # noqa: BLE001 - 逐个兜底
            errors.append(f"{loader.__name__}: {exc}")

    raise OntologyError("Failed to load the ontology; all parsers failed://n" + "\n".join(errors))


@st.cache_resource(show_spinner=False)
def get_ontology_index(path_str: str, mtime: float) -> OntologyIndex:
    """缓存本体索引；mtime 参与缓存键，文件被替换后会自动重新加载。"""
    return load_ontology(path_str)


def find_hbpo() -> Path | None:
    for p in HBPO_CANDIDATES:
        if p.exists():
            return p
    return None


# ---------------------------------------------------------- 本体检索实现 ----
def _build_concept_pattern(forms: Iterable[str]) -> str:
    """把概念的所有写法拼成一个正则：短英文词加词边界，避免 hbe 命中 the。"""
    parts: set[str] = set()
    for form in forms:
        form = (form or "").strip()
        if len(form) < 2:
            continue
        if form.isascii() and len(form) <= 4:
            parts.add(rf"(?<![a-z0-9]){re.escape(form)}(?![a-z0-9])")
        else:
            parts.add(re.escape(form))
    return "|".join(sorted(parts))


def ontology_search(
    df: pd.DataFrame,
    index: OntologyIndex,
    query: str,
    depth: int = 3,
    include_parents: bool = False,
    loose: bool = True,
) -> dict:
    """本体检索：概念映射 → 语义扩展 → 用扩展概念匹配 title / keywords。"""
    matched_keys = index.match(query)
    expanded_keys = index.expand(matched_keys, depth=depth, include_parents=include_parents) if matched_keys else []

    concept_masks: dict[str, pd.Series] = {}
    for key in expanded_keys:
        strict, trunk = index.surface_forms(key)
        forms = set(strict) | (set(trunk) if loose else set())
        pattern = _build_concept_pattern(forms)
        if not pattern:
            continue
        try:
            mask = df["_title_kw"].str.contains(pattern, regex=True, na=False)
        except re.error:
            mask = df["_title_kw"].str.contains(re.escape(pattern), regex=False, na=False)
        if mask.any():
            concept_masks[index.display(key)] = mask

    overall = pd.Series(False, index=df.index)
    for mask in concept_masks.values():
        overall |= mask

    result = df[overall].copy()
    if not result.empty:
        result["命中概念"] = [
            "；".join(name for name, mask in concept_masks.items() if mask.loc[idx])
            for idx in result.index
        ]
    else:
        result["命中概念"] = pd.Series(dtype="object")

    return {
        "matched_keys": matched_keys,
        "expanded_keys": expanded_keys,
        "matched_names": [index.display(k) for k in matched_keys],
        "expanded_names": [index.display(k) for k in expanded_keys],
        "hit_concepts": list(concept_masks),
        "result": result,
    }


# ==============================================================================
# ================================ UI LAYER ====================================
# ==============================================================================

MODE_KEYWORD = "Keyword Search"
MODE_ONTOLOGY = "Ontology Search"

RESULT_COLUMNS = ["Article ID", "Title", "Author", "Keywords", "Journal", "Year"]

_ARROW_STATE: bool | None = None

_HTML_TABLE_CSS = """
<style>
.lit-table-wrap {max-height: 660px; overflow: auto; border: 1px solid rgba(49,51,63,.15);
                 border-radius: 8px; margin-bottom: .5rem;}
table.lit-table {border-collapse: collapse; width: 100%; font-size: 13px; line-height: 1.5;}
table.lit-table th {position: sticky; top: 0; background: #f6f6f8; text-align: left;
                    padding: 8px 10px; border-bottom: 1px solid rgba(49,51,63,.18);
                    font-weight: 600; white-space: nowrap;}
table.lit-table td {padding: 8px 10px; border-bottom: 1px solid rgba(49,51,63,.08);
                    vertical-align: top; word-break: break-word;}
table.lit-table tr:nth-child(even) td {background: #fcfcfd;}
</style>
"""


def arrow_available() -> bool:
    """st.dataframe 依赖 pyarrow。部分环境（如 Anaconda 自带的 pyarrow DLL 损坏）
    导入时会抛 ImportError，这里检测一次并退化为 HTML 表格，保证页面可用。"""
    global _ARROW_STATE
    if FORCE_HTML_TABLE:
        return False
    if _ARROW_STATE is None:
        try:
            import pyarrow  # noqa: F401
            _ARROW_STATE = True
        except Exception:                                   # noqa: BLE001
            _ARROW_STATE = False
    return bool(_ARROW_STATE)


def render_table(display: pd.DataFrame) -> None:
    """渲染结果表格：优先 st.dataframe，pyarrow 不可用时退化为 HTML 表格。"""
    if arrow_available():
        st.dataframe(
            display,
            hide_index=True,
            width="stretch",
            height=min(38 + 35 * len(display), 720),
            column_config={
                "Article ID": st.column_config.TextColumn(width="small"),
                "Title": st.column_config.TextColumn(width="large"),
                "Author": st.column_config.TextColumn(width="medium"),
                "Keywords": st.column_config.TextColumn(width="medium"),
                "Journal": st.column_config.TextColumn(width="medium"),
                "Year": st.column_config.TextColumn(width="small"),
            },
        )
        return

    st.markdown(_HTML_TABLE_CSS + '<div class="lit-table-wrap">'
                + display.to_html(index=False, escape=True, border=0, classes="lit-table")
                + "</div>", unsafe_allow_html=True)


def init_state() -> None:
    defaults = {
        "search_mode": MODE_KEYWORD,
        "page": 1,
        "last_result": None,
        "active_ontology": None,
        "ontologies": {},
        "saved_uploads": set(),
        "pending_mode": None,
        "unmapped": None,
        "_mode_seen": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def register_ontologies() -> None:
    """Build the ontology registry: default HBPO plus previously uploaded files."""
    registry: dict[str, str] = {}
    hbpo = find_hbpo()
    if hbpo is not None:
        registry["HBPO"] = str(hbpo)

    if UPLOAD_DIR.exists():
        for file in sorted(UPLOAD_DIR.glob("*")):
            if file.suffix.lower() in ALLOWED_ONTOLOGY_SUFFIX:
                registry.setdefault(file.stem, str(file))

    st.session_state["ontologies"] = registry
    if st.session_state.get("active_ontology") not in registry:
        st.session_state["active_ontology"] = next(iter(registry), None)


def save_uploaded_ontology(uploaded) -> None:
    """Persist the uploaded ontology and register it in the dropdown (once per file)."""
    if uploaded is None:
        return
    suffix = Path(uploaded.name).suffix.lower()
    if suffix not in ALLOWED_ONTOLOGY_SUFFIX:
        st.error(
            f"Unsupported file type: {uploaded.name}. "
            f"Allowed: {', '.join(sorted(ALLOWED_ONTOLOGY_SUFFIX))}."
        )
        return

    key = f"{uploaded.name}:{uploaded.size}"
    if key in st.session_state["saved_uploads"]:
        return

    try:
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        target = UPLOAD_DIR / Path(uploaded.name).name
        target.write_bytes(uploaded.getbuffer())
    except Exception as exc:                            # noqa: BLE001
        st.error(f"Failed to save the uploaded file: {exc}")
        return

    st.session_state["saved_uploads"].add(key)
    name = target.stem
    st.session_state["ontologies"][name] = str(target)
    st.session_state["active_ontology"] = name
    st.session_state["ont_choice"] = name
    st.success(f"Imported ontology: {target.name} — added to the dropdown and selected.")


def show_concept_line(names: list[str], matched: list[str], max_inline: int = 24) -> None:
    """Show the expanded concept set above the result table."""
    if not names:
        return
    inline = names[:max_inline]
    text = ", ".join(f"`{n}`" for n in inline)
    if len(names) > max_inline:
        text += f" … ({len(names)} concepts in total)"
    st.markdown(f"**Concepts expanded by this ontology:** {text}")
    if matched:
        st.caption("Directly mapped concept(s): " + ", ".join(matched))
    if len(names) > max_inline:
        with st.expander("Show all expanded concepts"):
            st.write(", ".join(names))


def render_results(show_source: bool = False, show_hit: bool = False) -> None:
    payload = st.session_state.get("last_result")
    if not payload:
        if not st.session_state.get("unmapped"):
            st.info("No search has been run yet. Choose a search mode above and enter your query.")
        return

    df = payload["result"]
    total = len(df)
    fields_note = ""
    if payload["mode"] == MODE_KEYWORD:
        fields_note = " ｜ Fields: " + (
            "title + keywords + abstract" if payload.get("include_content") else "title + keywords"
        )
    st.caption(
        f"Mode: {payload['mode']} ｜ Query: `{payload['query']}` ｜ "
        f"Scope: {', '.join(payload['libraries'])}{fields_note} ｜ {payload['elapsed']:.2f}s"
    )

    if payload["mode"] == MODE_ONTOLOGY:
        show_concept_line(payload["expanded_names"], payload["matched_names"])

    if total == 0:
        st.warning("No matching literature found. Please adjust your keywords or search scope.")
        return

    pages = max(1, ceil(total / PER_PAGE))
    page = min(max(1, int(st.session_state["page"])), pages)
    st.session_state["page"] = page

    start = (page - 1) * PER_PAGE
    view = df.iloc[start:start + PER_PAGE]

    display = pd.DataFrame(
        {
            "Article ID": view["id"],
            "Title": view["title"],
            "Author": view["author"],
            "Keywords": view["keywords"],
            "Journal": view["journal"],
            "Year": view["year"],
        }
    ).fillna("")
    if show_source and len(payload["libraries"]) > 1:
        display["Source DB"] = view["library"].fillna("")
    if show_hit and payload["mode"] == MODE_ONTOLOGY:
        display["Matched Concept"] = view["命中概念"].fillna("")

    st.markdown(
        f"**{total:,} results**　(Page {page} of {pages} ｜ {PER_PAGE} per page)"
    )
    render_table(display)

    col_prev, col_info, col_next = st.columns([1, 2, 1])
    if col_prev.button("⬅ Previous", disabled=page <= 1, width="stretch", key="btn_prev"):
        st.session_state["page"] = page - 1
        st.rerun()
    col_info.markdown(
        f"<div style='text-align:center'>Showing rows {start + 1}–"
        f"{min(start + PER_PAGE, total)} of {total:,}</div>",
        unsafe_allow_html=True,
    )
    if col_next.button("Next ➡", disabled=page >= pages, width="stretch", key="btn_next"):
        st.session_state["page"] = page + 1
        st.rerun()

    csv = display.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇ Export current page as CSV",
        data=csv,
        file_name=f"literature_page{page}.csv",
        mime="text/csv",
        key="btn_csv",
    )


def run_keyword_search(
    libs: dict[str, pd.DataFrame],
    selected: list[str],
    query: str,
    include_content: bool = False,
) -> None:
    import time

    df = combine_libraries(libs, selected)
    started = time.perf_counter()
    with st.spinner("Searching…"):
        result = keyword_search(df, query, include_content=include_content)
    elapsed = time.perf_counter() - started

    st.session_state["unmapped"] = None
    st.session_state["last_result"] = {
        "mode": MODE_KEYWORD,
        "query": query.strip(),
        "result": result,
        "libraries": selected,
        "elapsed": elapsed,
        "expanded_names": [],
        "matched_names": [],
        "include_content": include_content,
    }
    st.session_state["page"] = 1


def run_ontology_search(
    libs: dict[str, pd.DataFrame],
    selected: list[str],
    query: str,
    ontology_path: str,
    depth: int,
    include_parents: bool,
    loose: bool,
) -> None:
    import time

    df = combine_libraries(libs, selected)
    started = time.perf_counter()
    with st.spinner("Loading ontology and expanding concepts…"):
        try:
            mtime = Path(ontology_path).stat().st_mtime
            index = get_ontology_index(ontology_path, mtime)
        except Exception as exc:                        # noqa: BLE001
            st.session_state["last_result"] = None
            st.error(f"Failed to load the ontology: {exc}")
            return

        outcome = ontology_search(
            df, index, query, depth=depth, include_parents=include_parents, loose=loose
        )
    elapsed = time.perf_counter() - started

    if not outcome["matched_keys"]:
        st.session_state["last_result"] = None
        st.session_state["unmapped"] = {
            "query": query.strip(),
            "ontology": index.name,
            "samples": [index.display(k) for k in list(index.nodes)[:24]],
        }
        return

    st.session_state["unmapped"] = None
    st.session_state["last_result"] = {
        "mode": MODE_ONTOLOGY,
        "query": query.strip(),
        "result": outcome["result"],
        "libraries": selected,
        "elapsed": elapsed,
        "expanded_names": outcome["expanded_names"],
        "matched_names": outcome["matched_names"],
        "ontology": index.name,
        "loader": index.loader,
    }
    st.session_state["page"] = 1


def apply_pending_mode() -> None:
    """Apply a queued mode switch before the radio widget is created."""
    pending = st.session_state.get("pending_mode")
    if pending:
        st.session_state["search_mode"] = pending
        st.session_state["pending_mode"] = None


def render_sidebar_header() -> None:
    """页面标题放在侧边栏顶部，主区域直接进入检索流程。"""
    with st.sidebar:
        st.title("HBV Literature Search")
        st.caption(
            "Search the English and Chinese HBV literature collections "
            "by keyword or by ontology concept."
        )
        st.divider()


def render_sidebar() -> None:
    with st.sidebar:
        st.header("Ontology")
        registry = st.session_state.get("ontologies") or {}
        if registry:
            for name, path in registry.items():
                st.caption(f"{name} → {Path(path).name}")
        else:
            st.caption("No ontology loaded")

        st.divider()
        with st.expander("How to use", expanded=True):
            st.markdown(
                "- **Keyword Search** — space-separated keywords, AND logic, case-insensitive, "
                "matches titles and keywords.\n"
                "- **Ontology Search** — maps your input to ontology concepts, expands "
                "subclasses / equivalent classes, then matches titles and keywords.\n"
                "- Example: `HBV_proteins` expands to `capsid_protein`, `envelope_protein`, "
                "`HBsAg`, `HBc`, …\n"
                "- To use another ontology, drop the file into `ontologies/` or upload it above."
            )


def main() -> None:
    st.set_page_config(
        page_title="HBV Literature Search",
        page_icon="🔍",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_state()
    render_sidebar_header()

    # ---------------------------- ① data & scope ----------------------------
    st.subheader("① Database & Search Scope")
    try:
        libs = build_libraries()
    except LibraryError as exc:
        st.error(f"Failed to load the literature data:\n\n{exc}")
        st.stop()
    except Exception as exc:                            # noqa: BLE001
        st.error(f"Unexpected error while loading data: {exc}")
        st.stop()

    register_ontologies()

    selected_libs = st.multiselect(
        "Target databases",
        options=list(libs.keys()),
        default=list(libs.keys()),
        key="target_libs",
    )
    if not selected_libs:
        st.warning("Please select at least one database.")

    # ---------------------------- ② search mode -----------------------------
    st.subheader("② Search Method")
    apply_pending_mode()
    mode = st.radio(
        "Search mode",
        options=[MODE_KEYWORD, MODE_ONTOLOGY],
        horizontal=True,
        key="search_mode",
        label_visibility="collapsed",
    )
    if st.session_state.get("_mode_seen") != mode:
        st.session_state["_mode_seen"] = mode
        st.session_state["last_result"] = None
        st.session_state["unmapped"] = None

    # 是否把文章正文/摘要（content 列）纳入关键词匹配
    include_content = st.checkbox(
        "Include article content (abstract) in matching",
        value=False,
        key="kw_include_content",
        disabled=(mode == MODE_ONTOLOGY),
        help=(
            "Applies to Keyword Search only. Off (default): match title + keywords. "
            "On: also scan the article content/abstract column — higher recall, "
            "but potentially more noisy hits."
        ),
    )

    query = ""
    trigger = False

    if mode == MODE_KEYWORD:
        fields = "title + keywords + abstract" if include_content else "title + keywords"
        query = st.text_input(
            f"Keywords (space-separated, AND logic, case-insensitive; matched against {fields})",
            placeholder="e.g.  HBV polymerase   or   乙肝 疫苗",
            key="kw_query",
        )
        trigger = st.button("Search", type="primary", key="kw_submit")
        if trigger and not query.strip():
            st.warning("Please enter at least one keyword before searching.")
            trigger = False
        if trigger and not selected_libs:
            st.warning("Please select at least one database.")
            trigger = False
        if trigger:
            run_keyword_search(libs, selected_libs, query, include_content=include_content)
            trigger = False
    else:
        # Create the uploader first (we need its value), then place it visually
        # below the dropdown through a pre-allocated container.
        slot = st.container()
        uploaded = st.file_uploader(
            "Import ontology (.owl / .ttl / .rdf / .xml / .n3)",
            type=["owl", "ttl", "rdf", "xml", "n3"],
            key="ont_upload",
            help="The uploaded file is added to the dropdown above and selected automatically.",
        )
        save_uploaded_ontology(uploaded)

        with slot:
            registry = st.session_state["ontologies"]
            if registry:
                names = list(registry.keys())
                active = st.session_state.get("active_ontology") or names[0]
                default_index = names.index(active) if active in names else 0
                choice = st.selectbox(
                    "Select ontology",
                    options=names,
                    index=default_index,
                    key="ont_choice",
                    format_func=lambda n: f"{n} (default)" if n == "HBPO" else n,
                )
                st.session_state["active_ontology"] = choice
            else:
                choice = None
                st.selectbox("Select ontology", options=["No ontology available"], index=0, disabled=True)
                st.warning(
                    "HBPO ontology not found. Put `HBPO.owl` in any of these locations:\n\n"
                    + "\n".join(f"- `{p}`" for p in HBPO_CANDIDATES[:2])
                    + "\n\nYou can also upload a `.owl` / `.ttl` file below."
                )

        query = st.text_input(
            "Concept name (mapped to an ontology concept, then semantically expanded)",
            placeholder="e.g.  HBV_proteins / envelope / HBsAg / HBV_structure_component",
            key="ont_query",
        )

        with st.expander("Advanced options", expanded=False):
            depth = st.slider("Expansion depth (subClassOf levels)", 1, 5, 3)
            include_parents = st.checkbox("Also include parent concepts", value=False)
            loose = st.checkbox(
                "Use concept stem for looser matching (e.g. envelope_of_HBV → envelope)",
                value=True,
            )

        trigger = st.button("Search", type="primary", key="ont_submit")
        if trigger and not query.strip():
            st.warning("Please enter a concept name before searching.")
            trigger = False
        if trigger and not selected_libs:
            st.warning("Please select at least one database.")
            trigger = False

        if trigger and not st.session_state["ontologies"]:
            st.error("No ontology is available. Please import one or switch to keyword search.")
            trigger = False

        if trigger:
            run_ontology_search(
                libs,
                selected_libs,
                query,
                st.session_state["ontologies"][st.session_state["active_ontology"]],
                depth,
                include_parents,
                loose,
            )
            trigger = False

        # Persistent hint when the input cannot be mapped to any concept
        unmapped = st.session_state.get("unmapped")
        if unmapped:
            st.warning(
                f"“{unmapped['query']}” could not be mapped to any concept of ontology "
                f"“{unmapped['ontology']}”. Try a concept name from the list below, "
                "or switch to keyword search."
            )
            with st.expander("Available concept examples", expanded=False):
                st.write(", ".join(f"`{s}`" for s in unmapped["samples"]))
            if st.button("Switch to keyword search", key="fallback_kw"):
                st.session_state["pending_mode"] = MODE_KEYWORD
                st.session_state["unmapped"] = None
                st.rerun()

    # ------------------------------- ③ results ------------------------------
    st.subheader("③ Results")
    opt1, opt2 = st.columns(2)
    show_source = opt1.checkbox("Show source database column", value=False, key="opt_show_lib")
    show_hit = opt2.checkbox(
        "Show matched concepts column (ontology search)", value=False, key="opt_show_hit"
    )
    render_results(show_source=show_source, show_hit=show_hit)

    render_sidebar()


if __name__ == "__main__":
    main()
