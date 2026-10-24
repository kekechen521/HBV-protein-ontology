# HBPO：乙型肝炎病毒蛋白本体

**HBPO（HBV Protein Ontology，乙型肝炎病毒蛋白本体）** 是一份面向乙型肝炎病毒（HBV）领域的轻量级 OWL 本体，用于统一描述 HBV 的病毒蛋白、结构组分与感染过程，并可作为语义检索与知识聚类的概念底座。

本仓库同时提供一个基于Streamlit的文献检索系统：既可用传统关键词检索，也可加载 HBPO 本体做概念级的语义检索。

---

## 作者与致谢

| 角色 | 姓名 |
|---|---|
| 本体创建 | **Chenheng** |
| 本体优化 | **Chenyirong** |

本体的类名、层级结构与术语体系由Chenheng建立；Chenyirong 在其基础上补充了规范标签（`rdfs:label`）、同义词（`oboInOwl:hasExactSynonym`，含中文译名）与中英文注释（`rdfs:comment`），并完善了检索侧的概念映射与语义扩展逻辑。

**联系方式：E-mail: chenheng_dr@163.com**

> 使用、引用或基于本本体开展研究时，请注明作者信息。

---

## 一、本体内容

HBPO 采用 OWL/XML 格式，共 **55 个类**，按三个顶层分支组织：

| 顶层分支 | 术语数 | 说明 |
|---|---|---|
| HBV infection process（感染过程） | 27 | 从病毒吸附、侵入、脱壳、cccDNA 形成、转录、翻译后加工、衣壳化、逆转录到成熟与释放的完整生命周期 |
| HBV protein（病毒蛋白） | 13 | 衣壳蛋白（HBc/HBcAg）、包膜蛋白（HBsAg/LHBs/MHBs/SHBs）、功能蛋白（P 蛋白/聚合酶、HBx）、分泌蛋白（HBe/HBeAg） |
| HBV structural component（结构组分） | 15 | 病毒基因组、pgRNA、衣壳、核心颗粒、包膜、Dane 颗粒、球形与管形亚病毒颗粒等 |

注解规模：

| 注解类型 | 数量 |
|---|---|
| `rdfs:label`（英文规范名） | 55 |
| `oboInOwl:hasExactSynonym`（同义词，含缩写、全称、中文译名） | 323（其中中文 163） |
| `rdfs:comment`（中英文释义） | 110（55 英文 + 55 中文） |

本体 IRI：`http://www.semanticweb.org/ontologies/2012/6/Ontology1342071125359.owl`

本体为 OWL/XML 格式，可用 Protégé 直接打开。

---

## 二、主要用途

### 1. 知识聚类

HBPO 提供了一套可直接用于概念级聚合的分类骨架：

- **按分支归类**：把文献、术语或实验结果归入「感染过程 / 病毒蛋白 / 结构组分」三大分支，形成主题簇。
- **按层次聚合**：沿 `rdfs:subClassOf` 自顶向下聚合，例如把涉及 `HBc`、`HBsAg`、`P_protein`、`HBx` 的文献统一归到父概念 `HBV protein` 之下，从而得到可解释的多级聚类结果。
- **同义词归一**：`HBcAg`、`核心抗原`、`core protein` 等不同表述在聚类前即被归并为同一概念，避免同一主题被拆散到多个簇。
- **跨语言统一**：中文文献与英文文献可映射到同一套英文概念节点，便于中英文语料合并分析。

### 2. 文献检索

本仓库的应用提供两种检索方式，均可在 A / B / C 三个逻辑库中按需选择范围：

- **关键词检索**：空格分隔多词按 AND 逻辑、不区分大小写，对标题与关键词做子串匹配；可选是否把文章摘要（content）一并纳入匹配。
- **本体检索**：把输入映射到 HBPO 概念，沿 `subClassOf` / `equivalentClass` / `sameAs` 做语义扩展，再用扩展出的概念集合匹配文献，并在结果中标注每条文献由哪个概念命中。

由于本体重在同义词与中文译名，本体检索对缩写（`HBc` 覆盖 `HBcAg`）与中文表述（`乙肝核心抗原`）的覆盖明显优于纯关键词匹配。

---

## 三、目录结构

```
检索系统/
├── app.py                              # 主程序（单文件 Streamlit 应用）
├── requirements.txt                    # 依赖清单
├── README.md                           # 本文档
├── HBPO.owl                            # 本体文件
├── ontologies/
│   ├── HBPO.owl                        # 页面默认加载的本体（与上者内容一致）
│   ├── _original/                      # 未加注解的原始本体备份
│   └── uploaded/                       # 页面历史上传的本体副本
```

---

## 四、快速开始

### 1. 环境要求

Python 3.10 及以上。

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

依赖说明：

| 包 | 用途 |
|---|---|
| `streamlit` | 页面框架 |
| `pandas` / `numpy` | 数据读取与匹配 |
| `openpyxl` | 读取 `.xlsx` |
| `owlready2` | 本体解析（主解析器） |
| `rdflib` | 通用 RDF 兜底（`.ttl` / `.rdf` 等） |
| `pyarrow` | 表格渲染；若本机 DLL 损坏可执行 `pip install --force-reinstall pyarrow` |

## 五、Acknowledgment
本研究得到了温州瓯粤海智计划创新人才项目（No.WZ2023YYCJH0169）、温州大学引进高层次人才科研启动基金（No.QD2021078）和温州大学挑战杯项目的支持。
This study was supported by Innovation Talent Project of Wenzhou Ou Yue Hai zhi Plan (No.WZ2023OYYCJH0169), the start foundation of Wenzhou University's scientific research for introducing high-level talents (No.QD2021078), and Challenge Cup Project of Wenzhou University.

## English summary

HBPO (HBV Protein Ontology) is a lightweight OWL ontology describing hepatitis B virus proteins, structural components and infection-process steps. It contains 55 classes in three top-level branches — infection process (27), viral proteins (13) and structural components (15)— annotated with English labels, 323 synonyms and 110 bilingual comments. The ontology can be used for concept-level knowledge clustering and for ontology-based literature retrieval. It was created by **Chenheng** and later refined by **Chenyirong**; contact: **chenheng_dr@163.com**.
