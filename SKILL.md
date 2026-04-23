---
name: contract-reviewer-wudi
display_name: 合同审核-WUDI
version: 2.9.1
description: |
  中文合同审核 Skill - 支持修订痕迹、批注气泡、清洁版生成与多轮迭代管理。

  ★ 修订痕迹（Track Changes）为本 Skill 的核心输出模式。
    每次审核应优先将具体修改建议写入文档修订痕迹；
    批注仅用于补充说明（结构性提示、法律背景说明等无法用修订表达的内容）。
    不接受"只有批注、没有修订"的输出方式。

  核心能力：
  1. 合同结构解析 - 提取条款、定义词、交叉引用
  2. 合同类型检测 - 自动识别41种合同类型（关键词匹配+置信度排序）
  3. 条款库 + 智能推荐引擎 - 内置22个专业条款模板，AI审核时自动匹配替代条款
  4. 交叉引用检查 - 检测孤立引用、缺失条款、自引用问题
  5. 修订痕迹生成 - 直接写入 Word Track Changes ★ 核心输出
  6. 批注气泡写入 - 直接写入 Word Comment（补充说明用）
  7. PDF/扫描件 OCR - 自动识别PDF类型，调用 Vision/Tesseract 转换
  8. 多轮迭代管理 - 版本跟踪、对比、回滚

  内置审核规则覆盖 41 种合同类型（350+ 关键条款，229+ 常见风险，82 个法律依据），
  涵盖《民法典》合同编主要类型、金融合规、PPP、电子商务等全场景。

  输出：修订批注版 docx + 清洁版 docx（仅 2 个文件，不输出 JSON 和报告）
license: GPL-3.0
author: WUDI
---

## 更新日志

### v2.9.1 | 2026-04-20
- **配置集成修复**：
  - 修复 config.py 未被任何脚本导入的 bug
  - apply_changes.py 现在正确导入 config 模块
  - 替换所有硬编码字体为 `get_system_font()`
  - 替换所有硬编码字号为 `DEFAULT_FONT_SIZE`
- **文件清理**：
  - 删除备份文件 `apply_changes_v2_backup.py`
- **功能集成**：
  - generate_comparison.py 现已集成到 pipeline
  - 自动生成修订对比 Markdown 表
- **测试修复**：
  - 修正测试导入路径问题

### v2.9.0 | 2026-04-20
- **代码质量改进**：
  - 添加异常处理补全（apply_changes.py 第 311 行）
  - 添加进度提示（每 5 条修订显示进度）
  - 改进 LibreOffice 缺失时的错误消息（提供 3 种解决方案）
- **新增模块**：
  - `config.py`：共享配置模块（集中管理字体、阈值、配色）
  - `generate_comparison.py`：对比视图生成器（Markdown 格式）
- **单元测试**：
  - 新增 `tests/test_contract_reviewer.py`
  - 覆盖：JSON schema 验证、格式检测、段落匹配
- **文档改进**：
  - 统一命名规范（修订痕迹版 vs 修订批注版）
  - 补充配置项说明

### v2.8.2 | 2026-04-20
- **代码缺陷修复**：
  - 修复 `apply_changes.py` 第 447-449 行重复代码（`return i` 后有不可达代码）
  - 该 bug 不会影响运行（Python 会忽略 return 后的代码），但会干扰静态分析
- **JSON Schema 验证**：
  - 新增 `validate_schema()` 函数，检查 revisions[].revised_text 是否为空
  - 检查 comments[].comment 是否为空
  - 空值字段直接报错而非静默失败
- **文档一致化**：
  - 统一"目录结构"章节命名（修订痕迹版 vs 修订批注版）
  - 所有输出文件名使用 `{合同名}-修订痕迹版.docx` 格式

### v2.8.1 | 2026-04-20
- **JSON 生成规范（防止解析失败）**：
  - 明确要求使用 `json.dumps()` 而非手动拼接 JSON 字符串
  - 字符串值内的引号使用单引号或全角引号，避免未转义的 ASCII 双引号
  - 生成后立即预检验，失败则重新生成而非修复损坏文件
- **输出文件规范**：
  - 严格限制为 2 个 docx 文件（修订痕迹版 + 清洁版）
  - 不输出 JSON 和 Markdown 报告（除非用户明确要求）
  - 移除 changes.json 和审核报告作为默认交付物
- **问题来源**：第二轮审核时 JSON 解析失败，手动拼接导致内嵌引号未转义；输出文件过多超出用户预期

### v2.8.0 | 2026-04-20
- **输出文件命名优化**：
  - 修订痕迹版：`{合同名}-修订痕迹版.docx`（原"审查版"改为更准确的"修订痕迹版"）
  - 清洁版：`{合同名}-修订痕迹版-清洁版.docx`
  - 审核报告：`{合同名}-审核报告.md`（新增自动生成）
- **清洁版生成失败优雅降级**：
  - 主方案失败时自动尝试 LibreOffice 命令行备用方案
  - 最终失败时提供手动操作指引（Word 中接受所有修订）
  - 不再因清洁版失败而中断整个流程
- **预检验子串匹配自动修复**：
  - 检测到子串匹配（非完整匹配）时，自动补全为完整段落原文
  - 减少人工修正成本
- **自动生成审核报告**：
  - Pipeline 完成后自动生成 Markdown 格式审核报告
  - 包含：风险统计、高风险问题清单、待填项提示、使用说明
  - 可直接作为交付物或内部存档

### v2.7.9 | 2026-04-20
- **修订批注自动配套（核心修复）**：
  - **问题**：v2.7.7 要求"每条修订都应有配套批注"，但 AI 需要手动在 `comments` 数组重复写一遍
  - **解决**：脚本现在自动处理 `revisions[].comment` 字段，修订成功后自动生成配套批注
  - **AI 使用方式**：只需在 `revisions` 中写一次 `comment` 字段，脚本自动完成：
    ```
    revisions[].comment → 自动生成批注（修订痕迹 + 批注气泡）
    ```
  - **JSON 格式简化**：
    ```json
    {
      "revisions": [
        {
          "paragraph_index": 42,
          "original_text": "原文",
          "revised_text": "修订后文本",
          "comment": "【修订批注-高风险】问题：... 依据：... 理由：...",
          "severity": "高风险",
          "highlight_text": "原文关键词（可选）"
        }
      ],
      "comments": []  // 仅用于无需修订的风险提示
    }
    ```
  - **向后兼容**：原有的 `comments` 数组仍然有效，用于无需修订的纯风险提示
- **移除 TXT 输出**：
  - 审查报告直接在 docx 中通过批注呈现，无需单独的 txt 文件
  - 输出简化为：修订批注版 docx + 清洁版 docx + changes.json

### v2.7.7 | 2026-04-18 晚
- **修订+批注配套原则（重要修复）**：
  - **问题发现**：甲方审查版有13处修订痕迹，但只有2条批注，大部分修订没有说明修改原因
  - **根本原因**：只应用了修订痕迹，忘记同时写入批注说明修改原因
  - **正确做法**：每条重要修订都应有对应批注，修订痕迹显示"改了什么"，批注说明"为什么改"
  - **批注写入两步**：①comments.xml 写入批注内容 ②document.xml 添加 commentRangeStart/End/Reference 引用
  - **批注ID匹配**：comments.xml 中的 ID 必须与 document.xml 中的引用 ID 完全一致
- **多轮审核批注保留**：
  - **问题发现**：第二轮审核时遗漏了甲方第一轮的批注
  - **正确做法**：第二轮审核时必须保留第一轮所有批注，新批注ID从上一轮最大ID+1开始
  - **批注ID分配**：甲方用ID 0-12，乙方用ID 20-31（中间留空便于插入）
- **定稿检查清单补充**：
  - 新增检查项：批注数量是否与修订数量匹配
  - 新增检查项：批注ID是否正确关联到文本位置
  - 新增检查项：多轮审核是否保留了前一轮的所有批注

### v2.7.6 | 2026-04-18
- **修订痕迹（Track Changes）升级为核心输出模式**：
  - SKILL.md 明确要求：每次审核必须包含修订痕迹（revisions），批注（comments）仅作补充
  - 字段规范已固定：`original_text`（原文）、`revised_text`（修订文）、`comment`（修订说明）
  - 经验：修订内容应尽量具体完整，填入合理默认值（天数、比例等），不要只说"建议修改"而不给具体文字
  - 批注的适用场景：①条款编号断裂等结构性提示；②法律背景说明；③无法精准定位原文的全局问题
  - 本次实战验证：14条修订 + 1条批注，成功写入 Track Changes ✅

### v2.7.5 | 2026-04-17
- **修订模式不显示的完整修复**（本轮两轮迭代）：
  - **第一轮**：修正 `w:delText` 嵌套位置（必须和 `w:rPr` 是兄弟节点，不是子元素）
  - **第二轮**：移除 `settings.xml` 中的 `<w:trackRevisions/>` 和 `<w:revisionView>`
    - Word 2016 原生参考文档的 `settings.xml` 里**完全没有**这两个元素
    - 加上反而可能导致 Word 渲染异常
  - compat mode 统一改为 14（与 Word 2016 原生一致）
  - 最终 XML 结构与 Word 2016 原生输出零差异：`w:del > w:r(w:rsidRPr) > w:rPr + w:delText`

### v2.7.4 | 2026-04-17
- 早期修订兼容尝试（后被 v2.7.5 替代）

### v2.7.2
- **预检验引擎（Preflight Check）**（本次修复核心）：
  - 新增 `scripts/preflight_check.py`（~350行，零外部依赖）
  - **5步自动诊断**：JSON 格式检测 → docx 段落读取 → 索引越界校验 → 原文匹配测试 → --fix 自动修正
  - **自动转换**：扁平 `{changes: [...]}` 格式 → 标准 `{revisions, comments}` 格式，无需手动重写
  - **自动修正**：全文扫描找最接近段落，修正 `paragraph_index` 并输出 `_fixed.json`
  - **内置 pipeline**：步骤 3.5（应用修订前）自动运行，发现问题退出并提示，修复后重跑即可
  - `apply_changes.py` 修复：`_replace_run_across_multiple` 中移除 XML 元素前去重，消除 `ValueError: list.remove(x)` 崩溃
  - `apply_changes.py` 修复：`comment['comment']` 同时支持 `comment` 和 `text` 两个字段名
  - **JSON 字段规范（必须严格遵守）**：
    - 修订：必须用 `revised_text`（非 `replacement_text`）
    - 批注：必须用 `comment`（非 `text`）；可选 `highlight_text` 指定高亮范围
  - **Python 版本兼容**：`full_pipeline.py` 类型注解改用 `Optional[Path]`（兼容 Python 3.9）
  - 经验教训：原文 `original_text` 必须从 document.xml 原始 `<w:t>` 提取，不能用解析器的中间输出

### v2.7.1
- **条款推荐引擎（Clause Recommender）**：
  - 新增 `scripts/clause_recommender.py`（~350行，零外部依赖）
  - **双因子评分算法**：合同类型匹配分（+5 精确命中 / +3 ID包含 / +1 类别包含）+ 风险关键词匹配分（+2 完全匹配 / +0.5 部分匹配）
  - **内置风险→条款映射**：17 种常见风险关键词自动映射到对应条款类别
  - **CLI 接口**：`--type`、`--risks`、`--keywords`、`--top-k`、`--all`、`--output-json`
  - 测试"违约金偏高+收款账户空白"：违约金条款和资金监管条款均获 **5星** 推荐 ✅
- **审核流程嵌入**：
  - `references/review-playbook.md` 新增"第五章：自动条款推荐"（含风险→条款映射表、4条使用规范）
  - AI 审核时，每识别风险自动调用推荐引擎，`suggestion` 字段写入推荐文本，`clause_source` 记录来源
  - 章节编号顺延（原来5-9章 → 6-10章）
- **SKILL.md 引擎表**：新增"条款推荐引擎"一行，完整记录7个引擎

### v2.7.0
- **合同类型自动检测引擎（P0）**：
  - `contract_parser.py` 新增 `detect_contract_type()` 函数，基于关键词匹配自动识别合同类型
  - 支持 41 种合同类型自动检测，置信度评分算法（归一化+命中密度双因子）
  - 新增 CLI 参数：`--detect-type [--top-k N]`，返回前 N 名候选类型
  - 内置配置缓存（避免重复加载 JSON 文件）
  - 房屋买卖合同测试置信度 1.771，大幅领先第二名（0.460）
- **条款库（Clause Library）（P1）**：
  - 新增 `references/clause-library/` 目录，含 22 个专业条款模板
  - 覆盖 9 大类别：付款（5）、违约金（3）、争议解决（4）、不可抗力（1）、知识产权（3）、交付验收（2）、保密（2）、解除终止（2）、综合（3）
  - 每条含 `clause_text`（可直接粘贴）、`position_tips`（甲方/乙方立场建议）、`legal_reference`
  - 适用合同类型标签，可按需选用
- **交叉引用一致性检查（P1）**：
  - 新增 `scripts/check_cross_refs.py`，零外部依赖
  - 检测四类问题：孤立引用（指向不存在的条款）、编号缺失（条款编号跳跃）、自引用（条款引用自身）、附件引用检查
  - 中文数字与阿拉伯数字互转，支持"第一条""第1条"混排合同
- **PDF/扫描件自动 OCR 集成（P1）**：
  - `full_pipeline.py` 新增步骤 0：自动检测 PDF 类型（文本型 vs 扫描件）
  - 方案A（macOS）：调用 Vision 框架零依赖 OCR，支持中文+英文，无需安装 tesseract
  - 方案B（跨平台）：回退 pytesseract + PyMuPDF（需 pip install）
  - OCR 生成的 docx 自动进入正常审查流程
- **JSON 修复**：批量修复 house-sale.json 和 ppp.json 的转义问题，所有 41 个配置文件全部通过 JSON 校验

### v2.6.0
- **合同类型审核要点库重大扩展**（从 21 种扩展至 41 种）：
  - 新增 20 种专业合同类型：合资、信托、基金、金融衍生品、进出口、数据合规、电信、娱乐、医疗、电商、SaaS、PPP、保证、抵押、质押、居间、借用、和解、仲裁、股权收购
  - 新增配置文件：`references/contract-types/` 目录下 41 个 JSON 文件（41 种类型 + 1 模板）
  - 新增统计：350 项关键条款、229 项常见风险、82 个法律依据引用
  - 新增分类体系：按房产/商事/公司/工程/知产/劳动/金融/物流/担保/民事/争议解决/国际贸易/合规/服务 14 个大类分类
  - 新增功能：自动识别合同类型 → 加载对应审核规则 → 执行针对性检查 → 生成立场建议
  - 覆盖范围：完整覆盖《民法典》合同编主要类型 + 主要商业/金融/监管类合同
  - 新增 README：`references/contract-types/README.md` 含完整清单、使用说明、分类体系和版本历史

### v2.5.0
- **批注 Word 兼容性修复**（经 WPS 生成文件实测）：
  - `apply_comment()` 重写批注范围结构：`commentRangeStart` 紧跟 `</w:pPr>` 之后，`commentRangeEnd` 在原文 runs 之后，`commentReference` run 在最后——完全符合 python-docx 实测 OOXML 标准
  - `add_to_comments_xml()` 新增 `w14:commentId` 属性（Word 2010 识别批注 ID 关键）
  - 新增 `_inject_w14_namespace()`：在 document.xml 根标签注入 `xmlns:w14` + `mc:Ignorable="w14"`（lxml 无法直接输出 `w14:` 前缀，改为序列化后文本注入）
  - 新增 `_generate_comments_extended()`：生成 `commentsExtended.xml`（Word 2013+ 批注扩展），并同步更新 rels 和 Content_Types
  - 新增 `_fix_comments_content_type()`：修正 `PartName="/comments.xml"` → `"/word/comments.xml"`（OPC 规范路径）
  - 脚本行数：996 行（v2.4.0: 833 行）

### v2.4.0
- apply_changes.py 重写（27KB）：段落精确匹配（语义搜索 → 精确匹配 → 降级批注）、原生批注样式（风险等级颜色）、修订失败自动降级为批注、新增 parse_tracked_changes() 读取对方修订痕迹
- round_analyzer.py 升级：捆绑式谈判识别（bundled）、parse_bundled_stances() 条款编号提取、立场分类增强（拒绝模式+精准接受）
- full_pipeline.py：新增修订痕迹读取报告步骤

### v2.3.1
- 明确：多轮迭代（round_analyzer）触发条件——需同时满足"对方批注版"+"告知己方立场"

### v2.3.0
- 新增：`scripts/round_analyzer.py` 多轮迭代自动分析引擎（批注立场分类+法律反驳库+回应生成）
- 改进：批注字体改为宋体五号（10.5pt）
- 移除：`references/legal_basis_db.md` 和 `legal-basis.md`（法律依据改为联网实时检索）

### v2.2.0
- 修复：立场分类 bug（拒绝优先接受、部分接受优先接受）
- 改进：批注格式含【风险等级】【修改原因】【法律依据】【修改建议】四段式

### v2.1.0
- 重大改进：批注内容增强，增加 reason 和 legal_basis 字段
- 新增：法律依据参考库 legal_basis_db.md

### v2.0.0
- 完整重写，GPL-3.0 许可协议
---

# 中文合同审核 Skill

## 概述

本 Skill 用于处理中文合同的审核、修订和批注任务。通过内置的 Python 引擎，实现从文档解析到输出生成的完整自动化流程。

---

## 核心引擎

| 引擎 | 脚本 | 功能 |
|------|------|------|
| **结构解析** | `contract_parser.py` | 解析 docx，提取条款结构、定义词、交叉引用 |
| **合同类型检测** | `contract_parser.py` | 自动识别合同类型（41种），置信度排序 |
| **条款推荐引擎** | `clause_recommender.py` | 根据风险和合同类型，智能推荐22条替代条款 |
| **交叉引用检查** | `check_cross_refs.py` | 检查条款编号引用一致性，报告孤立引用/缺失条款 |
| **预检验引擎** | `preflight_check.py` | 运行前自动诊断 JSON 格式、段落索引、文本匹配，失败则退出并生成修正版 |
| **修订写入** | `apply_changes.py` | 将修订建议写入 Word 修订模式（含元素去重修复） |
| **批注写入** | `apply_changes.py` | 将风险提示写入 Word 批注（含 Word/WPS 兼容修复） |
| **清洁版生成** | `generate_clean.py` | 接受修订、移除批注，生成清洁版 |
| **迭代管理** | `iteration_manager.py` | 管理多轮审核版本，支持对比和回滚 |
| **PDF/扫描件** | `full_pipeline.py` | 自动检测 PDF 类型，调用 Vision/Tesseract OCR 转换 |

---

## 使用场景

当你需要以下操作时，使用本 Skill：

- 审核中文合同并识别风险条款（支持 41 种合同类型自动识别）
- 生成带修订痕迹的 Word 文档
- 生成带批注气泡的 Word 文档
- 生成清洁版（接受所有修订）
- 管理多轮审核版本

---

## 合同类型审核要点库（v2.6.0 新增）

内置 41 种合同类型的专业审核规则，位于 `references/contract-types/` 目录。

### 支持的合同类型（41 种）

| 分类 | 类型 | 分类 | 类型 |
|------|------|------|------|
| 房产类 | 房屋买卖合同、租赁合同 | 商事类 | 买卖、经销、特许经营、保理、融资租赁、居间、电商、SaaS、PPP |
| 公司类 | 投资、合伙、股权收购、合资 | 工程类 | 建设工程 |
| 知产类 | 技术服务、IP许可、IP转让、保密协议 | 劳动类 | 劳动合同 |
| 金融类 | 借款、保险、信托、基金、金融衍生品 | 物流类 | 运输、仓储 |
| 担保类 | 保证、抵押、质押 | 民事类 | 委托、赠与、借用、和解 |
| 争议解决 | 仲裁 | 国际贸易 | 进出口 |
| 合规类 | 数据合规 | 服务业 | 电信、医疗、娱乐 |

**总计：350 项关键条款 · 229 项常见风险 · 82 个法律依据引用**

### 工作原理

1. **自动识别**：根据合同文本中的关键词（keywords_for_detection）自动识别合同类型
2. **加载规则**：加载对应合同类型的 JSON 配置文件
3. **执行检查**：按 key_clauses（关键条款）和 common_risks（常见风险）逐一检查
4. **立场建议**：根据 position_guidance（立场指导）为不同立场生成建议

### 配置结构

每个配置文件包含：
- `applicable_laws`：适用法律列表
- `keywords_for_detection`：合同类型识别关键词
- `key_clauses`：关键条款检查要点（含致命/重大/一般风险等级）
- `common_risks`：常见风险清单（含法律依据和修改建议）
- `position_guidance`：甲方/乙方各自的关键关注点和应避免问题

### 使用方式

审核时，AI 会自动：
1. 分析合同文本，匹配关键词，识别合同类型
2. 加载对应审核规则（可指定类型以提高准确度）
3. 按审核规则执行检查
4. 生成针对性的风险提示和修改建议

**指定合同类型可提高审核准确度**：
> "帮我审核这份PPP合同，站在社会资本方视角"

### 条款库（Clause Library）

内置专业条款模板库，位于 `references/clause-library/`，共 22 个模板，覆盖 9 大类别：

| 类别 | 数量 | 说明 |
|------|------|------|
| 付款条款 | 5 | 预付、里程碑、月结、资金监管（房屋）、工程资金监管 |
| 违约金 | 3 | 日万分之五、固定20%、劳动合同滞纳金 |
| 争议解决 | 4 | 原告所在地法院、指定法院、CIETAC仲裁、国内仲裁 |
| 不可抗力 | 1 | 标准不可抗力条款（含通知、后果、超期解除） |
| 知识产权 | 3 | 工作成果归属（甲方所有）、共有知识产权、反向许可 |
| 交付/验收 | 2 | 7天验收+逾期视为合格、第三方验收 |
| 保密条款 | 2 | 3年期限、永久保密（限核心商业秘密） |
| 解除/终止 | 2 | 随时解除（30天通知）、违约补救期（15天） |
| 综合类 | 3 | 书面变更、准据法、完整协议（Entire Agreement） |

每条条款含：
- `clause_text`：**可直接粘贴**的条款正文文本
- `position_tips`：甲方/乙方版本差异与谈判建议
- `legal_reference`：法律依据（《民法典》/《劳动合同法》等）
- `applicable_types`：适用的合同类型标签

**使用方式**：审核时，根据风险提示，从条款库选取对应替代文本，写入 changes.json：
```json
{
  "paragraph_index": 42,
  "original_text": "违约金为合同总金额的 30%",
  "revised_text": "任一方逾期履行合同义务的，每逾期一日，应按合同总价的万分之五向守约方支付违约金……",
  "reason": "违约金约定30%可能被认定为过高，建议按民法典585条调整为日万分之五",
  "clause_source": "clause-library/penalty/daily-5-per-10k.json"
}
```

---

## 审核框架

### ⚠️ 修订优先工作流（重要！）

每次合同审核必须遵循以下工作流：

```
1. 识别风险 → 判断能否精准定位原文
   ↓ 能定位（绝大多数情况）
2. 写修订条目（revisions）：
   {
     "paragraph_index": X,
     "original_text": "原文",
     "revised_text": "完整修订后文本",
     "comment": "修订原因"
   }
   ↓ 仅在以下情况才用批注
3. 写批注（comments）：
   - 条款编号断裂等结构性提示
   - 法律背景说明
   - 无法精准定位原文的全局问题
   ↓
4. 两者并存时：revisions 为主，comments 为辅
```

### ⚠️ 修订+批注配套原则（v2.7.7 新增）

**核心原则**：修订痕迹显示"改了什么"，批注说明"为什么改"

```
正确做法：
┌─────────────────────────────────────────────────────────┐
│ 修订痕迹：删除 "拾倍赔偿" → 插入 "30%违约金"            │
│ 批注内容：【修改原因】原"拾倍"比率畸高，依法可调减      │
└─────────────────────────────────────────────────────────┘

错误做法：
┌─────────────────────────────────────────────────────────┐
│ 修订痕迹：删除 "拾倍赔偿" → 插入 "30%违约金"            │
│ 批注内容：（无）← ❌ 对方不知道为什么要这样改           │
└─────────────────────────────────────────────────────────┘
```

**批注写入技术要点**：
1. comments.xml：写入批注内容（`<w:comment w:id="X">...</w:comment>`）
2. document.xml：添加三个引用标签：
   - `<w:commentRangeStart w:id="X"/>` - 批注开始位置
   - `<w:commentRangeEnd w:id="X"/>` - 批注结束位置
   - `<w:commentReference w:id="X"/>` - 批注引用
3. **ID必须完全一致**：comments.xml 中的 ID 与 document.xml 中的引用 ID 必须相同

**多轮审核批注ID分配建议**：
- 第一轮（甲方）：ID 0-9
- 第二轮（乙方）：ID 20-29（中间留空便于插入）
- 第三轮：ID 40-49
- 以此类推

**反面案例（已废弃）：**
- 只写批注气泡，不写修订痕迹 → ❌ 错误
- `revised_text` 只写"建议修改为合理比例"而不给具体文本 → ❌ 错误
- 留空项不填默认值（如"提前 天通知"空着） → ❌ 错误

**正面案例（正确做法）：**
- "提前 天通知" → `revised_text: "提前30天以书面形式通知乙方"` ✅
- "按第四条约定分成"（第四条不存在）→ `revised_text: "甲方占50%，乙方占50%"` ✅
- "拾倍赔偿" → `revised_text: "支付相当于瞒收学费总额30%的违约金"` ✅

---

### 四要素批注模板（v2.7.8 新增）

**核心改进**：从自由文本改为结构化四要素，每条批注必须包含：

```
【修订批注-风险等级】【修订方式：精准修改/完全重写】

问题：[原条款有什么问题]
修改：[具体怎么改]
依据：[法律/国标名称 + 第X条 + 内容节选]
理由：[为什么这样改，解决什么风险]
```

**示例**（质保金条款批注）：
```
【修订批注-高风险】【修订方式：精准修改（尾部追加）】

问题：原条款要求验收后付清全部款项，甲方在质保期内无任何资金保障。
      如乙方不履行质保义务，甲方需另行起诉追偿，成本高、周期长。

修改：验收合格后付至95%，扣留5%作为质量保证金，质保期满后10个工作日内无息付清。

依据：《民法典》第782条：承揽人交付的工作成果不符合质量要求的，定作人可以合理选择
      请求承揽人承担修理、重作、减少报酬、赔偿损失等违约责任；
      《建设工程质量保证金管理办法》第7条：保证金总预留比例不得高于工程价款结算总额的3%
      （家装承揽合同可参照适用，5%比例符合行业惯例且法院通常支持）。

理由：5%质保金是落实第782条的具体担保方式，确保乙方有动力履行质保义务；
      10个工作日为合理付款期限，避免乙方无限期拖延。
```

---

### 修订方式标注（v2.7.8 新增）

批注开头必须明确标注修订方式，帮助用户理解修改幅度：

| 标注 | 含义 | 适用场景 |
|------|------|----------|
| 【修订方式：精准修改】 | 原文框架可用，只改局部 | 改数字（30%→20%）、填空、插入句子 |
| 【修订方式：完全重写】 | 原文不可用，必须整段替换 | 缺失违约责任条款，需新增完整条款 |
| （尾部追加） | 在原文末尾添加内容 | 增加验收流程、补充协议条款 |
| （插入国标引用） | 在原文中间插入引用 | "符合国家标准"→"符合GB/T 8478" |
| （填空） | 填入空白处的默认值 | "质保[空白]年"→"质保5年" |

---

### 数值计算依据规范（v2.7.8 新增）

**问题**：数值类修改（5%、5日、万分之五）没有说明来源，缺乏说服力

**规范**：以下常见数值必须说明计算/惯例依据

| 数值 | 依据说明 |
|------|----------|
| **质保金5%** | 参照《建设工程质量保证金管理办法》上限3%，家装行业惯例3-5% |
| **违约金日万分之五** | 年化18.25%，参照LPR四倍（约15.4%），司法实践通常支持（《九民纪要》第50条） |
| **验收期限5个工作日** | 参照《建设工程施工合同（示范文本）》GF-2017-0201第32条"28天内组织验收"，家装规模小取5日合理 |
| **质保期5年/2年/5年** | 玻璃5年/配件2年/铝合金5年，参照GB/T 8478-2020、JGJ 113-2018行业通行标准 |
| **定金20%** | 《民法典》第586条强制性上限，超过部分不产生定金效力 |
| **解除权15日** | 《民法典》第563条"合理期限"，家装工程通常取15日 |
| **维修响应48小时** | 行业惯例，确保质量问题及时处理 |

**写作格式**：
```
日万分之五违约金：年化18.25%，略高于LPR四倍（约15.4%），
但司法实践中法院通常支持（《九民纪要》第50条）。
```

---

### 审核维度

按优先级从高到低审查：

| 维度 | 说明 | 示例问题 |
|------|------|----------|
| **准确性** | 表意是否唯一确定 | "合理期限" 缺乏量化标准 |
| **逻辑性** | 条款间是否协调 | 违约金与解约权冲突 |
| **规范性** | 是否符合行业惯例 | 付款条件缺少验收标准 |
| **简洁性** | 是否存在冗余表达 | 同一定义重复出现 |
| **流畅性** | 行文是否通顺 | 长句嵌套过多 |
| **工整性** | 格式是否统一 | 编号层级不规范 |

### 风险分级

| 等级 | 判定标准 | 处理方式 |
|------|----------|----------|
| **高风险** | 可能导致重大损失或法律无效 | 必须修改 |
| **中风险** | 可能引发争议或不利解释 | 建议修改 |
| **低风险** | 表述不当但不影响实质权利 | 可选修改 |

### 问题分类

- 法律风险
- 商业风险
- 表述问题
- 结构问题
- 格式问题

---

## 工作流程

### 首次审核

```
步骤 0: 检测文件类型（PDF → 自动 OCR；DOC → 自动转 DOCX）
        → python3 scripts/full_pipeline.py 合同.pdf changes.json
        （支持：.doc / .docx / PDF文本 / PDF扫描件）

步骤 1: 解析合同结构
        → python3 scripts/contract_parser.py 合同.docx 结构.json --detect-type
        （--detect-type 可自动识别合同类型，返回前3名候选+置信度）

步骤 2: AI 审核
        → 读取结构.json，识别风险条款
        → 生成 changes.json（修订建议 + 批注建议）
        → 可参考 clause-library/ 目录选取替代条款文本

步骤 3: 交叉引用一致性检查（可选，推荐）
        → python3 scripts/check_cross_refs.py 合同.docx 交叉引用报告.json
        （检测孤立引用、缺失条款、自引用、附件引用问题）

步骤 4: 写入修订和批注
        → python3 scripts/full_pipeline.py 合同.docx changes.json
        （pipeline 会在步骤 3.5 自动运行预检验，发现问题退出并提示修复）

        单独使用预检验（推荐先跑）：
        → python3 scripts/preflight_check.py 合同.docx changes.json --fix
        （5步诊断 + 自动修正 + 生成 _fixed.json）

步骤 5: 生成清洁版
        → python3 scripts/generate_clean.py 修订批注版.docx 清洁版.docx

步骤 6: 归档本轮
        → python3 scripts/iteration_manager.py init
        → python3 scripts/iteration_manager.py new-round
```

### 后续轮次

```
步骤 1: 查看历史
        → python3 scripts/iteration_manager.py status

步骤 2: 对比版本
        → python3 scripts/iteration_manager.py compare 1 2

步骤 3: AI 审核（基于上一轮清洁版）

步骤 4-6: 同首次审核
```

---

## 输入要求

### 必需信息

| 信息 | 说明 | 示例 |
|------|------|------|
| **合同文件** | 待审核的 docx 文件 | 采购合同.docx |
| **审核立场** | 偏甲方、偏乙方或中立 | 偏甲方 |

### 可选信息

| 信息 | 说明 |
|------|------|
| 合同类型 | 保密协议、采购合同、服务协议等 |
| 参考模板 | 优秀合同模板（如有） |
| 项目背景 | 交易背景信息 |
| 审核轮次 | 第几轮审核 |

---

## 输出规范

### 默认输出（v2.8.1 更新）

**⚠️ 输出文件数量：严格限制为 2 个 docx 文件**

| 文件 | 说明 | 必须 |
|------|------|------|
| `合同名-修订痕迹版.docx` | 含 Track Changes 和 Comments（可直接在 Word 中审阅） | ✅ 必须 |
| `合同名-修订痕迹版-清洁版.docx` | 接受所有修订后的版本（如生成失败会提供手动指引） | ✅ 必须 |
| ~~`合同名-审核报告.md`~~ | Markdown 格式审核报告 | ❌ 不输出（除非用户明确要求） |
| ~~`合同名-changes.json`~~ | 结构化变更记录 | ❌ 不输出（仅供内部调试） |

**输出原则：**
- 每轮审核只产生 2 个可见文件给用户
- JSON 文件仅在调试时使用，不作为交付物
- 审核报告如用户需要，再单独生成

### 目录结构

```
合同名-Output/
├── manifest.json              # 迭代管理清单（内部）
├── round-1/                   # 第一轮
│   ├── 合同名-修订痕迹版.docx    # ← 用户交付物
│   ├── 合同名-修订痕迹版-清洁版.docx  # ← 用户交付物
│   └── changes.json           # 内部调试用（不交付）
├── round-2/                   # 第二轮
│   ├── 合同名-修订痕迹版-乙方回复版.docx
│   └── ...
└── 合同名-结构解析.json        # 合同结构（全局，内部调试用）
```

---

## 审查报告格式

```
================================================================================
                          合同审查报告
================================================================================

一、合同概况
  - 合同名称：[合同名]
  - 合同类型：[类型]
  - 审核立场：[立场]
  - 审核轮次：第 N 轮
  - 审核日期：YYYY-MM-DD

二、风险问题清单

  【高风险】
  1. 第 X 条：[问题描述]
     风险说明：[为什么有风险]
     修改建议：[具体修改方案]

  【中风险】
  2. ...

  【低风险】
  3. ...

三、已修改内容
  - 第 X 条：原 "[...]" → 改为 "[...]"
    修改理由：[...]

四、格式问题
  - 定义词：发现 X 处不统一
  - 编号：第 X 条编号不连续
  - 标点：发现 X 处中英文标点混用

五、需确认事项
  1. [需要客户确认的问题]

六、统计数据
  - 本轮修订：X 处
  - 本轮批注：X 处
  - 高风险：X 处 | 中风险：X 处 | 低风险：X 处

================================================================================
                            报告结束
================================================================================
```

---

## 引擎使用

### 结构解析 + 合同类型检测

```bash
# 仅解析结构
python3 scripts/contract_parser.py 合同.docx 输出.json

# 解析结构 + 自动检测合同类型（返回前3名候选）
python3 scripts/contract_parser.py 合同.docx 输出.json --detect-type

# 解析结构 + 返回前5名候选类型
python3 scripts/contract_parser.py 合同.docx 输出.json --detect-type --top-k 5

# 输出包含：
# - 段落列表（索引、文本、样式）
# - 条款编号映射
# - 定义词列表
# - 交叉引用列表
# - 统计信息
# - _contract_type_detection（合同类型检测结果，含置信度和命中关键词）
```

### 交叉引用一致性检查

```bash
# 检查孤立引用、缺失条款、自引用
python3 scripts/check_cross_refs.py 合同.docx 报告.json

# 直接在终端查看摘要（无需输出文件）
python3 scripts/check_cross_refs.py 合同.docx
```

### 条款推荐引擎（自动嵌入审核流程）

```bash
# 根据合同类型 + 风险关键词推荐替代条款
python3 scripts/clause_recommender.py \
  --type house-sale \
  --risks "违约金偏高" "收款账户空白" \
  --top-k 2

# 显示所有类别（预览）
python3 scripts/clause_recommender.py --all

# 指定关键词（无需合同类型）
python3 scripts/clause_recommender.py \
  --keywords "定金" "过户" "违约金" \
  --top-k 3

# 输出 JSON（供 AI 写入 changes.json）
python3 scripts/clause_recommender.py \
  --type house-sale \
  --risks "违约金不对等" \
  -o /tmp/clause-recs.json
```

**AI 审核时的自动工作流（无需人工干预）：**

```
识别风险关键词（如"违约金偏高"）
    → 调用 clause_recommender.py --risks "违约金偏高" --type <合同类型>
    → 获取匹配得分最高的替代条款文本
    → 在 changes.json 的 suggestion 字段写入推荐文本
    → clause_source 字段记录来源：clause-library/penalty/daily-5-per-10k.json
```

**内置风险 → 条款映射规则（自动匹配）：**

| 风险关键词 | 推荐条款类别 | 首选文件 |
|-----------|------------|---------|
| 违约金 / 违约金偏高 | penalty | daily-5-per-10k.json |
| 收款账户空白 | payment | escrow.json |
| 知识产权归属 | ip | work-for-hire.json |
| 无不可抗力 | force-majeure | standard.json |
| 保密期限缺失 | confidentiality | nda-3years.json |
| 解除权不对等 | termination | breach-cure-15days.json |
| 无争议解决 | dispute | court-specified.json |
| 变更无书面 | general | amendment-written.json |

**输出字段说明**：
```json
{
  "statistics": {
    "total_clauses": 30,
    "dangling_refs": 2,    // 孤立引用（⚠️）
    "missing_clauses": 1,   // 缺失条款（⚠️）
    "self_refs": 0,
    "appendix_issues": 0
  },
  "issues": [
    {
      "type": "dangling_reference",
      "ref_num": 5,
      "ref_text": "第5条",
      "from_paragraph_index": 18,
      "from_text_snippet": "……详见第5条约定……",
      "severity": "中风险",
      "suggestion": "合同中引用了第5条，但该条款不存在……"
    }
  ]
}
```

### 预检验（必跑）

**每次运行 pipeline 前强烈推荐先跑预检验**——它会在写入前诊断出问题并自动修复，而不是静默失败。

```bash
# 先诊断（不修改任何文件）
python3 scripts/preflight_check.py 合同.docx changes.json

# 自动修正 + 生成 _fixed.json（推荐方式）
python3 scripts/preflight_check.py 合同.docx changes.json --fix

# 输出示例（发现问题）：
#   ⚠️ 扁平格式 (7 条) → 已自动转换为标准格式
#   ⚠️  [rev@24] 部分匹配 → 建议改为段落 26（匹配度 45）
#   ✅ 已生成修正版: changes_fixed.json
#   → 重新运行: python3 full_pipeline.py 合同.docx changes_fixed.json
```

**预检验做什么**：
| 步骤 | 检查内容 | 失败时 |
|------|---------|--------|
| 1/5 | JSON 格式（扁平 vs 标准） | 自动转换，覆盖原文件 |
| 2/5 | 读取原始 docx 段落（精确 XML 文本） | 退出（docx 读取失败） |
| 3/5 | paragraph_index 越界检测 | 退出（索引超出范围） |
| 4/5 | 原文精确匹配测试 | 退出并生成 _fixed.json |
| 5/5 | --fix 时写入修正版 | 输出新文件路径 |

> ⚠️ **经验教训**：`original_text` 必须从 `document.xml` 原始 `<w:t>` 提取，不能用 `contract_parser.py` 的中间输出——两者格式可能不同（空格、数字写法等），导致匹配失败。

### 应用变更

```bash
# 完整工作流（含自动预检验）：推荐方式
python3 scripts/full_pipeline.py 合同.docx changes.json \
  --output-dir ~/Desktop \
  --suffix "-审查版"

# pipeline 内部流程：步骤 3（应用修订）前内置预检验
#   发现格式错误 → 自动转换
#   发现匹配失败 → 生成 _fixed.json，退出并提示
#   全部通过 → 正常写入

# 单独应用修订+批注（仅 docx，已解压模式）
python3 scripts/apply_changes.py 合同.docx changes.json

# 可选参数：
# --dry-run  仅验证，不写入
```

### 旧版用法（已弃用）

```bash
# v2.x 版本需要手动解压/重压缩（已弃用）
# 1. 先解压 docx
python3 ~/.../docx/scripts/unpack.py 合同.docx unpacked/

# 2. 应用变更
python3 scripts/apply_changes.py unpacked/ changes.json

# 3. 重新打包
python3 ~/.../docx/scripts/pack.py unpacked/ 合同-修订批注版.docx
```

### changes.json 格式

**⚠️ 字段名必须严格遵守（与 apply_changes.py 内部实现一致）：**
- 修订条目：`original_text`（原文）+ `revised_text`（修订文）+ `comment`（可选，说明理由）
- 批注条目：`highlight_text`（高亮文本）+ `comment`（批注正文）+ `severity` + `reason` + `legal_basis`

**⚠️ JSON 生成规范（v2.8.1 新增，防止解析失败）：**
- **禁止手动拼接 JSON 字符串**，必须使用 `json.dumps(data, ensure_ascii=False, indent=2)` 生成
- **字符串值内的引号处理**：批注内容如需引用原文，用**单引号**或**全角引号**，禁止用未转义的 ASCII 双引号
- **生成后立即预检验**：调用 `preflight_check.py` 或 `json.loads()` 验证，失败则重新生成而非修复
- **反面案例**（禁止）：
  ```python
  # ❌ 错误：手动拼接，内嵌双引号未转义
  comment = '甲方修订"任何时候"均可查阅'
  json_str = f'{{"comment": "{comment}"}}'  # 解析失败！
  ```
- **正面案例**（正确）：
  ```python
  # ✅ 正确：使用 json.dumps() 自动转义
  import json
  data = {"comment": '甲方修订"任何时候"均可查阅'}
  json_str = json.dumps(data, ensure_ascii=False, indent=2)
  # 输出：{"comment": "甲方修订\"任何时候\"均可查阅"}  # 自动转义
  ```
  ```python
  # ✅ 或使用单引号/全角引号避免冲突
  data = {"comment": "甲方修订'任何时候'均可查阅"}
  ```

```json
{
  "author": "Claude",
  "date": "2026-04-18T12:00:00Z",
  "revisions": [
    {
      "paragraph_index": 42,
      "original_text": "违约金为合同总金额的 30%",
      "revised_text": "违约金为合同总金额的 20%，任一方逾期履行超过15日的，守约方有权解除合同。",
      "comment": "【修订】30%可能被认定为过高；参照《民法典》第585条调整至合理范围。"
    }
  ],
  "comments": [
    {
      "paragraph_index": 55,
      "highlight_text": "争议解决",
      "comment": "建议增加仲裁选项",
      "severity": "中风险",
      "reason": "仅约定诉讼管辖，未提供仲裁选项，限制了当事人的争议解决方式选择权",
      "legal_basis": "《仲裁法》第16条；参照《最高人民法院关于适用〈中华人民共和国民事诉讼法〉的解释》",
      "suggestion": "建议增加仲裁选项："甲乙双方约定争议可提交XX仲裁委员会仲裁，或向有管辖权的人民法院提起诉讼""
    }
  ]
}
```

**修订条目写作规范（重要）：**
- `revised_text` 必须写**完整的修订后条款文本**（可长可短，完整即可），不是只说"建议修改"
- 对于留空项（如"提前 天通知"），`revised_text` 应填入合理默认值（如"提前30天通知"），而非空着
- 对于缺失条款引用（如引用不存在的第四条），`revised_text` 应直接写入正确的核心内容，而非仅删除引用
- 对于畸高违约金，`revised_text` 应写明具体比例（如"30%"）及法律依据，而非只标注"偏高"
- 批注仅用于：①结构性提示（如条款编号断裂）；②法律背景说明；③确实无法精准定位原文的全局问题

### 生成清洁版

```bash
python3 scripts/generate_clean.py 修订批注版.docx 清洁版.docx

# 支持的模式：
# --mode accept_and_remove_comments  # 默认：接受修订+移除批注
# --mode accept_only                  # 仅接受修订
# --mode remove_comments_only        # 仅移除批注
```

### 迭代管理

```bash
# 初始化项目
python3 scripts/iteration_manager.py init 合同名-Output --name "合同名"

# 创建新一轮
python3 scripts/iteration_manager.py new-round 合同名-Output \
  --revised round-2/修订批注版.docx \
  --clean round-2/清洁版.docx

# 查看状态
python3 scripts/iteration_manager.py status 合同名-Output

# 对比两轮
python3 scripts/iteration_manager.py compare 合同名-Output 1 2

# 回滚
python3 scripts/iteration_manager.py rollback 合同名-Output 1
```

---

## 审核原则

### 核心原则

1. **修订优先原则 ★** - 所有可操作的风险，必须写入修订痕迹（Track Changes）；批注仅用于结构性提示、法律背景说明、全局问题，无法精准定位原文的情况
2. **最小修改原则** - 只改有问题的内容，不做整篇重写
3. **具体可执行** - `revised_text` 必须写完整修订后文本，填入合理默认值，不留空
4. **风险导向** - 优先处理高风险问题
5. **立场明确** - 在立场不明时标注需确认

### 修改标准

- 修改必须说明理由
- 批注必须包含建议条款
- 风险必须量化或具体化
- 格式问题统一列出

### 不接受的做法

- 整篇替换原文
- 泛泛而谈的建议（如"建议完善"）
- 不说明理由的修改
- 修改无具体条款的批注

---

## 前置确认

审核前必须确认：

- [ ] 合同类型
- [ ] 审核立场（甲方/乙方/中立）
- [ ] 输入文件（docx 格式）
- [ ] 审核轮次（第几轮）
- [ ] 是否需要清洁版
- [ ] 是否允许新增完整条款

---

## 定稿检查

输出前必须检查：

### 修订痕迹检查
- [ ] **修订痕迹（revisions）已写入 changes.json**——每条可操作风险都有对应修订条目，不是只有批注
- [ ] `revised_text` 填入完整文本和合理默认值，无留空项
- [ ] 修订痕迹数量与 changes.json 中的 revisions 数组长度一致

### 批注检查（v2.7.7 新增）
- [ ] **每条重要修订都有对应批注说明修改原因**
- [ ] 批注数量与修订数量匹配（或接近）
- [ ] 批注ID在 comments.xml 和 document.xml 中完全匹配
- [ ] document.xml 中每个批注都有完整的 commentRangeStart + commentRangeEnd + commentReference

### 多轮审核检查（v2.7.7 新增）
- [ ] 第二轮审核保留了第一轮的所有批注
- [ ] 新批注ID从上一轮最大ID+1开始（建议间隔10，如甲方0-12，乙方20-31）
- [ ] 前一轮的修订痕迹仍然可见（不被覆盖）

### 文档质量检查
- [ ] 定义词是否统一
- [ ] 编号是否连续
- [ ] 标点是否规范
- [ ] 援引是否准确
- [ ] 格式是否一致
- [ ] changes.json 是否生成
- [ ] 迭代管理是否归档

---

## 依赖说明

### 零外部依赖

所有 Python 脚本仅使用标准库：
- `zipfile` - 读写 docx
- `xml.etree.ElementTree` - XML 处理
- `json` - 数据格式
- `pathlib` - 路径操作

### 可选依赖

如需使用 `unpack.py` 和 `pack.py`，需安装 docx skill。

---

## 故障排除

| 问题 | 解决方案 |
|------|----------|
| pipeline 运行后输出 0 处修订/0 处批注 | changes.json 格式错误；运行 `preflight_check.py --fix` 自动诊断并修复 |
| 报错 `KeyError: 'revised_text'` | 字段名错误：`revised_text` 是正确字段名（不是 `new_text` 或 `replacement_text`）|
| 报错 `ValueError: list.remove(x): x not in list` | `apply_changes.py` 已修复（v2.7.2）；升级到最新版本 |
| 预检验显示"未匹配"但内容明明在文档里 | 原文 `original_text` 格式与 XML 不一致；运行 `--fix` 自动修正段落索引 |
| 修订痕迹不显示 | 检查 Word"审阅"→"显示标记"是否开启 |
| 批注气泡不显示 | 检查 comments.xml 是否正确注册；v2.5.0 已修复 WPS/Word 兼容性问题 |
| WPS 打开的批注在 Word 中丢失 | v2.5.0 已修复：补充 `commentRangeStart/End`、`w14:commentId`、`commentsExtended.xml` |
| 清洁版生成失败 | 检查输入文件是否损坏 |
| 迭代状态丢失 | 检查 manifest.json 格式 |
| 只有批注没有修订 | 这是**错误的输出模式**；应将所有可操作风险写入 `revisions` 数组，`comments` 仅作补充说明 |
| **修订有痕迹但无批注说明原因** | v2.7.7 修复：每条重要修订都必须有对应批注；检查 changes.json 中 revisions[].comment 是否为空 |
| **批注在 Word 中看不到** | 批注ID不匹配；检查 comments.xml 和 document.xml 中的 ID 是否一致 |
| **第二轮审核丢失第一轮批注** | v2.7.7 修复：第二轮必须保留第一轮所有批注，新批注ID从上一轮最大值+1开始 |

---

## 许可证

本 Skill 以 MIT 协议发布。

**核心审核方法论**基于法律审核通用实践，属于行业公共知识。

**引擎实现代码**由 OpenClaw AI 编写，以 MIT 协议开源。

---

*版本 2.8.1 | 2026-04-20*
