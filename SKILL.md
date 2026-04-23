---
name: contract-reviewer-wudi
display_name: 合同审核-WUDI
version: 2.9.1
description: |
  中文合同审核 Skill - 支持修订痕迹、批注气泡、清洁版生成与多轮迭代管理。
  覆盖41种合同类型，350+关键条款，229+常见风险，82个法律依据。
  核心能力：合同结构解析、合同类型检测、条款智能推荐、交叉引用检查、修订痕迹生成、批注气泡写入、PDF/扫描件OCR、多轮迭代管理。

license: GPL-3.0
author: WUDI
---## 更新日志

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

> 完整文档请参阅 GitHub 仓库。
