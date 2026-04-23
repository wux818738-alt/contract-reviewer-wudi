---
name: contract-reviewer-wudi
display_name: 合同审核-WUDI
version: 2.9.3
description: |
  中文合同审核 Skill - 支持修订痕迹、批注气泡、清洁版生成与多轮迭代管理。覆盖41种合同类型、350+关键条款、229+常见风险、82个法律依据。核心输出为 Word Track Changes（修订痕迹），不接受"只有批注、没有修订"的输出方式。
license: GPL-3.0
author: WUDI
homepage: https://github.com/wux818738-alt/contract-reviewer-wudi
---

# 合同审核 WUDI | Contract Reviewer WUDI

> 以 **修订痕迹（Track Changes）** 为核心输出的中文合同审核 Skill。
> 审核修改直接写入 Word 文档，红删绿增，一目了然。

## ★ 核心输出原则

**修订痕迹优先**。每次审核必须将具体修改建议写入文档修订痕迹；批注仅用于补充说明（法律背景、结构性提示等无法用修订表达的内容）。不接受"只有批注、没有修订"的输出方式。

## 核心能力

1. **合同结构解析** - 提取条款、定义词、交叉引用
2. **合同类型检测** - 自动识别41种合同类型
3. **条款智能推荐** - 内置22个专业条款模板
4. **交叉引用检查** - 检测孤立引用、缺失条款
5. **修订痕迹生成** - 直接写入 Word Track Changes
6. **批注气泡写入** - Word Comment（补充说明）
7. **PDF/扫描件 OCR** - Vision/Tesseract 自动转换
8. **多轮迭代管理** - 版本跟踪、对比、回滚

## 效果展示

<p align="center">
  <img src="https://raw.githubusercontent.com/wux818738-alt/contract-reviewer-wudi/main/images/01_input.png" width="400" alt="合同输入">
  <img src="https://raw.githubusercontent.com/wux818738-alt/contract-reviewer-wudi/main/images/02_type_detection.png" width="400" alt="类型识别">
</p>

<p align="center">
  <img src="https://raw.githubusercontent.com/wux818738-alt/contract-reviewer-wudi/main/images/03_risk_scan.png" width="400" alt="风险扫描">
  <img src="https://raw.githubusercontent.com/wux818738-alt/contract-reviewer-wudi/main/images/04_track_changes.png" width="400" alt="修订痕迹">
</p>

<p align="center">
  <img src="https://raw.githubusercontent.com/wux818738-alt/contract-reviewer-wudi/main/images/05_comments.png" width="400" alt="批注气泡">
  <img src="https://raw.githubusercontent.com/wux818738-alt/contract-reviewer-wudi/main/images/06_clean_version.png" width="400" alt="清洁版输出">
</p>

## 覆盖范围

| 类型 | 数量 |
|------|------|
| 合同类型 | 41 种（买卖、租赁、建设工程、技术服务、劳动、股权、PPP、金融等） |
| 关键条款 | 350+ 条 |
| 常见风险 | 229+ 点 |
| 法律依据 | 82 个（《民法典》合同编为主） |

## 使用方式

### OpenClaw 用户

```
openclaw skills install contract-reviewer-wudi
```

或手动 clone：

```bash
git clone https://github.com/wux818738-alt/contract-reviewer-wudi.git ~/.qclaw/skills/contract-reviewer-wudi
```

### 触发方式

上传 Word/PDF 合同文件，AI 自动：
1. 识别合同类型 → 匹配审核规则
2. 逐条扫描 → 识别风险条款
3. 生成修订 → 写入 Word Track Changes
4. 添加批注 → 法律背景说明
5. 输出清洁版 → 接受全部修改

### 输出文件

- `{合同名}-修订痕迹版.docx` - 包含所有修订痕迹和批注
- `{合同名}-清洁版.docx` - 接受全部修改后的干净文档

## 目录结构

```
contract-reviewer-wudi/
├── SKILL.md                    # Skill 主文件（本文档）
├── README.md                   # GitHub 介绍（含效果截图）
├── scripts/                    # Python 脚本引擎
│   ├── contract_parser.py      # 合同结构解析
│   ├── apply_changes.py        # 修订痕迹写入 ★
│   ├── generate_clean.py       # 清洁版生成
│   └── full_pipeline.py        # 完整流水线
├── references/                 # 详细文档
│   ├── review-playbook.md      # 审核操作手册
│   ├── engine-guide.md         # 引擎使用指南
│   ├── clause-library/         # 22 个条款模板
│   └ contract-types/           # 41 种合同定义
└── images/                     # 效果截图
```

## 详细文档

- **审核操作手册**: `references/review-playbook.md`
- **引擎使用指南**: `references/engine-guide.md`
- **输出规范**: `references/output-spec.md`
- **GitHub 完整介绍**: https://github.com/wux818738-alt/contract-reviewer-wudi

## 许可证

GPL-3.0 © 2026 WUDI

---

> 本 Skill 的审核方法论基于法律审核通用实践，属于行业公共知识。
