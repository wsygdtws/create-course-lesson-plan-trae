# Lesson Plan Generator / 教案生成器

[English](#english) | [中文](#中文)

---

<a id="english"></a>
## English

A [Trae IDE Skill](https://docs.trae.ai/en/skills) that generates Chinese vocational-school lesson plan `.doc` files from Word templates, following the **工学一体化 (Work-Learning Integration)** teaching style.

### Features

- Generates weekly lesson plans from `.doc` templates (regular and 机动周 / flexibility-week templates)
- Fills 3 tables per template: header info, signature block, and teaching content
- Inserts images (extracted from PDFs or generated) with captions into the "获取信息" section
- Preserves template formatting by only modifying cell text and font properties
- Auto-retry mechanism for Word COM instability (RPC server unavailable)
- Keeps Table 1 and Table 2 on the same page to avoid cross-page layout breaks
- Learning objectives follow 工学一体化 style ("学生能够..." with measurable verbs)
- Mandatory user confirmation for uncertain key information before generation

### Installation

1. Copy the `lesson-plan-generator` folder to your Trae skills directory:
   ```
   <workspace>/.trae/skills/lesson-plan-generator/
   ```
2. Restart Trae IDE — the skill is auto-detected.

### Prerequisites

- **Windows** with Microsoft Word installed (uses `win32com.client`)
- Python packages:
  ```bash
  pip install pywin32 PyMuPDF requests
  ```
- Optional: Baidu Netdisk MCP server configured in Trae for course content download

### Usage

Just ask Trae in natural language, e.g.:

> "根据授课计划编写第1-5周教案"

The skill will be auto-invoked. **Before generating, it MUST ask you about:**

1. 课时安排 (period schedule — e.g. 周二、周三下午7-10节, 周四上午1-4节)
2. 班级信息 (class names, multiple separated by "、")
3. 模板选择 (which `.doc` template, including 机动周 template)
4. 授课日期 (dates confirmed from school calendar)
5. 课题内容来源 (百度网盘 path or local files)
6. 排课方式 (which class on which day)

### Workflow Overview

1. **Gather information** — read teaching plan, templates, school calendar, course content
2. **Ask user** about any uncertain key information (CRITICAL RULE: never assume)
3. **Generate lesson plans** — fill Table 1 (header), Table 2 (signature), Table 3 (teaching activities)
4. **Insert images** — extracted from PDFs via PyMuPDF or generated
5. **Download course content** from Baidu Netdisk via MCP or PCS API
6. **Verify output** — check row counts, page layout, image insertion, date format

### Template Structure

Each template contains 3 tables:

| Table | Purpose | Rows × Cols |
|-------|---------|-------------|
| Table 1 | Header info (课程/日期/班级/课题/课时/教学方法/课后作业/学习目标/重点/难点/教学回顾/备注) | 10 × 4 |
| Table 2 | Signature (授课教师/部长/主任签名/提交/审阅日期) | 2 × 3 |
| Table 3 | Teaching content — 工学一体化 phases | 8 (regular) / 7 (机动周) |

**Table 3 phases (regular template):**

| Row | Time | Phase |
|-----|------|-------|
| 1 | — | Headers (DO NOT MODIFY) |
| 2 | — | 课前六件事 (DO NOT MODIFY) |
| 3 | 10 min | 安全教育 + 明确任务 |
| 4 | 40 min | 获取信息 (main content + images) |
| 5 | 80 min | 实施任务 (hands-on practice) |
| 6 | 10 min | 质量检验 |
| 7 | 20 min | 评价反馈 |
| 8 | — | 课后四件事 (DO NOT MODIFY) |

### Font Specifications

| Element | Font | Size | Bold |
|---------|------|------|------|
| Body text | 宋体 (SimSun) | 12pt | No |
| Section headers | 宋体 | 12pt | Yes |
| Numbered sub-items | 宋体 | 12pt | Yes |
| Image captions | 宋体 | 10.5pt | No |
| Table headers | 宋体 | 12pt | Yes |
| Title "教案内容" | 黑体 (SimHei) | 22pt | Yes (in template) |

### Date Format Example

```
2026年9月8日、9月9日、9月10日 第2周 周二、周三7-10节、周四1-4节
```

### Learning Objectives Style

Objectives MUST start with "通过X节课的学习，学生能够：" and use measurable verbs.

**Example:**
```
通过4节课的学习，学生能够：
1.对照AiNova和JetAuto实物，指出感知、决策、执行和通信四大系统对应的主要硬件模块；
2.按照实训室安全规范，正确完成设备取放和通电检查；
3.填写设备认识记录表，准确记录两种实训平台的主要模块名称及功能。
```

### Common Issues & Solutions

| Issue | Solution |
|-------|----------|
| Word COM crashes (RPC unavailable) | Use `kill_word()` + retry (up to 3 times) |
| Table 2 spills to page 2 | Set rows 6-9 `HeightRule=2` (Exactly), reduce height |
| Duplicate images in output | Always `shutil.copy2()` template fresh before modifying |
| Content overflow in cells | Keep text concise, use bullet points |
| Chinese encoding issues | Add `# -*- coding: utf-8 -*-` to all Python scripts |

### File Structure

```
lesson-plan-generator/
├── SKILL.md      # Skill definition (workflow + Python code)
└── README.md     # This file
```

### License

MIT License — feel free to use and modify.

---

<a id="中文"></a>
## 中文

一个 [Trae IDE Skill](https://docs.trae.ai/en/skills)，基于 Word 模板生成中国职业院校教案 `.doc` 文件，遵循**工学一体化**教学风格。

### 功能特点

- 基于 `.doc` 模板生成每周教案（常规模板和机动周模板）
- 填充模板中的 3 个表格：表头信息、签名区、教学内容
- 在"获取信息"区域插入带说明的图片（从 PDF 提取或 AI 生成）
- 仅修改单元格文本和字体属性，保留模板原有格式
- 针对 Word COM 不稳定（RPC 服务器不可用）自动重试
- 保持表 1 与表 2 在同一页，避免跨页布局错乱
- 学习目标遵循工学一体化风格（"学生能够..."开头，动词可衡量）
- 生成前必须就关键不确定信息主动询问用户

### 安装方法

1. 将 `lesson-plan-generator` 文件夹复制到 Trae 的 skills 目录：
   ```
   <工作区>/.trae/skills/lesson-plan-generator/
   ```
2. 重启 Trae IDE —— 技能将自动被识别。

### 环境要求

- **Windows** 系统，已安装 Microsoft Word（使用 `win32com.client`）
- Python 依赖包：
  ```bash
  pip install pywin32 PyMuPDF requests
  ```
- 可选：在 Trae 中配置百度网盘 MCP 服务，用于下载课程资料

### 使用方法

直接用自然语言对 Trae 说即可，例如：

> "根据授课计划编写第1-5周教案"

技能会自动触发。**生成前，它必须询问你以下信息：**

1. 课时安排（如：周二、周三下午7-10节，周四上午1-4节）
2. 班级信息（多个班级用"、"分隔）
3. 模板选择（使用哪个 `.doc` 模板，是否有机动周模板）
4. 授课日期（根据校历确认具体日期）
5. 课题内容来源（百度网盘路径或本地文件）
6. 排课方式（哪天哪个班级上课）

### 工作流程概览

1. **收集信息** —— 读取授课计划、模板、校历、课题资料
2. **询问用户** —— 就任何不确定的关键信息提问（关键规则：切勿自作主张）
3. **生成教案** —— 填充表 1（表头）、表 2（签名）、表 3（教学活动）
4. **插入图片** —— 使用 PyMuPDF 从 PDF 提取，或 AI 生成
5. **下载课题资料** —— 通过 MCP 或 PCS API 从百度网盘下载
6. **验证输出** —— 检查行数、页面布局、图片插入、日期格式

### 模板结构

每个模板包含 3 个表格：

| 表格 | 用途 | 行 × 列 |
|------|------|---------|
| 表 1 | 表头信息（课程/日期/班级/课题/课时/教学方法/课后作业/学习目标/重点/难点/教学回顾/备注） | 10 × 4 |
| 表 2 | 签名（授课教师/部长/主任签名/提交/审阅日期） | 2 × 3 |
| 表 3 | 教学内容 —— 工学一体化各阶段 | 8（常规）/ 7（机动周） |

**表 3 阶段划分（常规模板）：**

| 行 | 时间 | 阶段 |
|----|------|------|
| 1 | — | 表头（请勿修改） |
| 2 | — | 课前六件事（请勿修改） |
| 3 | 10 分钟 | 安全教育 + 明确任务 |
| 4 | 40 分钟 | 获取信息（主要内容 + 图片） |
| 5 | 80 分钟 | 实施任务（动手实践） |
| 6 | 10 分钟 | 质量检验 |
| 7 | 20 分钟 | 评价反馈 |
| 8 | — | 课后四件事（请勿修改） |

### 字体规范

| 元素 | 字体 | 字号 | 加粗 |
|------|------|------|------|
| 正文 | 宋体 | 12pt | 否 |
| 阶段标题 | 宋体 | 12pt | 是 |
| 编号子项 | 宋体 | 12pt | 是 |
| 图片说明 | 宋体 | 10.5pt | 否 |
| 表格表头 | 宋体 | 12pt | 是 |
| 标题"教案内容" | 黑体 | 22pt | 是（模板已有） |

### 日期格式示例

```
2026年9月8日、9月9日、9月10日 第2周 周二、周三7-10节、周四1-4节
```

### 学习目标风格

学习目标必须以"通过X节课的学习，学生能够："开头，使用可衡量的动词。

**示例：**
```
通过4节课的学习，学生能够：
1.对照AiNova和JetAuto实物，指出感知、决策、执行和通信四大系统对应的主要硬件模块；
2.按照实训室安全规范，正确完成设备取放和通电检查；
3.填写设备认识记录表，准确记录两种实训平台的主要模块名称及功能。
```

### 常见问题与解决方案

| 问题 | 解决方案 |
|------|----------|
| Word COM 崩溃（RPC 不可用） | 使用 `kill_word()` + 重试（最多 3 次） |
| 表 2 跑到第 2 页 | 将第 6-9 行 `HeightRule` 设为 2（固定值），减小行高 |
| 输出图片重复 | 修改前始终用 `shutil.copy2()` 重新复制模板 |
| 单元格内容溢出 | 保持文字简洁，使用项目符号 |
| 中文编码问题 | 所有 Python 脚本添加 `# -*- coding: utf-8 -*-` |

### 文件结构

```
lesson-plan-generator/
├── SKILL.md      # 技能定义（工作流程 + Python 代码）
└── README.md     # 本文件
```

### 许可证

MIT License —— 欢迎自由使用和修改。
