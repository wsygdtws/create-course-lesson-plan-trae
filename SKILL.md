---
name: "lesson-plan-generator"
description: "Generates lesson plan .doc files from templates. Invoke when user asks to create/write/generate lesson plans (教案), fill in teaching plan content, or modify existing lesson plans. MUST ask user about uncertain info before proceeding."
---

# Lesson Plan Generator (教案生成器)

Generates lesson plan `.doc` files based on Word templates, following the 工学一体化 (work-learning integration) style used in Chinese vocational schools.

## CRITICAL RULE: Ask Before Assuming

**Before generating any lesson plan, you MUST ask the user about the following uncertain information. NEVER make assumptions:**

1. **课时安排**: How many class periods per week? Which day/period? (e.g., "周二、周三下午7-10节，周四上午1-4节")
2. **班级信息**: Which classes? Multiple classes separated by "、"
3. **模板选择**: Which template to use? Is there a 机动周 (flexibility week) template?
4. **授课日期**: Confirm dates from the school calendar (校历)
5. **课题内容来源**: Where to find course content? (百度网盘 path, local files, etc.)
6. **排课方式**: Are all periods on the same day or spread across days? Which class on which day?

**If ANY key information is unclear, STOP and ask the user. Do NOT guess or assume.**

## Tools (独立脚本)

本 skill 的核心功能已封装为独立 Python 脚本，位于 `tools/` 目录：

| 脚本 | 功能 | 用法 |
|------|------|------|
| `tools/word_lesson_tool.py` | Word COM 操作：读取/填充教案模板、单元格设置、图片插入 | `from word_lesson_tool import process_week, read_doc_tables, set_simple_cell, set_complex_cell, add_image_to_cell, kill_word` |
| `tools/pdf_tool.py` | PDF 文本与图片提取（PyMuPDF） | `from pdf_tool import extract_text, extract_images, batch_extract` |
| `tools/baidu_download_tool.py` | 百度网盘文件下载（PCS API） | `from baidu_download_tool import download_file, list_files` |

**调用方式**：将 `tools/` 加入 Python path，或复制脚本到工作目录。

## Complete Workflow

### Step 1: Gather Information
1. **授课计划 (.docx)** - course name, classes, teacher, weekly topics, homework
2. **教案模板 (.doc)** - regular template + 机动周 template
3. **校历 (.xlsx)** - exact dates, holidays
4. **课题内容** - from 百度网盘 or local PDFs

### Step 2: Ask User About Uncertain Info
Confirm 课时安排, 班级, 模板, 日期, 内容来源, 排课方式.

### Step 3: Extract Course Content

```python
import sys; sys.path.insert(0, r'.trae/skills/lesson-plan-generator/tools')
from pdf_tool import extract_text, extract_images

# 提取 PDF 文本
text = extract_text('download_from_baidunet/1.1.1_认识AiNova.pdf')

# 提取图片
imgs = extract_images('download_from_baidunet/1.1.1_认识AiNova.pdf', 'lesson_images', '1_1_1_认识AiNova')
```

### Step 4: Generate Lesson Plans

#### Template Structure (3 tables per template)
- **Table 1** (10 rows × 4 cols): Header info
- **Table 2** (2 rows × 3 cols): Signature
- **Table 3**: Lesson content (工学一体化 style)

#### Table 3 Structure (Regular - 8 rows / 机动周 - 7 rows)
| Row | Content |
|-----|---------|
| 1 | Headers (DO NOT MODIFY) |
| 2 | 课前六件事 (DO NOT MODIFY) |
| 3 | 10min: 安全教育 + 明确任务 |
| 4 | 40min: 获取信息 (main content + images) |
| 5 | 80min: 实施任务 (hands-on practice) |
| 6 | 10min: 质量检验 |
| 7 | 20min: 评价反馈 |
| 8 | 课后四件事 (DO NOT MODIFY) |

#### Generate using word_lesson_tool

```python
import sys; sys.path.insert(0, r'.trae/skills/lesson-plan-generator/tools')
from word_lesson_tool import process_week

data = {
    'course': '智能网联汽车概论',
    'date': '2026年9月8日、9月9日、9月10日 第2周 周二、周三7-10节、周四1-4节',
    'class_name': '24新能源高级3班、4班、2班',
    'topic': '智能网联汽车与实训安全',
    'hours': '4',
    'methods': '讲授法、小组合作学习',
    'tools': '多媒体、实物、微课',
    'homework': '认识两种实训平台',
    'homework_time': '1学时',
    'objectives': '通过4节课的学习，学生能够：1.对照...；2.按照...；3.填写...',
    'key_points': '1.感知、决策、执行和通信的基本关系；2.实训室安全规范。',
    'difficulties': '理解感知、决策、执行和通信的关系。',
    'table3_rows': [
        (3, ['10min', [('安全教育...', True),('明确任务...', False)], '讲授法、PPT', '激发兴趣']),
        (4, ['40min', [('1.基本概念', True),('...', False)], '讲授法、PPT', '掌握内容']),
        # ...
    ]
}
images = [('AiNova结构图', 'lesson_images/1_1_1_xxx.png', 180)]
process_week('教案模板.doc', '智能网联汽车概论_第2周教案.doc', data, images)
```

**关键规则**：
- **不要**修改 row.Height/HeightRule/AllowBreakAcrossPages（模板行格式）
- **不要**修改 cell 内容的 paragraph.KeepWithNext/SpaceBefore/SpaceAfter（内容单元格）
- **仅**修改 cell text 和 font 属性 (Name, NameFarEast, Size, Bold)
- "教案内容"标题必须位于第2页顶部（process_week 已内置 `PageBreakBefore=True` 自动处理，确保不留在第1页末尾）
- 用 `shutil.copy2()` 复制模板再修改副本
- Word COM 不稳定，已内置 `kill_word()` + 重试机制
- 逐个文件处理，避免 COM 连接问题

## Font and Format Requirements

| Element | Font | Size | Bold |
|---------|------|------|------|
| Body text | 宋体 | 12pt | No |
| Section headers | 宋体 | 12pt | Yes |
| Numbered sub-items | 宋体 | 12pt | Yes |
| Image captions | 宋体 | 10.5pt | No |
| Title "教案内容" | 黑体 | 22pt | Yes (in template) |

## Learning Objectives Style (工学一体化)

MUST start with "通过X节课的学习，学生能够：" + specific measurable verbs (not "学会"/"了解" alone).

## Common Issues

| Issue | Solution |
|-------|----------|
| Word COM crashes | `kill_word()` + retry (内置) |
| Table 2 on page 2 | Set rows 6-9 HeightRule=2, reduce height (内置) |
| "教案内容"留在第1页末尾 | process_week 已内置：`PageBreakBefore=True` + `KeepWithNext=True` 强制标题到第2页顶部 |
| Duplicate images | Always `shutil.copy2()` fresh template |
| Chinese encoding | `# -*- coding: utf-8 -*-` in all scripts |
