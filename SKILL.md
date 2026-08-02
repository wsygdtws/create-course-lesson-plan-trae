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
7. **下载位置**: 当任务涉及从百度网盘、云盘或任何网络位置下载文档、素材或资源到本地时，**必须先主动询问用户**下载到哪个本地目录（如 `download_from_baidunet/`、项目根目录下指定子目录等），不得擅自决定下载路径。询问时应给出默认建议路径（如项目根目录下的 `download_from_baidunet/`）供用户确认或修改。只有在用户明确给出下载位置（或确认使用默认建议路径）后才能执行下载。

**If ANY key information is unclear, STOP and ask the user. Do NOT guess or assume.**

## Tools (独立脚本)

本 skill 的核心功能已封装为独立 Python 脚本，位于 `tools/` 目录：

| 脚本 | 功能 | 用法 |
|------|------|------|
| `tools/word_lesson_tool.py` | Word COM 操作：读取/填充教案模板、单元格设置、图片插入、图片质量筛选 | `from word_lesson_tool import process_week, read_doc_tables, set_simple_cell, set_complex_cell, add_image_to_cell, kill_word, is_good_image, pick_images, calc_width` |
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

## Content Quality Requirements (内容质量要求)

教案内容必须充实、具体，不得空洞。以下是对 Table 3 各环节的详细要求：

### 课题 (Table 1 Row 3 Col 2) 必须与授课计划保持一致

**教案首页课题必须与授课计划中对应周的"章节（课题）"完全一致。**

授课计划中，一个课题可能跨越多周（连续几周的"章节（课题）"列只在第一周填写，后续周为空），教案课题应继承该课题，即使连续好几周都是同一课题也没错。教案的具体教学内容则根据授课计划中对应周的"主要内容"分别设计。

**正确做法**：
- 读取授课计划表3，找到对应周的"章节（课题）"列
- 如果该周课题为空，向上查找最近一个非空课题（继承上一周的课题）
- 教案首页课题必须与授课计划中的课题完全一致（包括标点符号）
- 例如：授课计划第3-5周课题均为"AiNova声光与变速控制"，则这3周的教案课题都应填写"AiNova声光与变速控制"，但每周教案的80min实操内容根据该周"主要内容"分别设计

### 40min 环节（Row 4 - 获取信息）内容要求

40min 环节是教师讲授环节，内容必须充实、技术细节丰富：

1. **内容长度**：至少 600 字符，避免过于简略
2. **技术细节**：必须包含具体的参数、命令、操作步骤等，不能只有概念性描述
   - 例如：写"速度控制"要列出具体命令 `rostopic pub /jetauto_controller/cmd_vel geometry_msgs/Twist`、参数范围（线速度 -0.7~0.7、角速度 -3.5~3.5）等
   - 例如：写"配件安装"要列出具体螺丝规格（M3×6）、接口编号（②号接口、⑤号接口）等
3. **分点结构**：用编号分点（1. 2. 3.）组织内容，每点有小标题（加粗）
4. **必须包含教师活动和学生活动**：
   - `('教师活动：', True)` - 描述教师讲授、演示的具体内容
   - `('学生活动：', True)` - 描述学生听讲、观察、记录等具体行为

### 80min 环节（Row 5 - 实施任务）内容要求

80min 环节是学生实操环节，必须依据教学计划当周课题，结合网盘资料或网络检索结果，按照工学一体化特点和技工院校学生特点，设计丰富详细的实操内容：

1. **内容长度**：至少 400 字符，实操步骤必须具体可执行
2. **三段式结构**：
   - `1.安全教育` - 列出与本周内容相关的安全注意事项（2-3条）
   - `2.分组实操` - 详细列出每一步操作（依据课题资料设计检查清单、参数测试、数据记录等环节）
   - `3.填写工作单` - 说明工作单填写要求
3. **实操内容设计依据**：
   - 依据教学计划中当周课题，结合百度网盘下载的 PDF 资料（产品手册、教程文档等）或网络检索结果，提炼出具体的操作步骤、检查项目、参数测试、数据记录等实操内容
   - 例如：课题为"速度控制"，应从 ROS 教程资料中提炼出 `rostopic pub` 命令、线速度/角速度参数范围、矩形与圆形轨迹行驶方法等实操步骤
   - 例如：课题为"AiNova运动控制"，应从产品资料中提炼出转速参数（50/70/30rpm）、转弯逻辑（电机1正电机2负为右转）、方形路线测试等实操内容
   - 实操步骤要体现工学一体化特点，让学生在"做中学"，每一步都明确操作对象、操作方法、预期结果
4. **必须包含教师活动和学生活动**：
   - `('教师活动：', True)` - 描述教师巡回指导、示范操作、检查记录等具体行为
   - `('学生活动：', True)` - 描述学生按步骤实施、记录数据、协作完成等具体行为

### 10min 环节（Row 3 - 明确任务）和 Row 6/7 也需包含师生活动

- **Row 3 (10min)**：安全教育 + 明确任务 + 教师活动 + 学生活动
- **Row 6 (10min)**：质量检验 + 教师活动 + 学生活动
- **Row 7 (20min)**：评价反馈 + 教师活动 + 学生活动

## Image Quality Requirements (配图质量要求)

配图必须清晰可见，不得使用尺寸异常的图片。以下是图片选择和处理要求：

### 问题背景

从 PDF 提取的图片中常包含横幅图（如页眉、分隔线，尺寸如 1267×80），这类图片在 Word 中以默认宽度 180pt 显示时，高度仅约 18pt，几乎无法看清。

### 图片筛选规则

使用 `is_good_image()` 和 `pick_images()` 函数筛选合格图片：

```python
from word_lesson_tool import is_good_image, pick_images, calc_width

# 筛选条件：
# - 高度 >= 200 像素
# - 宽高比 <= 4.5（过滤横幅图）
# - 高宽比 <= 3.0（过滤过高的竖图）

# 按前缀筛选合格图片
imgs = pick_images(img_dir, 'JetAuto_配件安装教程', count=3, min_h=200)
# 返回 [(filepath, w, h), ...]

# 根据宽高比计算合适的显示宽度
for fp, w, h in imgs:
    width_pt = calc_width(w, h)  # 竖图150pt, 普通图180pt, 超宽图240pt
    images.append(('配件安装示意图', fp, width_pt))
```

### 图片显示尺寸要求

- 插入 Word 后的图片**显示高度不得低于 50pt**（约 1.76cm），否则视为异常
- 使用 `calc_width(w, h)` 根据宽高比自动选择合适的显示宽度
- 如果某个前缀的所有图片均不合格，应更换其他前缀的图片或从网络下载替代配图

### 验证图片质量

生成教案后，应验证 InlineShapes 的显示尺寸：

```python
import win32com.client
word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open(filepath)
for i in range(1, doc.InlineShapes.Count + 1):
    h = doc.InlineShapes(i).Height
    if h < 50:
        print(f"图片 {i} 异常：高度仅 {h:.1f}pt，需更换")
```

### 无法调整时的替代方案

如果某个主题实在没有合格的配图：
1. 从其他相关 PDF 中寻找替代图片（更换前缀搜索）
2. 从网络下载相关配图
3. 绘制示意图替代

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
| 配图显示过小/看不清 | 使用 `is_good_image()` 过滤横幅图，用 `calc_width()` 按宽高比选宽度，显示高度≥50pt |
| 课题与内容不符 | 教案课题必须与授课计划中对应周的"章节（课题）"完全一致，空课题周继承上一周课题；教学内容根据"主要内容"分别设计 |
| 80min实操内容空洞 | 依据教学计划当周课题，结合网盘PDF资料/网络检索结果，按工学一体化特点设计检查清单、参数测试、数据记录等实操步骤，包含教师/学生活动 |
| table3_rows IndexError | 不要填充 row 1/2/8（模板已有内容），仅填充 row 3-7；col2 必须是 `(text, is_bold)` 元组列表 |
| set_complex_cell paragraphs[0] 报错 | col2 不能传纯字符串，必须是 `[(text, is_bold), ...]` 列表 |
