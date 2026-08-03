---
name: "lesson-plan-generator"
description: "Generates lesson plan .doc files from templates. Invoke when user asks to create/write/generate lesson plans (教案), fill in teaching plan content, or modify existing lesson plans. MUST ask user about uncertain info before proceeding."
---

# Lesson Plan Generator (教案生成器)

基于 Word 模板生成 `.doc` 教案文件，遵循技工院校工学一体化风格。

## CRITICAL RULE: 先问再做

**生成教案前，必须向用户确认以下信息，不得自行假设：**

1. **上课日期（最重要）**：每周几上课？周几的第几节？（如"周二、周三下午7-10节，周四上午1-4节"）。**必须先确认此信息，才能与校历匹配判断哪些周正常上课、哪些周机动。**
2. **班级信息**：哪些班？多班用"、"分隔
3. **模板选择**：用哪个模板？是否有机动周模板？
4. **课题内容来源**：课程内容从哪获取？（百度网盘路径、本地文件等）
5. **下载位置**：涉及下载时，必须先询问下载到哪个本地目录，给出默认建议（如 `download_from_baidunet/`）供用户确认。

**任何关键信息不明确时，停止并询问用户。**

## Tools (独立脚本)

核心功能封装为 `tools/` 目录下的 Python 脚本：

| 脚本 | 功能 | 关键函数 |
|------|------|----------|
| `word_lesson_tool.py` | Word COM 操作：读取/填充模板、单元格设置、图片插入、质量筛选 | `process_week, read_doc_tables, set_simple_cell, set_complex_cell, add_image_to_cell, kill_word, is_good_image, pick_images, calc_width` |
| `pdf_tool.py` | PDF 文本与图片提取（PyMuPDF） | `extract_text, extract_images, batch_extract` |
| `baidu_download_tool.py` | 百度网盘文件下载（PCS API） | `download_file, list_files` |

调用方式：`sys.path.insert(0, r'.trae/skills/lesson-plan-generator/tools')`

## Complete Workflow

### Step 1: 确认上课日期与收集资料

向用户确认上课日期后，收集以下文件：
- **授课计划 (.docx)**：课程名、班级、教师、每周课题、主要内容、作业
- **教案模板 (.doc)**：常规模板 + 机动周模板
- **校历 (.xlsx)**：每周日期范围、工作内容、放假信息
- **课题内容**：百度网盘或本地 PDF

### Step 2: 校历分析与机动周判定

**这是教案正确性的基础。** 根据上课日期和校历，逐周判定是正常上课还是机动：

| 情况 | 判定 |
|------|------|
| 校历"工作内容"为"正式按课程表上课"**之前**的周 | 机动 |
| 上课日期**在**放假日期范围内 | 机动 |
| 上课日期**不在**放假日期范围内（即使该周有放假） | 正常上课 |
| 校历"工作内容"为"期末考试" | 机动 |

**关键**：放假周不能简单判定为机动，必须核对上课日期是否落在放假日期内。例如国庆放假10月1日-7日，若上课日期为10月8日（周四），则该周正常上课。

**机动周使用机动周模板**（Table 3 有 7 行），正常周使用常规模板（Table 3 有 8 行）。

校历读取脚本：
```python
import openpyxl

wb = openpyxl.load_workbook(filepath)
ws = wb.active
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0]:  # 周次非空
        week, date_range, work, holiday = row[0], str(row[1] or ''), str(row[2] or ''), str(row[3] or '')
        # 根据 work/holiday 和上课日期判断是否机动
```

### Step 3: 与授课计划匹配

授课计划"教学周历"表包含每周的章节（课题）、主要内容、课时、课外作业。匹配规则：

1. **周次对应**：校历周次 = 授课计划周次 = 教案周次（三者一一对应）
2. **日期**：教案日期 = 校历对应周次的上课日期，格式如 `2026年9月8日、9月9日、9月10日 第2周 周二、周三7-10节、周四1-4节`
3. **课题**：教案课题 = 授课计划对应周的"章节（课题）"。若该周课题为空，**向上继承**最近一个非空课题（一个课题可跨多周，教案课题保持一致，但教学内容按每周"主要内容"分别设计）
4. **作业**：教案作业 = 授课计划对应周的"课外作业"，**必须包含"（1学时）"后缀**。Table 1 中 Row 5 Col 2 填完整作业（如 `认识两种实训平台（1学时）`），Row 5 Col 4 填 `1学时`

### Step 4: 提取课程内容

```python
import sys; sys.path.insert(0, r'.trae/skills/lesson-plan-generator/tools')
from pdf_tool import extract_text, extract_images

text = extract_text('download_from_baidunet/1.1.1_认识AiNova.pdf')
imgs = extract_images('download_from_baidunet/1.1.1_认识AiNova.pdf', 'lesson_images', '1_1_1_认识AiNova')
```

### Step 5: 生成教案

#### 模板结构（3 张表）

| 表 | 结构 | 说明 |
|----|------|------|
| Table 1 | 10行×4列 | 首页信息（课程、日期、课题、作业等） |
| Table 2 | 2行×3列 | 签名 |
| Table 3 | 常规8行 / 机动7行 | 教学内容 |

Table 3 行布局（常规模板）：

| Row | 内容 | 是否修改 |
|-----|------|----------|
| 1 | 表头 | ❌ 不修改 |
| 2 | 课前六件事 | ❌ 不修改 |
| 3 | 10min：安全教育 + 明确任务 | ✅ |
| 4 | 40min：获取信息（含配图） | ✅ |
| 5 | 80min：实施任务（实操） | ✅ |
| 6 | 10min：质量检验 | ✅ |
| 7 | 20min：评价反馈 | ✅ |
| 8 | 课后四件事 | ❌ 不修改 |

#### 生成代码

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
    'homework': '认识两种实训平台（1学时）',
    'homework_time': '1学时',
    'objectives': '通过4节课的学习，学生能够：1.对照...；2.按照...；3.填写...',
    'key_points': '1.感知、决策、执行和通信的基本关系；2.实训室安全规范。',
    'difficulties': '理解感知、决策、执行和通信的关系。',
    'table3_rows': [
        (3, ['10min', [('安全教育...', True),('明确任务...', False)], '讲授法、PPT', '激发兴趣']),
        (4, ['40min', [('1.基本概念', True),('...', False)], '讲授法、PPT', '掌握内容']),
        # row 5-7 同理
    ]
}
images = [('AiNova结构图', 'lesson_images/1_1_1_xxx.png', 180)]
process_week('教案模板.doc', '智能网联汽车概论_第2周教案.doc', data, images)
```

**生成规则**：
- 用 `shutil.copy2()` 复制模板再修改副本，避免重复图片
- **仅**修改 cell text 和 font 属性（Name, NameFarEast, Size, Bold）
- **不要**修改 row.Height/HeightRule/AllowBreakAcrossPages 等模板格式
- **不要**修改 cell 的 paragraph.KeepWithNext/SpaceBefore/SpaceAfter
- 逐个文件处理，避免 COM 连接问题；Word COM 不稳定已内置 `kill_word()` + 重试
- col2 必须是 `[(text, is_bold), ...]` 元组列表，不能传纯字符串

### Step 6: 验证一致性

生成全部教案后，逐一验证：

| 验证项 | 要求 |
|--------|------|
| 课题 | 教案课题 = 授课计划课题（空课题继承上一周） |
| 作业 | 教案作业 = 授课计划作业（含"（1学时）"后缀） |
| 日期 | 教案日期 = 校历对应周次的上课日期 |
| 周次 | 文件名周次 = 校历周次 = 授课计划周次 |

```python
import sys; sys.path.insert(0, r'.trae/skills/lesson-plan-generator/tools')
from word_lesson_tool import read_doc_tables, kill_word

tables = read_doc_tables('教案.doc')
t1 = tables[0]
date = t1[0][3]      # 日期周次
topic = t1[2][1]     # 课题
homework = t1[4][1]  # 作业
hours = t1[2][3]     # 课时
```

## Font and Format Requirements

| Element | Font | Size | Bold |
|---------|------|------|------|
| Body text | 宋体 | 12pt | No |
| Section headers | 宋体 | 12pt | Yes |
| Numbered sub-items | 宋体 | 12pt | Yes |
| Image captions | 宋体 | 10.5pt | No |
| Title "教案内容" | 黑体 | 22pt | Yes (in template) |

## Content Quality Requirements (内容质量要求)

教案内容必须充实、具体，不得空洞。

### 40min 环节（Row 4 - 获取信息）

教师讲授环节，内容必须充实、技术细节丰富：

1. **内容长度**：至少 600 字符
2. **技术细节**：必须包含具体参数、命令、操作步骤，不能只有概念性描述
   - 例如：写"速度控制"要列出 `rostopic pub /jetauto_controller/cmd_vel geometry_msgs/Twist`、参数范围（线速度 -0.7~0.7、角速度 -3.5~3.5）
   - 例如：写"配件安装"要列出螺丝规格（M3×6）、接口编号（②号接口、⑤号接口）
3. **分点结构**：用编号分点（1. 2. 3.）组织，每点有小标题（加粗）
4. **师生活动**：必须包含 `('教师活动：', True)` 和 `('学生活动：', True)`

### 80min 环节（Row 5 - 实施任务）

学生实操环节，依据教学计划当周课题，结合网盘资料或网络检索，按工学一体化特点设计：

1. **内容长度**：至少 400 字符，实操步骤具体可执行
2. **三段式结构**：
   - `1.安全教育` — 与本周内容相关的安全注意事项（2-3条）
   - `2.分组实操` — 详细操作步骤（检查清单、参数测试、数据记录等）
   - `3.填写工作单` — 工作单填写要求
3. **实操依据**：从 PDF 资料/网络检索中提炼操作步骤、检查项目、参数测试、数据记录。例如"速度控制"应提炼 `rostopic pub` 命令、线速度/角速度范围、轨迹行驶方法；"AiNova运动控制"应提炼转速参数（50/70/30rpm）、转弯逻辑、路线测试
4. **师生活动**：必须包含 `('教师活动：', True)`（巡回指导、示范、检查）和 `('学生活动：', True)`（按步骤实施、记录、协作）

### Row 3/6/7 也需包含师生活动

- **Row 3 (10min)**：安全教育 + 明确任务 + 教师活动 + 学生活动
- **Row 6 (10min)**：质量检验 + 教师活动 + 学生活动
- **Row 7 (20min)**：评价反馈 + 教师活动 + 学生活动

## Image Quality Requirements (配图要求)

PDF 提取的图片常含横幅图（页眉/分隔线，如 1267×80），在 Word 中显示高度仅约 18pt，无法看清。

**筛选规则**（`is_good_image()`）：高度 ≥ 200px、宽高比 ≤ 4.5、高宽比 ≤ 3.0

```python
from word_lesson_tool import pick_images, calc_width

imgs = pick_images(img_dir, 'JetAuto_配件安装教程', count=3, min_h=200)
for fp, w, h in imgs:
    width_pt = calc_width(w, h)  # 竖图150pt, 普通图180pt, 超宽图240pt
    images.append(('配件安装示意图', fp, width_pt))
```

**显示高度不得低于 50pt**。若某前缀所有图片均不合格，更换其他前缀或从网络下载替代配图。

生成后验证：
```python
for i in range(1, doc.InlineShapes.Count + 1):
    if doc.InlineShapes(i).Height < 50:
        print(f"图片 {i} 异常，需更换")
```

## Learning Objectives Style (工学一体化)

MUST start with "通过X节课的学习，学生能够：" + specific measurable verbs (not "学会"/"了解" alone).

## Common Issues

| Issue | Solution |
|-------|----------|
| Word COM crashes | `kill_word()` + retry (内置) |
| Table 2 on page 2 | Set rows 6-9 HeightRule=2, reduce height (内置) |
| "教案内容"留在第1页末尾 | process_week 已内置 `PageBreakBefore=True` + `KeepWithNext=True` |
| Duplicate images | Always `shutil.copy2()` fresh template |
| Chinese encoding | `# -*- coding: utf-8 -*-` in all scripts |
| table3_rows IndexError | 仅填充 row 3-7（row 1/2/8 不修改）；col2 必须是 `[(text, is_bold), ...]` 列表 |
| set_complex_cell 报错 | col2 不能传纯字符串，必须是 `[(text, is_bold), ...]` 列表 |
| 更新已有教案（非重新生成） | 用 win32com 直接打开文件，更新 Table 1/3 单元格后 Save；不要修改 row.Height 等模板格式 |
