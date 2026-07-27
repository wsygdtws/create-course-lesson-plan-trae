---
name: "lesson-plan-generator"
description: "Generates lesson plan .doc files from templates. Invoke when user asks to create/write/generate lesson plans (教案), fill in teaching plan content, or modify existing lesson plans. MUST ask user about uncertain info before proceeding."
---

# Lesson Plan Generator (教案生成器)

This skill generates lesson plan .doc files based on provided templates, following the 工学一体化 (work-learning integration) style used in Chinese vocational schools.

## CRITICAL RULE: Ask Before Assuming

**Before generating any lesson plan, you MUST ask the user about the following uncertain information. NEVER make assumptions:**

1. **课时安排**: How many class periods per week? Which day/period? (e.g., "周二、周三下午7-10节，周四上午1-4节")
2. **班级信息**: Which classes? Multiple classes separated by "、"
3. **模板选择**: Which template to use? Is there a 机动周 (flexibility week) template?
4. **授课日期**: Confirm dates from the school calendar (校历)
5. **课题内容来源**: Where to find course content? (百度网盘 path, local files, etc.)
6. **排课方式**: Are all periods on the same day or spread across days? Which class on which day?

**If ANY key information is unclear, STOP and ask the user. Do NOT guess or assume.**

## Complete Workflow

### Step 1: Gather Information (收集信息)

Read and collect:
1. **授课计划 (Teaching Plan)** - .docx file containing:
   - Course name, class names, teacher name
   - Weekly hours, total hours
   - Weekly topics and main content
   - Homework assignments
2. **教案模板 (Lesson Plan Template)** - .doc file(s):
   - Regular template (教案模板.doc)
   - 机动周 template (教案模板（机动周）.doc) if exists
3. **校历 (School Calendar)** - .xlsx file:
   - Determine exact dates for each week
   - Identify holidays (国庆, 中秋, etc.)
4. **课题内容 (Course Content)** - from 百度网盘 or local files:
   - Download PDFs and extract text/images
   - Use content to fill "获取信息" section

### Step 2: Ask User About Uncertain Info

Use `AskUserQuestion` tool to confirm:
- 课时安排 (period schedule)
- 班级分配 (class assignment per day)
- Any other unclear information

### Step 3: Generate Lesson Plans

#### Template Structure (模板结构)

Each template has 3 tables:
- **Table 1** (10 rows × 4 cols): Header info (课程/日期/班级/课题/课时/教学方法/课后作业/学习目标/重点/难点/教学回顾/备注)
- **Table 2** (2 rows × 3 cols): Signature (授课教师/部长签名/主任签名/提交日期/审阅日期)
- **Table 3**: Lesson content (工学一体化 style)

#### Table 3 Structure (Regular Template - 8 rows):
| Row | Content |
|-----|---------|
| 1 | Headers (时间分配/教学活动内容/教学方法与手段/设计意图) - DO NOT MODIFY |
| 2 | 课前六件事 - DO NOT MODIFY |
| 3 | 10min: 安全教育 + 明确任务 |
| 4 | 40min: 获取信息 (main teaching content) |
| 5 | 80min: 实施任务 (hands-on practice) |
| 6 | 10min: 质量检验 |
| 7 | 20min: 评价反馈 |
| 8 | 课后四件事 - DO NOT MODIFY |

#### Table 3 Structure (机动周 Template - 7 rows):
Same structure but rows 3-6 are for 机动 content (教学准备/设备检查/备课研讨/场地布置).

### Step 4: Implementation (Python Script with win32com)

#### Critical Rules for Format Preservation:
1. **DO NOT** modify `row.Height`, `row.HeightRule`, or `row.AllowBreakAcrossPages`
2. **DO NOT** modify `paragraph.KeepWithNext`, `paragraph.SpaceBefore`, `paragraph.SpaceAfter`
3. **ONLY** modify cell text content and font properties (Name, NameFarEast, Size, Bold)
4. Use `shutil.copy2()` to copy template, then modify the copy

#### Helper Functions:

```python
import win32com.client
import shutil
import os
import subprocess
import time

def kill_word():
    """Kill Word process to handle COM errors"""
    subprocess.call('taskkill /F /IM WINWORD.EXE', shell=True, 
                    stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
    time.sleep(3)

def set_simple_cell(cell, text):
    """Set simple text in a cell, font=宋体 12pt"""
    rng = cell.Range
    rng.MoveEnd(1, -1)  # Exclude end-of-cell marker
    rng.Text = text
    rng.Font.Name = "宋体"
    rng.Font.NameFarEast = "宋体"
    rng.Font.Size = 12
    rng.Font.Bold = False

def set_complex_cell(cell, paragraphs):
    """Set multi-paragraph content with mixed bold/normal.
    paragraphs: list of (text, is_bold) tuples
    """
    rng = cell.Range
    rng.MoveEnd(1, -1)
    rng.Text = paragraphs[0][0]
    rng.Font.Name = "宋体"
    rng.Font.NameFarEast = "宋体"
    rng.Font.Size = 12
    rng.Font.Bold = paragraphs[0][1]
    
    for text, is_bold in paragraphs[1:]:
        rng.Collapse(0)
        rng.InsertParagraphAfter()
        rng.Collapse(0)
        rng.InsertAfter(text)
        rng.MoveStart(1, -len(text))
        rng.Font.Name = "宋体"
        rng.Font.NameFarEast = "宋体"
        rng.Font.Size = 12
        rng.Font.Bold = is_bold
        rng.Collapse(0)

def add_image_to_cell(cell, caption, img_path, width_pt=180):
    """Add caption + image at end of cell content"""
    rng = cell.Range
    rng.MoveEnd(1, -1)
    rng.Collapse(0)
    rng.InsertParagraphAfter()
    rng.Collapse(0)
    rng.InsertAfter(caption)
    rng.Font.Name = "宋体"
    rng.Font.NameFarEast = "宋体"
    rng.Font.Size = 10.5
    rng.Font.Bold = False
    rng.Collapse(0)
    rng.InsertParagraphAfter()
    rng.Collapse(0)
    shape = rng.InlineShapes.AddPicture(FileName=img_path)
    shape.LockAspectRatio = -1
    shape.Width = width_pt

def process_week(template_name, output_name, data, images=None, retry=0):
    """Process one week's lesson plan.
    
    data: dict with keys: course, date, class_name, topic, hours, methods, 
          tools, homework, homework_time, objectives, key_points, 
          difficulties, review, notes, table3_rows
    images: list of (caption, image_path, width_pt) tuples
    """
    base = os.getcwd()
    fpath_template = os.path.join(base, template_name)
    fpath_output = os.path.join(base, output_name)
    
    # Step 1: Copy template
    shutil.copy2(fpath_template, fpath_output)
    
    word = None
    doc = None
    saved = False
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(fpath_output)
        
        # Fill Table 1
        t1 = doc.Tables(1)
        set_simple_cell(t1.Cell(1, 2), data['course'])
        set_simple_cell(t1.Cell(1, 4), data['date'])
        set_simple_cell(t1.Cell(2, 2), data['class_name'])
        set_simple_cell(t1.Cell(3, 2), data['topic'])
        set_simple_cell(t1.Cell(3, 4), data['hours'])
        set_simple_cell(t1.Cell(4, 2), data['methods'])
        set_simple_cell(t1.Cell(4, 4), data['tools'])
        set_simple_cell(t1.Cell(5, 2), data['homework'])
        set_simple_cell(t1.Cell(5, 4), data['homework_time'])
        set_simple_cell(t1.Cell(6, 2), data['objectives'])
        set_simple_cell(t1.Cell(7, 2), data['key_points'])
        set_simple_cell(t1.Cell(8, 2), data['difficulties'])
        set_simple_cell(t1.Cell(9, 2), data.get('review', ''))
        set_simple_cell(t1.Cell(10, 2), data.get('notes', ''))
        
        # Fill Table 3
        t3 = doc.Tables(3)
        for row_num, row_data in data['table3_rows']:
            col1, col2, col3, col4 = row_data
            if col1 is not None:
                set_simple_cell(t3.Cell(row_num, 1), col1)
            if col2 is not None:
                set_complex_cell(t3.Cell(row_num, 2), col2)
            if col3 is not None:
                set_simple_cell(t3.Cell(row_num, 3), col3)
            if col4 is not None:
                set_simple_cell(t3.Cell(row_num, 4), col4)
        
        # Add images
        if images:
            cell = t3.Cell(4, 2)
            for caption, img_path, width_pt in images:
                if os.path.exists(img_path):
                    add_image_to_cell(cell, caption, img_path, width_pt)
        
        # Keep Table 1 + Table 2 on same page
        t2 = doc.Tables(2)
        t1_end = t1.Range.End
        t2_start = t2.Range.Start
        if t2_start > t1_end:
            between_range = doc.Range(t1_end, t2_start)
            for para in between_range.Paragraphs:
                para.KeepWithNext = True
                para.SpaceBefore = 0
                para.SpaceAfter = 0
        
        # Set rows 6-9 to Exactly height to prevent growing
        for ri in range(6, 10):
            t1.Rows(ri).HeightRule = 2  # wdRowHeightExactly
            t1.Rows(ri).Height = 75  # Slightly smaller than template's 85
        
        doc.Repaginate()
        
        # Check if Table 2 is on page 1
        t2_page = t2.Range.Information(3)
        if t2_page > 1:
            # Reduce row heights further
            for ri in range(6, 10):
                t1.Rows(ri).Height = 60
            doc.Repaginate()
        
        time.sleep(1)
        doc.Save()
        saved = True
        
    except Exception as e:
        if not saved and retry < 3:
            try:
                if doc: doc.Close(False)
            except: pass
            try:
                if word: word.Quit()
            except: pass
            kill_word()
            return process_week(template_name, output_name, data, images, retry + 1)
        raise
    finally:
        try:
            if doc: doc.Close(False)
        except: pass
        try:
            if word: word.Quit()
        except: pass
        kill_word()
    
    return saved
```

#### Important Notes:
- Word COM is unstable. Use retry mechanism with `kill_word()` between retries.
- Process files ONE AT A TIME to avoid COM connection issues.
- Use `doc.Save()` not `doc.SaveAs2()` to avoid format conversion.
- Process each week with separate script execution if COM keeps failing.

### Step 5: Extract Images from PDF

Use PyMuPDF (`fitz`) to extract images from downloaded PDF files:

```python
import fitz
import os

def extract_pdf_images(pdf_path, output_dir, prefix=""):
    """Extract all images from a PDF file."""
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(doc.page_count):
        page = doc[page_num]
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            if pix.n - pix.alpha > 3:  # CMYK
                pix = fitz.Pixmap(fitz.csRGB, pix)
            img_name = f"{prefix}_p{page_num+1}_{img_index}.png"
            img_path = os.path.join(output_dir, img_name)
            pix.save(img_path)
            images.append((img_name, pix.width, pix.height, os.path.getsize(img_path)))
            pix = None
    doc.close()
    return images
```

### Step 6: Download from Baidu Netdisk

Use MCP tools (`mcp_baidu-netdisk`) or Baidu PCS API to download course content:

```python
import requests
import urllib.parse

def download_baidu_file(access_token, remote_path, local_path):
    """Download a file from Baidu Netdisk via PCS API."""
    encoded_path = urllib.parse.quote(remote_path, safe='')
    url = f"https://d.pcs.baidu.com/rest/2.0/pcs/file?method=download&access_token={access_token}&path={encoded_path}"
    resp = requests.get(url, stream=True, timeout=60)
    if resp.status_code == 200:
        with open(local_path, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        return True
    return False
```

### Step 7: Verify Generated Files

After generation, verify:
1. Table 1 has correct row count and content
2. Table 2 (signature) is on page 1 (same page as Table 1)
3. Table 3 has correct content in each phase
4. Images are properly inserted
5. Date format includes correct period numbers
6. Learning objectives start with "学生能够" (工学一体化 style)

## Learning Objectives Style (工学一体化)

Learning objectives MUST:
1. Start with "通过X节课的学习，学生能够："
2. Use "能够" + specific actionable verb (not "学会"/"了解"/"掌握" alone)
3. Be specific and measurable (e.g., "对照实物指认..." not "了解结构")
4. Include 2-3 objectives per lesson

Example:
```
通过4节课的学习，学生能够：
1.对照AiNova和JetAuto实物，指出感知、决策、执行和通信四大系统对应的主要硬件模块；
2.按照实训室安全规范，正确完成设备取放和通电检查；
3.填写设备认识记录表，准确记录两种实训平台的主要模块名称及功能。
```

## Font and Format Requirements

| Element | Font | Size | Bold |
|---------|------|------|------|
| Body text | 宋体 | 12pt | No |
| Section headers (安全教育/明确任务/获取信息 etc.) | 宋体 | 12pt | Yes |
| Numbered sub-items (1.xxx/2.xxx) | 宋体 | 12pt | Yes |
| Sub-item content | 宋体 | 12pt | No |
| Image captions | 宋体 | 10.5pt | No |
| Table headers | 宋体 | 12pt | Yes |
| Title "教案内容" | 黑体 | 22pt | Yes (already in template) |

## Date Format

Date string format for Table 1 Cell(1,4):
```
2026年9月8日、9月9日、9月10日 第2周 周二、周三7-10节、周四1-4节
```
- Multiple dates separated by "、"
- Multiple classes on different days
- Period numbers specified (e.g., "7-10节" for afternoon, "1-4节" for morning)

## Common Issues and Solutions

1. **Word COM crashes (RPC server unavailable)**: Use `kill_word()` + retry mechanism
2. **Table 2 on page 2**: Set rows 6-9 HeightRule=2 (Exactly) and reduce height
3. **Duplicate images**: Always copy template fresh with `shutil.copy2()` before modifying
4. **Content overflow in cells**: Keep text concise, use bullet points
5. **Chinese encoding issues**: Use `# -*- coding: utf-8 -*-` in all Python scripts
6. **PDF text extraction**: Use PyMuPDF (`fitz`) with `page.get_text()`
7. **Image extraction from PDF**: Use `page.get_images(full=True)` + `fitz.Pixmap`
