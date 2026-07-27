# -*- coding: utf-8 -*-
"""Word COM 工具：读取/填充教案模板，含 kill_word 重试机制"""
import win32com.client
import subprocess
import time
import os

def kill_word():
    """Kill Word process to handle COM errors"""
    subprocess.call('taskkill /F /IM WINWORD.EXE', shell=True,
                    stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
    time.sleep(3)

def read_doc_tables(filepath, retry=0):
    """读取 .doc 文件中所有表格内容，返回 list[list[list[str]]]
    每个 table -> rows -> cells (text)
    """
    word = None; doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(filepath))
        tables = []
        for i in range(1, doc.Tables.Count + 1):
            table = doc.Tables(i)
            rows = []
            for r in range(1, table.Rows.Count + 1):
                row = []
                for c in range(1, table.Columns.Count + 1):
                    try:
                        cell = table.Cell(r, c)
                        text = cell.Range.Text.replace('\r\x07', '').replace('\x07', '')
                        row.append(text)
                    except:
                        row.append("")
                rows.append(row)
            tables.append(rows)
        return tables
    except Exception as e:
        if retry < 3:
            try:
                if doc: doc.Close(False)
            except: pass
            try:
                if word: word.Quit()
            except: pass
            kill_word()
            return read_doc_tables(filepath, retry + 1)
        raise
    finally:
        try:
            if doc: doc.Close(False)
        except: pass
        try:
            if word: word.Quit()
        except: pass

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
    shape = rng.InlineShapes.AddPicture(FileName=os.path.abspath(img_path))
    shape.LockAspectRatio = -1
    shape.Width = width_pt

def process_week(template_name, output_name, data, images=None, retry=0):
    """Process one week's lesson plan.
    data: dict with keys: course, date, class_name, topic, hours, methods,
          tools, homework, homework_time, objectives, key_points,
          difficulties, review, notes, table3_rows
    images: list of (caption, image_path, width_pt) tuples
    Returns True if saved successfully.
    """
    import shutil
    base = os.getcwd()
    fpath_template = os.path.join(base, template_name)
    fpath_output = os.path.join(base, output_name)
    shutil.copy2(fpath_template, fpath_output)

    word = None; doc = None; saved = False
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
            t1.Rows(ri).Height = 75

        doc.Repaginate()
        t2_page = t2.Range.Information(3)
        if t2_page > 1:
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

if __name__ == '__main__':
    # 用法示例
    import sys
    if len(sys.argv) < 2:
        print("用法: python word_lesson_tool.py <doc文件路径>  读取表格内容")
        print("     或作为模块导入: from word_lesson_tool import process_week, read_doc_tables")
        sys.exit(1)
    tables = read_doc_tables(sys.argv[1])
    for i, t in enumerate(tables):
        print(f"\n===== Table {i+1} =====")
        for row in t:
            print(" | ".join(row))
