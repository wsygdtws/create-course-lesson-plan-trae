# -*- coding: utf-8 -*-
"""PDF 文本与图片提取工具"""
import fitz
import os

def extract_text(pdf_path, save_to=None):
    """提取 PDF 全部文本，可保存到文件。
    返回文本字符串。
    """
    doc = fitz.open(pdf_path)
    parts = []
    for page_num in range(doc.page_count):
        page = doc[page_num]
        text = page.get_text()
        parts.append(f"--- Page {page_num+1} ---\n{text}")
    doc.close()
    full = "\n".join(parts)
    if save_to:
        with open(save_to, 'w', encoding='utf-8') as f:
            f.write(full)
    return full

def extract_images(pdf_path, output_dir, prefix=""):
    """提取 PDF 中所有图片，保存为 PNG。
    prefix: 文件名前缀，如 "1_1_1_认识AiNova"
    返回图片路径列表 [(img_path, width, height, size_bytes), ...]
    """
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(doc.page_count):
        page = doc[page_num]
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            if pix.n - pix.alpha > 3:  # CMYK -> RGB
                pix = fitz.Pixmap(fitz.csRGB, pix)
            # 清理 prefix 中的特殊字符
            clean_prefix = prefix.replace("/", "_").replace("\\", "_")
            img_name = f"{clean_prefix}_p{page_num+1}_{img_index}.png"
            img_path = os.path.join(output_dir, img_name)
            pix.save(img_path)
            images.append((img_path, pix.width, pix.height, os.path.getsize(img_path)))
            pix = None
    doc.close()
    return images

def batch_extract(pdf_dir, output_dir, prefix_map=None):
    """批量提取目录下所有 PDF 的文本和图片。
    prefix_map: 可选 dict {filename: prefix}，未指定则用文件名（去扩展名）
    """
    results = {}
    for fname in sorted(os.listdir(pdf_dir)):
        if not fname.lower().endswith('.pdf'):
            continue
        pdf_path = os.path.join(pdf_dir, fname)
        prefix = prefix_map.get(fname, os.path.splitext(fname)[0]) if prefix_map else os.path.splitext(fname)[0]
        imgs = extract_images(pdf_path, output_dir, prefix)
        results[fname] = {
            'text': extract_text(pdf_path),
            'images': imgs
        }
        print(f"  {fname}: {len(imgs)} images")
    return results

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 3:
        print("用法: python pdf_tool.py <pdf路径> <输出目录> [前缀]")
        print("     提取图片到输出目录")
        sys.exit(1)
    pdf_path = sys.argv[1]
    out_dir = sys.argv[2]
    prefix = sys.argv[3] if len(sys.argv) > 3 else ""
    imgs = extract_images(pdf_path, out_dir, prefix)
    print(f"提取完成: {len(imgs)} 张图片 -> {out_dir}")
    # 同时提取文本
    text = extract_text(pdf_path)
    print(f"文本长度: {len(text)} 字符")
