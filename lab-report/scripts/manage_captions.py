#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
管理图表题注
- 插入题注行
- 格式：图X-Y 标题 / 表X-Y 标题
- 居中、五号字
"""

import win32com.client
import os
import shutil
import re


def insert_caption(doc, range_obj, caption_text, is_table=False):
    """
    在指定位置插入题注
    
    Args:
        doc: Word 文档对象
        range_obj: 插入位置
        caption_text: 题注文本（不含"图X-Y"前缀）
        is_table: 是否为表格题注
    """
    # 插入新段落
    range_obj.InsertParagraphAfter()
    caption_range = range_obj.Paragraphs(1).Next.Range
    
    # 设置题注文本
    prefix = "表" if is_table else "图"
    caption_range.Text = f"{prefix} {caption_text}\n"
    
    # 设置格式
    caption_range.Font.Size = 10.5  # 五号
    caption_range.Font.Name = "宋体"
    caption_range.ParagraphFormat.Alignment = 1  # 居中
    caption_range.ParagraphFormat.FirstLineIndent = 0  # 无缩进


def add_image_placeholder(doc, range_obj, chapter_num, img_num, caption_text):
    """
    添加图片占位符和题注
    
    Args:
        doc: Word 文档对象
        range_obj: 插入位置
        chapter_num: 章节号
        img_num: 图片序号
        caption_text: 题注内容
    """
    # 插入图片占位行
    range_obj.InsertParagraphAfter()
    placeholder_range = range_obj.Paragraphs(1).Next.Range
    placeholder_range.Text = "【此处插入图片】\n"
    placeholder_range.ParagraphFormat.Alignment = 1  # 居中
    
    # 插入题注
    insert_caption(doc, placeholder_range, f"{chapter_num}-{img_num} {caption_text}", is_table=False)


def add_table_caption(table, chapter_num, table_num, caption_text):
    """
    为表格添加题注
    
    Args:
        table: Word 表格对象
        chapter_num: 章节号
        table_num: 表格序号
        caption_text: 题注内容
    """
    # 在表格后插入段落
    table.Range.InsertParagraphAfter()
    caption_range = table.Range.Paragraphs(1).Next.Range
    
    # 设置题注文本
    caption_range.Text = f"表 {chapter_num}-{table_num} {caption_text}\n"
    
    # 设置格式
    caption_range.Font.Size = 10.5  # 五号
    caption_range.Font.Name = "宋体"
    caption_range.ParagraphFormat.Alignment = 1  # 居中
    caption_range.ParagraphFormat.FirstLineIndent = 0  # 无缩进


def process_captions(docx_path, output_path=None):
    """
    处理文档中的图表题注
    - 查找现有图表
    - 添加或修复题注格式
    
    Args:
        docx_path: 文档路径
        output_path: 输出路径
    """
    temp_dir = os.path.expanduser('~')
    temp_docx = os.path.join(temp_dir, 'temp_report.docx')
    shutil.copy2(docx_path, temp_docx)
    
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        
        doc = word.Documents.Open(temp_docx)
        
        # 处理表格题注
        table_count = {}
        for table in doc.Tables:
            # 尝试获取表格所在章节
            try:
                # 获取表格前的段落
                before_range = table.Range.Previous()
                if before_range:
                    text = before_range.Text
                    # 查找章节号（如"五、"）
                    match = re.search(r'(\d+)[、\.]', text)
                    if match:
                        chapter = int(match.group(1))
                    else:
                        chapter = 1
                else:
                    chapter = 1
            except:
                chapter = 1
            
            # 统计该章节的表格数
            if chapter not in table_count:
                table_count[chapter] = 0
            table_count[chapter] += 1
            
            # 检查是否已有题注
            try:
                after_para = table.Range.Next()
                if after_para and ('表' in after_para.Text or '图' in after_para.Text):
                    # 已有题注，修复格式
                    after_para.Range.Font.Size = 10.5
                    after_para.Range.Font.Name = "宋体"
                    after_para.Range.ParagraphFormat.Alignment = 1
                    after_para.Range.ParagraphFormat.FirstLineIndent = 0
                else:
                    # 添加题注
                    add_table_caption(table, chapter, table_count[chapter], "标题")
            except:
                add_table_caption(table, chapter, table_count[chapter], "标题")
        
        # 保存
        if output_path is None:
            output_path = docx_path
        
        doc.SaveAs(output_path)
        doc.Close()
        
        print(f"题注处理完成：{output_path}")
        return output_path
        
    except Exception as e:
        print(f"处理题注时出错：{e}")
        raise
    finally:
        word.Quit()
        if os.path.exists(temp_docx):
            os.remove(temp_docx)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("用法：python manage_captions.py <docx_path>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    process_captions(docx_path)
