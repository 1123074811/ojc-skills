#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一文档格式
- 正文：小四、宋体、Times New Roman（英文/数字）
- 首行缩进 2 字符
"""

import win32com.client
import os
import shutil
import re


def set_font_for_range(range_obj, is_english=False):
    """为范围设置字体"""
    range_obj.Font.Name = "Times New Roman" if is_english else "宋体"
    range_obj.Font.Size = 12  # 小四


def is_english_or_number(char):
    """判断字符是否为英文或数字"""
    return bool(re.match(r'[a-zA-Z0-9]', char))


def format_document(docx_path, output_path=None):
    """
    统一文档格式
    
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
        
        # 遍历所有段落
        for para in doc.Paragraphs:
            # 检查是否在表格内
            try:
                # 如果在表格内，跳过（表格单独处理）
                if para.Range.Information(12):  # wdWithInTable = 12
                    continue
            except:
                pass
            
            # 设置段落格式
            para.Range.Font.Size = 12  # 小四
            para.Range.Font.Name = "宋体"
            para.Range.Font.NameFarEast = "宋体"
            
            # 首行缩进 2 字符（约 24 磅）
            para.FirstLineIndent = 24
            
            # 处理英文和数字（设置为 Times New Roman）
            text = para.Range.Text
            for i, char in enumerate(text):
                if is_english_or_number(char):
                    # 设置单个字符为 Times New Roman
                    try:
                        char_range = doc.Range(para.Range.Start + i, para.Range.Start + i + 1)
                        char_range.Font.Name = "Times New Roman"
                    except:
                        pass
        
        # 处理表格
        for table in doc.Tables:
            for row in range(1, table.Rows.Count + 1):
                for col in range(1, table.Columns.Count + 1):
                    cell = table.Cell(row, col)
                    cell.Range.Font.Size = 10.5  # 五号
                    cell.Range.Font.Name = "宋体"
                    
                    # 处理单元格内的英文和数字
                    text = cell.Range.Text
                    for i, char in enumerate(text):
                        if is_english_or_number(char):
                            try:
                                char_range = doc.Range(cell.Range.Start + i, cell.Range.Start + i + 1)
                                char_range.Font.Name = "Times New Roman"
                            except:
                                pass
        
        # 保存
        if output_path is None:
            output_path = docx_path
        
        doc.SaveAs(output_path)
        doc.Close()
        
        print(f"文档格式调整完成：{output_path}")
        return output_path
        
    except Exception as e:
        print(f"格式调整时出错：{e}")
        raise
    finally:
        word.Quit()
        if os.path.exists(temp_docx):
            os.remove(temp_docx)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("用法：python format_document.py <docx_path>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    format_document(docx_path)
