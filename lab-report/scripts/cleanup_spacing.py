#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理多余空行和分页
- 标题间最多1个空行
- 最大标题前分页
"""

import win32com.client
import os
import shutil
import re


def is_title(text):
    """判断是否为标题"""
    text = text.strip()
    # 匹配模式：一、  1.  1.1  第X章  等
    patterns = [
        r'^[一二三四五六七八九十]+[、\.]',  # 一、二、
        r'^\d+[、\.]',  # 1、1.
        r'^\d+\.\d+',  # 1.1
        r'^第[一二三四五六七八九十\d]+章',  # 第一章
        r'^[（(]\d+[)）]',  # (1)（1）
    ]
    for pattern in patterns:
        if re.match(pattern, text):
            return True
    return False


def is_main_title(text):
    """判断是否为最大标题（一级标题）"""
    text = text.strip()
    # 匹配一级标题：一、  1.  第X章
    patterns = [
        r'^[一二三四五六七八九十]+[、\.].{2,20}$',  # 一、实验目的
        r'^\d+[、\.].{2,20}$',  # 1. 实验目的
        r'^第[一二三四五六七八九十\d]+章',  # 第一章
    ]
    for pattern in patterns:
        if re.match(pattern, text):
            return True
    return False


def cleanup_spacing(docx_path, output_path=None):
    """
    清理文档中的多余空行和分页
    
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
        
        # 第一遍：删除多余空行
        paras_to_delete = []
        empty_count = 0
        prev_was_title = False
        
        for i, para in enumerate(doc.Paragraphs, 1):
            text = para.Range.Text.strip()
            
            if text == '':
                empty_count += 1
                # 如果连续空行超过1个，标记删除
                if empty_count > 1:
                    paras_to_delete.append(para)
            else:
                if is_title(text):
                    # 如果当前是标题，且前面有空行，保留1个
                    if empty_count > 1:
                        # 删除多余的空行（保留最后1个）
                        for j in range(len(paras_to_delete) - empty_count + 1, len(paras_to_delete)):
                            if j >= 0:
                                paras_to_delete[j].Range.Delete()
                    empty_count = 0
                    prev_was_title = True
                else:
                    empty_count = 0
                    prev_was_title = False
        
        # 第二遍：在最大标题前插入分页符
        for i, para in enumerate(doc.Paragraphs, 1):
            text = para.Range.Text.strip()
            
            if is_main_title(text):
                # 检查前一段是否已有分页符
                try:
                    prev_para = para.Previous()
                    if prev_para:
                        # 在标题前插入分页
                        para.Range.InsertBreak(7)  # wdPageBreak = 7
                except:
                    pass
        
        # 保存
        if output_path is None:
            output_path = docx_path
        
        doc.SaveAs(output_path)
        doc.Close()
        
        print(f"空行清理完成：{output_path}")
        return output_path
        
    except Exception as e:
        print(f"清理空行时出错：{e}")
        raise
    finally:
        word.Quit()
        if os.path.exists(temp_docx):
            os.remove(temp_docx)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("用法：python cleanup_spacing.py <docx_path>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    cleanup_spacing(docx_path)
