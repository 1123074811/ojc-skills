#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理多余空行和分页 - 使用 python-docx
替代原有的 win32com.client 依赖版本
- 标题间最多1个空行
- 最大标题前分页
"""

import os
import sys
import re
from typing import Optional

# 添加脚本目录到路径
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from docx_handler import LabReportDocument


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


def cleanup_spacing(docx_path: str, output_path: Optional[str] = None) -> str:
    """
    清理文档中的多余空行和分页
    
    Args:
        docx_path: 文档路径
        output_path: 输出路径（可选）
        
    Returns:
        str: 输出文件路径
    """
    print(f"正在清理文档格式: {docx_path}")
    
    doc = LabReportDocument(docx_path)
    
    if not doc.open():
        print("错误: 无法打开文档")
        return docx_path
    
    try:
        # 使用内置的清理空行功能
        success = doc.cleanup_empty_paragraphs(max_empty=1)
        if not success:
            print("警告: 空行清理可能不完整")
        
        # 在最大标题前添加分页符
        if hasattr(doc.doc, 'paragraphs'):
            for para in doc.doc.paragraphs:
                text = para.text.strip()
                if is_main_title(text):
                    # 在标题前添加分页符
                    success = doc.add_page_break_before_heading(text)
                    if not success:
                        print(f"警告: 分页符添加失败: {text}")
        
        # 保存文档
        output = output_path or docx_path
        if doc.close(output):
            print(f"空行清理完成: {output}")
            return output
        else:
            print("保存文档失败")
            return docx_path
            
    except Exception as e:
        print(f"清理空行时出错: {e}")
        doc.close()
        return docx_path


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='清理文档格式工具')
    parser.add_argument('docx_path', help='DOCX 文件路径')
    parser.add_argument('--output', '-o', help='输出文件路径（可选）')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.docx_path):
        print(f"错误: 文件不存在: {args.docx_path}")
        sys.exit(1)
    
    output_path = cleanup_spacing(args.docx_path, args.output)
    sys.exit(0 if output_path != args.docx_path else 1)


if __name__ == '__main__':
    main()
