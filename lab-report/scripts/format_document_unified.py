#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一文档格式处理 - 使用 python-docx
替代原有的 win32com.client 依赖版本
- 正文：小四、宋体、Times New Roman（英文/数字）
- 首行缩进 2 字符
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


def is_english_or_number(char):
    """判断字符是否为英文或数字"""
    return bool(re.match(r'[a-zA-Z0-9]', char))


def format_document(docx_path: str, output_path: Optional[str] = None) -> str:
    """
    统一文档格式
    
    Args:
        docx_path: 文档路径
        output_path: 输出路径（可选）
        
    Returns:
        str: 输出文件路径
    """
    print(f"正在格式化文档: {docx_path}")
    
    doc = LabReportDocument(docx_path)
    
    if not doc.open():
        print("错误: 无法打开文档")
        return docx_path
    
    try:
        # 遍历所有段落并设置格式
        if hasattr(doc.doc, 'paragraphs'):
            for para in doc.doc.paragraphs:
                # 跳过表格内的段落
                try:
                    # python-docx 没有直接的表格内检测，这里简化处理
                    if not para.text.strip():
                        continue
                        
                    # 设置段落格式
                    success = doc.set_paragraph_format(
                        para,
                        font_name_cn="宋体",
                        font_name_en="Times New Roman", 
                        font_size=12,
                        first_line_indent=0.35  # 2字符缩进约0.35英寸
                    )
                    if not success:
                        print(f"警告: 段落格式设置失败: {para.text[:20]}...")
                        
                except Exception as e:
                    print(f"处理段落时出错: {e}")
                    continue
        
        # 处理表格格式
        if hasattr(doc.doc, 'tables'):
            for table in doc.doc.tables:
                try:
                    for row in table.rows:
                        for cell in row.cells:
                            # 设置表格内文字格式
                            for para in cell.paragraphs:
                                if para.text.strip():
                                    # 表格使用五号字
                                    success = doc.set_paragraph_format(
                                        para,
                                        font_name_cn="宋体",
                                        font_name_en="Times New Roman",
                                        font_size=10.5,  # 五号字
                                        first_line_indent=0  # 表格内不缩进
                                    )
                                    if not success:
                                        print(f"警告: 表格单元格格式设置失败")
                except Exception as e:
                    print(f"处理表格时出错: {e}")
                    continue
        
        # 保存文档
        output = output_path or docx_path
        if doc.close(output):
            print(f"文档格式调整完成: {output}")
            return output
        else:
            print("保存文档失败")
            return docx_path
            
    except Exception as e:
        print(f"格式调整时出错: {e}")
        doc.close()
        return docx_path


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='统一文档格式工具')
    parser.add_argument('docx_path', help='DOCX 文件路径')
    parser.add_argument('--output', '-o', help='输出文件路径（可选）')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.docx_path):
        print(f"错误: 文件不存在: {args.docx_path}")
        sys.exit(1)
    
    output_path = format_document(args.docx_path, args.output)
    sys.exit(0 if output_path != args.docx_path else 1)


if __name__ == '__main__':
    main()
