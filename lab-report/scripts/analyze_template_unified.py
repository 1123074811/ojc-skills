#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析实验报告模板 - 使用 python-docx
替代原有的 win32com.client 依赖版本
提取关键信息：报告名称、章节结构、表格位置、封面字段
"""

import os
import sys
import re
from typing import Dict, List, Any

# 添加脚本目录到路径
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from docx_handler import LabReportDocument


def analyze_template(template_path: str) -> Dict[str, Any]:
    """
    分析模板文档，提取关键信息
    
    Args:
        template_path: 模板文件路径
        
    Returns:
        dict: 包含提取的信息
        {
            'report_name': '实验报告名称',
            'has_cover': True/False,
            'sections': ['一、实验目的', '二、实验内容', ...],
            'table_positions': [{'section': '五、实验结果', 'headers': [...]}],
            'cover_fields': {'学号': '', '姓名': '', '班级': '', '日期': ''}
        }
    """
    print(f"正在分析模板: {template_path}")
    
    doc = LabReportDocument(template_path)
    
    if not doc.open():
        print("错误: 无法打开模板文档")
        return {}
    
    try:
        result = {
            'report_name': '',
            'has_cover': False,
            'sections': [],
            'table_positions': [],
            'cover_fields': {}
        }
        
        # 提取所有文本
        full_text = ""
        if hasattr(doc.doc, 'paragraphs'):
            for para in doc.doc.paragraphs:
                full_text += para.text + "\n"
        
        # 1. 识别实验报告名称
        report_patterns = [
            r'《(.+?)》实验报告',
            r'实验报告[（(]([一二三四五六七八九十]+)[)）]',
            r'实验报告\s*([一二三四五六七八九十]+)',
            r'实验[一二三四五六七八九十]',
        ]
        
        for pattern in report_patterns:
            match = re.search(pattern, full_text)
            if match:
                result['report_name'] = match.group(0)
                break
        
        # 2. 分析段落结构
        if hasattr(doc.doc, 'paragraphs'):
            for para in doc.doc.paragraphs:
                text = para.text.strip()
                if not text:
                    continue
                
                # 识别章节标题（如"一、实验目的"）
                section_match = re.match(r'^([一二三四五六七八九十]+)[、.．]\s*(.+)', text)
                if section_match:
                    section_title = text
                    result['sections'].append(section_title)
                    
                    # 检查是否是表格插入位置（实验结果/测试用例相关章节）
                    if any(keyword in text for keyword in ['实验结果', '测试用例', '测试结果', '实验数据']):
                        result['table_positions'].append({
                            'section': section_title,
                            'headers': None  # 后续根据内容推断
                        })
                
                # 识别封面字段
                if any(keyword in text for keyword in ['学号', '姓名', '班级', '日期', '专业']):
                    result['has_cover'] = True
                    for field in ['学号', '姓名', '班级', '日期', '专业', '课程']:
                        if field in text:
                            result['cover_fields'][field] = text
        
        # 3. 推断表格表头（根据实验类型）
        if result['table_positions']:
            # 默认使用测试用例表头
            result['table_positions'][0]['headers'] = [
                "测试用例编号", "测试项目", "测试子项", "测试用例描述",
                "前置条件", "测试输入", "操作步骤", "预期结果",
                "实际结果", "测试结果"
            ]
        
        doc.close()
        
        print(f"模板分析完成：")
        print(f"  - 报告名称：{result['report_name'] or '未识别'}")
        print(f"  - 章节数：{len(result['sections'])}")
        print(f"  - 表格位置：{len(result['table_positions'])}")
        print(f"  - 封面字段：{list(result['cover_fields'].keys())}")
        
        return result
        
    except Exception as e:
        print(f"分析模板时出错: {e}")
        doc.close()
        return {}


def extract_report_name_from_filename(filename: str) -> str:
    """
    从文件名中提取实验报告名称
    
    Args:
        filename: 文件名
        
    Returns:
        str: 实验报告名称
    """
    # 移除扩展名
    name = os.path.splitext(filename)[0]
    
    # 常见模式匹配
    patterns = [
        r'实验报告[（(]([一二三四五六七八九十]+)[)）]',
        r'实验([一二三四五六七八九十]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, name)
        if match:
            return match.group(0)
    
    # 如果匹配不到，返回清理后的文件名
    # 移除常见的冗余文字
    clean_name = re.sub(r'[《》]', '', name)
    clean_name = re.sub(r'[（(].*?[)）]', '', clean_name)
    clean_name = re.sub(r'\s+', ' ', clean_name).strip()
    
    return clean_name


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='分析实验报告模板工具')
    parser.add_argument('template_path', help='模板文件路径')
    parser.add_argument('--output', '-o', help='输出分析结果到JSON文件')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.template_path):
        print(f"错误: 模板文件不存在: {args.template_path}")
        sys.exit(1)
    
    result = analyze_template(args.template_path)
    
    if args.output:
        try:
            import json
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"分析结果已保存到: {args.output}")
        except Exception as e:
            print(f"保存分析结果失败: {e}")
    
    print("\n分析结果：")
    print(result)


if __name__ == '__main__':
    main()
