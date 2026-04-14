#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析实验报告模板，提取关键信息
"""

import win32com.client
import os
import shutil
import re


def analyze_template(template_path):
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
    temp_dir = os.path.expanduser('~')
    temp_docx = os.path.join(temp_dir, 'temp_template.docx')
    
    # 处理中文路径
    shutil.copy2(template_path, temp_docx)
    
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        
        doc = word.Documents.Open(temp_docx)
        
        result = {
            'report_name': '',
            'has_cover': False,
            'sections': [],
            'table_positions': [],
            'cover_fields': {}
        }
        
        # 提取所有文本
        full_text = doc.Content.Text
        
        # 1. 识别实验报告名称
        # 常见模式：文件名中的"实验报告（X）"或文档标题
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
        current_section = None
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue
            
            # 识别章节标题（如"一、实验目的"）
            section_match = re.match(r'^([一二三四五六七八九十]+)[、.．]\s*(.+)', text)
            if section_match:
                section_title = text
                result['sections'].append(section_title)
                current_section = section_title
                
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
        
        doc.Close()
        
        print(f"模板分析完成：")
        print(f"  - 报告名称：{result['report_name'] or '未识别'}")
        print(f"  - 章节数：{len(result['sections'])}")
        print(f"  - 表格位置：{len(result['table_positions'])}")
        print(f"  - 封面字段：{list(result['cover_fields'].keys())}")
        
        return result
        
    except Exception as e:
        print(f"分析模板时出错：{e}")
        raise
    finally:
        word.Quit()
        if os.path.exists(temp_docx):
            os.remove(temp_docx)


def extract_report_name_from_filename(filename):
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


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("用法：python analyze_template.py <template_path>")
        sys.exit(1)
    
    template_path = sys.argv[1]
    result = analyze_template(template_path)
    print("\n分析结果：")
    print(result)
