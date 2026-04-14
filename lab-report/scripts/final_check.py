#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最终格式检查
- 封面信息
- 字体字号
- 缩进
- 题注格式
- 表格宽度
"""

import win32com.client
import os
import shutil


def final_check(docx_path):
    """
    对文档进行最终格式检查
    
    Args:
        docx_path: 文档路径
        
    Returns:
        检查报告字典
    """
    temp_dir = os.path.expanduser('~')
    temp_docx = os.path.join(temp_dir, 'temp_report.docx')
    shutil.copy2(docx_path, temp_docx)
    
    report = {
        'cover': {},
        'format': [],
        'tables': [],
        'warnings': []
    }
    
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        
        doc = word.Documents.Open(temp_docx)
        
        # 1. 检查封面信息
        cover_info = {'student_id': None, 'name': None, 'class': None, 'date': None}
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            
            if '学' in text and '号' in text and ('：' in text or ':' in text):
                cover_info['student_id'] = text.split('：')[-1].split(':')[-1].strip()
            elif '姓' in text and '名' in text and ('：' in text or ':' in text):
                cover_info['name'] = text.split('：')[-1].split(':')[-1].strip()
            elif '班' in text and '级' in text and ('：' in text or ':' in text):
                cover_info['class'] = text.split('：')[-1].split(':')[-1].strip()
            elif ('日' in text or '期' in text) and ('：' in text or ':' in text):
                cover_info['date'] = text.split('：')[-1].split(':')[-1].strip()
        
        report['cover'] = cover_info
        
        # 验证封面信息
        expected = {
            'student_id': '20232501306',
            'name': '欧劲聪',
            'class': '软件23-4'
        }
        
        for key, value in expected.items():
            if cover_info.get(key) != value:
                report['warnings'].append(f"封面{key}不匹配：期望 '{value}'，实际 '{cover_info.get(key)}'")
        
        # 2. 检查正文格式
        for i, para in enumerate(doc.Paragraphs[:50]):  # 检查前50段
            try:
                # 跳过表格
                if para.Range.Information(12):  # wdWithInTable
                    continue
                
                font_name = para.Range.Font.Name
                font_size = para.Range.Font.Size
                first_indent = para.FirstLineIndent
                
                # 检查字号（应为12pt小四）
                if font_size != 12 and font_size != 0:  # 0表示混合字号
                    report['format'].append({
                        'paragraph': i + 1,
                        'issue': f'字号异常：{font_size}pt（期望12pt）',
                        'text_preview': para.Range.Text[:30]
                    })
                
                # 检查缩进（应为约24磅）
                if first_indent < 20 or first_indent > 30:
                    if para.Range.Text.strip():  # 非空段落
                        report['format'].append({
                            'paragraph': i + 1,
                            'issue': f'首行缩进异常：{first_indent}磅（期望24磅）',
                            'text_preview': para.Range.Text[:30]
                        })
                        
            except Exception as e:
                pass
        
        # 3. 检查表格
        for i, table in enumerate(doc.Tables, 1):
            try:
                # 检查表格宽度
                total_width = sum(table.Columns(j).Width for j in range(1, table.Columns.Count + 1))
                page_width = doc.PageSetup.TextColumns.Width
                
                if total_width > page_width:
                    report['tables'].append({
                        'table': i,
                        'issue': f'表格宽度超出页面：{total_width:.1f}pt > {page_width:.1f}pt'
                    })
                
                # 检查表格字体
                cell = table.Cell(1, 1)
                if cell.Range.Font.Size != 10.5:
                    report['tables'].append({
                        'table': i,
                        'issue': f'表格字号异常：{cell.Range.Font.Size}pt（期望10.5pt五号）'
                    })
                    
            except Exception as e:
                pass
        
        doc.Close()
        
        # 打印报告
        print("=" * 50)
        print("最终格式检查报告")
        print("=" * 50)
        
        print("\n【封面信息】")
        for key, value in cover_info.items():
            status = "✓" if value else "✗"
            print(f"  {status} {key}: {value}")
        
        if report['warnings']:
            print("\n【警告】")
            for warning in report['warnings']:
                print(f"  ! {warning}")
        
        if report['format']:
            print(f"\n【格式问题】发现 {len(report['format'])} 处")
            for item in report['format'][:5]:  # 只显示前5个
                print(f"  第{item['paragraph']}段: {item['issue']}")
                print(f"    预览: {item['text_preview']}")
        else:
            print("\n【格式检查】✓ 未发现明显问题")
        
        if report['tables']:
            print(f"\n【表格问题】发现 {len(report['tables'])} 处")
            for item in report['tables']:
                print(f"  表{item['table']}: {item['issue']}")
        else:
            print("\n【表格检查】✓ 未发现明显问题")
        
        print("\n" + "=" * 50)
        
        return report
        
    except Exception as e:
        print(f"检查时出错：{e}")
        raise
    finally:
        word.Quit()
        if os.path.exists(temp_docx):
            os.remove(temp_docx)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("用法：python final_check.py <docx_path>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    final_check(docx_path)
