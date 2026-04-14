#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
实验报告处理统一工作流
整合所有功能，提供完整的实验报告自动化处理流程
"""

import os
import sys
import json
import argparse
from datetime import datetime
from typing import Dict, List, Optional, Any

# 添加脚本目录到路径
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from docx_handler import LabReportDocument
from analyze_template_unified import analyze_template, extract_report_name_from_filename
from insert_test_cases import insert_test_case_table, generate_sample_test_cases
from rewrite_aigc import rewrite_text, load_api_config
from format_document_unified import format_document
from cleanup_spacing_unified import cleanup_spacing


def process_lab_report(template_path: str, student_info: Dict[str, str], 
                      options: Optional[Dict[str, Any]] = None) -> str:
    """
    处理实验报告的完整工作流
    
    Args:
        template_path: 模板文件路径
        student_info: 学生信息 {'student_id': '', 'name': '', 'class_name': ''}
        options: 处理选项 {
            'output_path': '',  # 输出路径（可选）
            'use_aigc': False,  # 是否使用AIGC改写
            'api_config_path': '',  # API配置文件路径
            'insert_test_cases': False,  # 是否插入测试用例
            'test_case_type': 'default',  # 测试用例类型
            'format_document': True,  # 是否格式化文档
            'cleanup_spacing': True,  # 是否清理空行
        }
    
    Returns:
        str: 输出文件路径
    """
    if options is None:
        options = {}
    
    print("=" * 60)
    print("实验报告自动化处理开始")
    print("=" * 60)
    
    # 验证输入
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"模板文件不存在: {template_path}")
    
    required_fields = ['student_id', 'name', 'class_name']
    for field in required_fields:
        if not student_info.get(field):
            raise ValueError(f"缺少必填信息: {field}")
    
    # 1. 分析模板
    print("\n1. 分析模板结构...")
    template_info = analyze_template(template_path)
    
    # 生成输出文件名
    report_name = template_info.get('report_name', extract_report_name_from_filename(os.path.basename(template_path)))
    output_filename = f"{student_info['name']}-{student_info['class_name']}-{student_info['student_id']}-{report_name}.docx"
    
    if options.get('output_path'):
        output_path = options['output_path']
    else:
        output_path = os.path.join(os.path.dirname(template_path), output_filename)
    
    # 复制模板到输出路径
    import shutil
    shutil.copy2(template_path, output_path)
    print(f"已创建工作副本: {output_path}")
    
    # 2. 填写封面信息
    print("\n2. 填写封面信息...")
    doc = LabReportDocument(output_path)
    if doc.open():
        success = doc.fill_cover_info(
            student_id=student_info['student_id'],
            name=student_info['name'],
            class_name=student_info['class_name'],
            date=datetime.now().strftime("%Y年%m月%d日")
        )
        if success:
            print("封面信息填写完成")
        else:
            print("警告: 封面信息填写可能不完整")
        doc.close()
    
    # 3. 插入测试用例表格（可选）
    if options.get('insert_test_cases', False):
        print("\n3. 插入测试用例表格...")
        test_cases = generate_sample_test_cases()
        
        # 根据类型选择表头
        from insert_test_cases import DEFAULT_HEADERS, SOFTWARE_TEST_HEADERS, PERFORMANCE_TEST_HEADERS
        test_case_type = options.get('test_case_type', 'default')
        
        if test_case_type == 'software':
            headers = SOFTWARE_TEST_HEADERS
        elif test_case_type == 'performance':
            headers = PERFORMANCE_TEST_HEADERS
        else:
            headers = DEFAULT_HEADERS
        
        success = insert_test_case_table(output_path, test_cases, headers)
        if success:
            print("测试用例表格插入完成")
        else:
            print("警告: 测试用例表格插入可能不完整")
    
    # 4. AIGC降重改写（可选）
    if options.get('use_aigc', False):
        print("\n4. 执行AIGC降重改写...")
        api_config_path = options.get('api_config_path', 'aigc_config.json')
        api_config = load_api_config(api_config_path)
        
        if api_config:
            # 这里可以扩展为对文档特定段落进行改写
            print("AIGC配置已加载，但段落级改写功能需要进一步开发")
            print("当前版本仅提供提示词模式")
        else:
            print("未找到API配置，跳过AIGC改写")
    
    # 5. 格式化文档
    if options.get('format_document', True):
        print("\n5. 统一文档格式...")
        format_document(output_path)
    
    # 6. 清理空行和分页
    if options.get('cleanup_spacing', True):
        print("\n6. 清理多余空行和分页...")
        cleanup_spacing(output_path)
    
    print("\n" + "=" * 60)
    print(f"实验报告处理完成: {output_path}")
    print("=" * 60)
    
    return output_path


def create_sample_config():
    """创建示例配置文件"""
    from rewrite_aigc import create_sample_config
    create_sample_config()


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description='实验报告自动化处理工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 基本用法
  python lab_report_workflow.py template.docx --student-id 20232501306 --name 张三 --class 软件23-4
  
  # 完整功能
  python lab_report_workflow.py template.docx \\
    --student-id 20232501306 --name 张三 --class 软件23-4 \\
    --output result.docx --insert-test-cases --use-aigc --api-config aigc.json
        """
    )
    
    # 必填参数
    parser.add_argument('template', help='模板文件路径')
    parser.add_argument('--student-id', required=True, help='学号')
    parser.add_argument('--name', required=True, help='姓名')
    parser.add_argument('--class', dest='class_name', required=True, help='班级')
    
    # 可选参数
    parser.add_argument('--output', '-o', help='输出文件路径')
    parser.add_argument('--insert-test-cases', action='store_true', help='插入测试用例表格')
    parser.add_argument('--test-case-type', choices=['default', 'software', 'performance'], 
                       default='default', help='测试用例类型')
    parser.add_argument('--use-aigc', action='store_true', help='使用AIGC降重改写')
    parser.add_argument('--api-config', help='AIGC API配置文件路径')
    parser.add_argument('--no-format', action='store_true', help='跳过文档格式化')
    parser.add_argument('--no-cleanup', action='store_true', help='跳过空行清理')
    parser.add_argument('--create-config', action='store_true', help='创建示例配置文件')
    
    args = parser.parse_args()
    
    # 创建配置文件
    if args.create_config:
        create_sample_config()
        return
    
    # 验证模板文件
    if not os.path.exists(args.template):
        print(f"错误: 模板文件不存在: {args.template}")
        sys.exit(1)
    
    # 准备学生信息
    student_info = {
        'student_id': args.student_id,
        'name': args.name,
        'class_name': args.class_name
    }
    
    # 准备处理选项
    options = {
        'output_path': args.output,
        'insert_test_cases': args.insert_test_cases,
        'test_case_type': args.test_case_type,
        'use_aigc': args.use_aigc,
        'api_config_path': args.api_config,
        'format_document': not args.no_format,
        'cleanup_spacing': not args.no_cleanup,
    }
    
    try:
        # 执行处理
        output_path = process_lab_report(args.template, student_info, options)
        print(f"\n✅ 处理成功! 输出文件: {output_path}")
        
    except Exception as e:
        print(f"\n❌ 处理失败: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
