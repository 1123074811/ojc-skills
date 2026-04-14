"""
填写封面信息
使用 docx_handler 模块处理 DOCX 文件
"""

import os
import sys
from datetime import datetime

# 添加脚本目录到路径
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from docx_handler import LabReportDocument


def fill_cover_info(docx_path: str, student_id: str, name: str, 
                    class_name: str, date: str = None) -> bool:
    """
    填写实验报告封面信息
    
    Args:
        docx_path: DOCX 文件路径
        student_id: 学号
        name: 姓名
        class_name: 班级
        date: 日期（可选，默认当天）
    
    Returns:
        bool: 是否成功
    """
    if date is None:
        date = datetime.now().strftime("%Y年%m月%d日")
    
    print(f"正在填写封面信息: {docx_path}")
    print(f"  学号: {student_id}")
    print(f"  姓名: {name}")
    print(f"  班级: {class_name}")
    print(f"  日期: {date}")
    
    doc = LabReportDocument(docx_path)
    
    if not doc.open():
        print("错误: 无法打开文档")
        return False
    
    try:
        # 填写封面信息
        success = doc.fill_cover_info(student_id, name, class_name, date)
        
        if success:
            print("封面信息填写成功")
        else:
            print("警告: 封面信息填写可能不完整")
        
        # 关闭文档（保存）
        return doc.close()
        
    except Exception as e:
        print(f"填写封面时出错: {e}")
        doc.close()
        return False


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='填写实验报告封面信息')
    parser.add_argument('docx_path', help='DOCX 文件路径')
    parser.add_argument('--student-id', '-s', required=True, help='学号')
    parser.add_argument('--name', '-n', required=True, help='姓名')
    parser.add_argument('--class', '-c', dest='class_name', required=True, help='班级')
    parser.add_argument('--date', '-d', help='日期（格式：YYYY年MM月DD日）')
    
    args = parser.parse_args()
    
    success = fill_cover_info(
        args.docx_path,
        args.student_id,
        args.name,
        args.class_name,
        args.date
    )
    
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
