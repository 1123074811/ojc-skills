"""
插入测试用例表格
使用 docx_handler 模块处理 DOCX 文件
"""

import os
import sys

# 添加脚本目录到路径
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from docx_handler import LabReportDocument


# 默认测试用例表头（根据新疆大学软件学院实验报告模板）
DEFAULT_HEADERS = [
    "测试用例编号",
    "测试项目",
    "测试子项",
    "测试用例描述",
    "前置条件",
    "测试输入",
    "操作步骤",
    "预期结果",
    "实际结果"
]


def insert_test_case_table(docx_path: str, test_cases: list,
                          headers: list = None) -> bool:
    """
    插入测试用例表格
    
    Args:
        docx_path: DOCX 文件路径
        test_cases: 测试用例数据列表，每项为一个列表
        headers: 表头列表（可选，使用默认表头）
    
    Returns:
        bool: 是否成功
    """
    if headers is None:
        headers = DEFAULT_HEADERS
    
    print(f"正在插入测试用例表格: {docx_path}")
    print(f"  表头: {headers}")
    print(f"  测试用例数: {len(test_cases)}")
    
    doc = LabReportDocument(docx_path)
    
    if not doc.open():
        print("错误: 无法打开文档")
        return False
    
    try:
        # 插入表格
        success = doc.insert_table(
            rows=len(test_cases) + 1,  # +1 for header row
            cols=len(headers),
            headers=headers,
            data=test_cases,
            style="Table Grid"
        )
        
        if success:
            print("测试用例表格插入成功")
        else:
            print("警告: 表格插入可能不完整")
        
        # 关闭文档（保存）
        return doc.close()
        
    except Exception as e:
        print(f"插入表格时出错: {e}")
        doc.close()
        return False


def generate_sample_test_cases() -> list:
    """生成示例测试用例数据"""
    return [
        ["TC-001", "新增客户信息", "必填项-公司名称", 
         "填写公司名称及所有必填项，提交表单",
         "用户已登录系统，具有新增客户权限",
         "公司名称：测试科技有限公司；申请类型：企业",
         "1. 登录系统\n2. 进入【客户管理】模块\n3. 点击"新增客户"\n4. 填写公司名称\n5. 选择申请类型"企业"\n6. 点击"确定"",
         "客户信息保存成功，提示"新增成功"",
         "待填写"],
        
        ["TC-002", "新增客户信息", "完整信息填写",
         "填写所有字段后提交",
         "用户已登录系统",
         "公司名称：新疆创新科技有限公司；申请类型：企业；公司背景：国有企业；公司地址：新疆乌鲁木齐；主营产品：软件开发",
         "1. 登录系统\n2. 进入【客户管理】模块\n3. 点击"新增客户"\n4. 依次填写所有字段\n5. 点击"确定"",
         "客户信息保存成功，所有字段正确存储",
         "待填写"],
    ]


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='插入测试用例表格')
    parser.add_argument('docx_path', help='DOCX 文件路径')
    parser.add_argument('--sample', '-s', action='store_true', 
                       help='使用示例测试用例数据')
    
    args = parser.parse_args()
    
    if args.sample:
        test_cases = generate_sample_test_cases()
    else:
        print("错误: 请提供测试用例数据或使用 --sample 参数")
        sys.exit(1)
    
    success = insert_test_case_table(args.docx_path, test_cases)
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
