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


# 软件测试实验报告专用表头
SOFTWARE_TEST_HEADERS = [
    "测试用例编号",
    "测试模块",
    "测试功能点",
    "测试用例描述",
    "前置条件",
    "测试数据",
    "操作步骤",
    "预期结果",
    "测试结果"
]


# 性能测试实验报告表头
PERFORMANCE_TEST_HEADERS = [
    "测试编号",
    "测试场景",
    "性能指标",
    "测试描述",
    "环境配置",
    "测试参数",
    "执行步骤",
    "预期性能",
    "实际性能"
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
    """生成示例测试用例数据 - 软件测试实验报告相关"""
    return [
        ["TC-001", "登录功能测试", "正常登录", 
         "使用正确的用户名和密码登录系统",
         "系统正常运行，用户已注册",
         "用户名：test01；密码：123456",
         "1. 打开登录页面\n2. 输入用户名test01\n3. 输入密码123456\n4. 点击登录按钮",
         "登录成功，跳转到主页面",
         "登录成功，页面正常跳转"],
        
        ["TC-002", "登录功能测试", "密码错误", 
         "使用错误的密码登录系统",
         "系统正常运行，用户已注册",
         "用户名：test01；密码：wrongpass",
         "1. 打开登录页面\n2. 输入用户名test01\n3. 输入错误密码wrongpass\n4. 点击登录按钮",
         "登录失败，提示密码错误",
         "登录失败，显示错误提示信息"],
        
        ["TC-003", "用户注册测试", "信息完整注册", 
         "填写所有必填信息进行用户注册",
         "系统正常运行，用户名未被占用",
         "用户名：newuser；密码：pass123；确认密码：pass123；邮箱：user@test.com",
         "1. 打开注册页面\n2. 填写用户名newuser\n3. 填写密码pass123\n4. 确认密码pass123\n5. 填写邮箱user@test.com\n6. 点击注册按钮",
         "注册成功，跳转到登录页面",
         "注册成功，显示成功提示"],
        
        ["TC-004", "数据查询测试", "按条件查询", 
         "根据指定条件查询数据",
         "数据库中有测试数据",
         "查询条件：日期范围2024-01-01到2024-12-31；状态：已完成",
         "1. 进入数据查询页面\n2. 选择日期范围\n3. 选择状态为已完成\n4. 点击查询按钮",
         "显示符合条件的查询结果",
         "查询结果正确显示，数据完整"],
        
        ["TC-005", "表单验证测试", "必填项验证", 
         "提交表单时不填写必填项",
         "表单包含必填字段",
         "不填写用户名和密码，直接提交",
         "1. 打开注册页面\n2. 不填写任何信息\n3. 直接点击提交按钮",
         "提示必填项不能为空",
         "显示验证错误信息，阻止提交"],
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
