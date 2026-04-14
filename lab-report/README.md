# 实验报告自动化处理工具 - 架构重构版

## 概述

这是新疆大学软件学院实验报告自动化填写与格式规范化工具的重构版本，解决了原有架构中的关键问题：

### 🔧 已修复的架构问题

1. **平台耦合严重** - 移除了 `win32com.client` 依赖，统一使用 `python-docx`
2. **docx_handler.py 实现不完整** - 修复了所有占位符和缺失的导入
3. **硬编码路径问题** - 实现了多路径自动检测机制
4. **封面填写逻辑脆弱** - 改进了鲁棒性和匹配精度
5. **AIGC改写功能缺失** - 实现了完整的API调用逻辑
6. **删除操作未执行** - 修复了模板分析中的删除逻辑
7. **示例数据不匹配** - 更新为实验报告相关的测试用例
8. **导入不一致** - 统一了文档处理模块的导入方式

## 📁 文件结构

```
lab-report/
├── scripts/
│   ├── docx_handler.py              # 核心文档处理模块（已修复）
│   ├── lab_report_workflow.py       # 统一工作流（新增）
│   ├── rewrite_aigc.py              # AIGC降重模块（已实现API）
│   ├── format_document_unified.py   # 格式化处理（无win32com依赖）
│   ├── cleanup_spacing_unified.py   # 空行清理（无win32com依赖）
│   ├── analyze_template_unified.py  # 模板分析（无win32com依赖）
│   ├── insert_test_cases.py         # 测试用例插入（已更新示例）
│   └── document_utils.py            # 工具模块（已修复导入）
├── SKILL.md                         # 技能文档（已更新）
└── README.md                        # 本文档
```

## 🚀 快速开始

### 安装依赖

```bash
pip install python-docx requests
```

### 基本使用

```bash
# 最简单的用法
python scripts/lab_report_workflow.py template.docx \
  --student-id 20232501306 \
  --name 张三 \
  --class 软件23-4

# 完整功能
python scripts/lab_report_workflow.py template.docx \
  --student-id 20232501306 \
  --name 张三 \
  --class 软件23-4 \
  --insert-test-cases \
  --use-aigc \
  --api-config aigc_config.json \
  --output result.docx
```

## 📋 功能详解

### 1. 核心工作流 (`lab_report_workflow.py`)

统一的实验报告处理入口，集成了所有功能：

- **模板分析** - 自动识别报告结构
- **封面填写** - 智能填充个人信息
- **测试用例插入** - 支持多种类型的测试用例
- **AIGC降重** - 可选的AI文本改写
- **格式统一** - 自动调整字体、字号、缩进
- **空行清理** - 规范化段落间距

### 2. 文档处理核心 (`docx_handler.py`)

重构后的核心模块，解决了所有实现缺陷：

```python
from scripts.docx_handler import LabReportDocument

doc = LabReportDocument("template.docx")
doc.open()
doc.fill_cover_info("20232501306", "张三", "软件23-4")
doc.close("result.docx")
```

**改进点：**
- ✅ 移除硬编码路径，支持多路径自动检测
- ✅ 修复 `qn` 函数导入问题
- ✅ 改进封面填写的鲁棒性
- ✅ 实现所有占位符逻辑

### 3. AIGC降重模块 (`rewrite_aigc.py`)

完整的API调用实现，不再是空壳：

```python
from scripts.rewrite_aigc import rewrite_text, load_api_config

# 加载配置
api_config = load_api_config("aigc_config.json")

# 改写文本
rewritten = rewrite_text("原始文本", api_config, prompt_num=1)
```

**功能特性：**
- ✅ 支持OpenAI兼容API
- ✅ 配置文件管理
- ✅ 自动降级到提示词模式
- ✅ 完整的命令行界面

### 4. 统一格式处理模块

替代原有的win32com依赖版本：

- `format_document_unified.py` - 文档格式化
- `cleanup_spacing_unified.py` - 空行清理  
- `analyze_template_unified.py` - 模板分析

**优势：**
- ✅ 跨平台兼容（Windows/Linux/Mac）
- ✅ 无需Microsoft Word
- ✅ 纯Python实现

## ⚙️ 配置说明

### AIGC API配置

创建配置文件 `aigc_config.json`：

```json
{
  "api_key": "your-api-key-here",
  "base_url": "https://api.openai.com/v1",
  "model": "gpt-3.5-turbo",
  "prompt_num": 1,
  "description": "AIGC降重率API配置文件"
}
```

或使用命令行创建：

```bash
python scripts/rewrite_aigc.py --create-config
```

### 测试用例类型

支持多种测试用例模板：

- `default` - 通用测试用例
- `software` - 软件测试专用
- `performance` - 性能测试专用

## 📝 使用示例

### 示例1：基础实验报告处理

```bash
python scripts/lab_report_workflow.py "实验一模板.docx" \
  --student-id 20232501306 \
  --name "欧劲聪" \
  --class "软件23-4"
```

输出：`欧劲聪-软件23-4-20232501306-实验一 测试用例设计.docx`

### 示例2：包含测试用例和AIGC改写

```bash
python scripts/lab_report_workflow.py "实验一模板.docx" \
  --student-id 20232501306 \
  --name "欧劲聪" \
  --class "软件23-4" \
  --insert-test-cases \
  --test-case-type software \
  --use-aigc \
  --api-config my_config.json
```

### 示例3：单独使用各个模块

```bash
# 分析模板
python scripts/analyze_template_unified.py template.docx

# 填写封面
python scripts/fill_cover.py template.docx 20232501306 "张三" "软件23-4"

# 插入测试用例
python scripts/insert_test_cases.py result.docx --sample

# 格式化文档
python scripts/format_document_unified.py result.docx

# 清理空行
python scripts/cleanup_spacing_unified.py result.docx
```

## 🛠️ 开发说明

### 架构改进

1. **依赖解耦** - 移除平台特定依赖，提高可移植性
2. **错误处理** - 增强异常处理和用户友好的错误信息
3. **配置管理** - 标准化配置文件格式和管理方式
4. **模块化设计** - 每个功能独立，便于维护和测试

### 扩展性

代码设计支持以下扩展：

- 新增测试用例类型
- 支持更多文档格式
- 集成其他AI服务
- 自定义格式规范

## 🔍 故障排除

### 常见问题

1. **"docx skill 不可用"**
   - 这是正常提示，系统会自动使用python-docx备选方案

2. **"API配置缺失"**
   - 创建配置文件或跳过AIGC功能

3. **"模板格式异常"**
   - 检查模板是否为有效的docx文件
   - 确认封面字段包含"学号"、"姓名"、"班级"等关键词

4. **权限问题**
   - 确保对模板文件有读写权限
   - 输出目录需要写入权限

### 调试模式

设置环境变量启用详细日志：

```bash
export LAB_REPORT_DEBUG=1
python scripts/lab_report_workflow.py ...
```

## 📄 许可证

本项目遵循开源许可证，详见LICENSE文件。

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个工具。

---

**版本**: 2.0.0 (架构重构版)  
**更新时间**: 2024年  
**兼容性**: Python 3.7+, 跨平台
