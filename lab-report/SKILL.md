---
name: lab-report
version: 2.0.0
description: 新疆大学软件学院实验报告自动化填写与格式规范化工具。支持封面信息填写、测试用例表格插入、图表题注管理、格式统一调整（宋体/Times New Roman、字号、缩进等），并提供降AIGC率文本改写功能。
description_zh: 新疆大学软件学院实验报告自动化填写与格式规范化工具
description_en: Xinjiang University Software School Lab Report Automation Tool
allowed-tools: Read, Write, Bash, Python
---

# lab-report（实验报告）

自动化完成新疆大学软件学院实验报告的填写与格式规范化。

## 最小用户输入（新对话首条消息）

```
$lab-report
模板地址：<template_docx>
学号：<student_id>
姓名：<name>
班级：<class_name>
是否需要降AIGC率改写：[是/否]
```

**说明**：
- 只需提供模板文件的绝对地址和个人信息
- Skill 会自动分析模板，提取实验报告名称
- 根据模板结构自动填写内容
- 输出文件自动命名为：`<姓名>-<班级>-<学号>-<实验报告名>.docx`

## 默认约束

### 文件命名规范
- 完成后的文件命名为：`<姓名>-<班级>-<学号>-<实验报告名>.docx`
- 实验报告名根据模板自动提取（如"实验一 测试用例设计"）

### 封面格式
- 仅填写：学号、姓名、班级、日期
- 不修改任何其他格式（字体、字号、间距等保持原模板）
- 日期使用当日日期

### 正文格式规范
- **正文字号**：小四（12pt）
- **中文字体**：宋体
- **英文/数字字体**：Times New Roman
- **段落缩进**：首行缩进 2 字符
- **标题与正文间距**：本标题结尾与下一个标题之间最多 1 个空行（多余删除）
- **分页**：最大标题与上一段内容之间进行分页

### 图表格式规范
- **表格内字号**：五号（10.5pt）
- **图片/表格注释字号**：五号（10.5pt）
- **题注格式**：`图X-Y` 或 `表X-Y`（X为章节号，Y为序号）
- **注释格式**：`图X-Y-本标题下的第几张图片（表格）-图片（表格）内容说明`
- **对齐方式**：图片行和注释行都居中（无缩进）
- **图片占位**：图片留空位给用户导入，保留题注行

### 内容生成
- 使用降AIGC率提示词（优先使用提示词1）对AI生成内容进行改写
- 改写要点：
  - 增加冗余与解释性（动词短语扩展）
  - 系统性词汇替换（采用/使用→运用/选用，基于→鉴于，通过→借助等）
  - 括号内容整合或移除
  - 使用"把"字句
  - 保持技术术语不变
  - 采用"分-总-分"结构

## 工作流

### 主流程

1. **分析用户意图**
   - **填写实验报告？** → 执行"填写工作流"
   - **导出表格？** → 执行"导出工作流"
   - **处理PDF？** → 执行"PDF工作流"

### 填写工作流

2. **分析模板**
   - 读取模板 DOCX 文件
   - 提取实验报告名称（从文件名或文档内容）
   - 识别封面字段位置（学号、姓名、班级、日期）
   - 识别正文章节结构
   - 识别表格插入位置和表头要求

3. **填写封面信息**
   - 学号：用户提供
   - 姓名：用户提供
   - 班级：用户提供
   - 日期：当日日期
   - **仅填写内容，不修改任何格式**（字体、字号、间距保持原模板）

4. **生成/改写内容**
   - 使用降AIGC率提示词1进行改写
   - 确保字数与原文相符
   - 避免过于口语化（不要有"至于xxx呢"等形式）
   - 不使用第一人称
   - 采用"分-总-分"结构

5. **插入测试用例表格**
   - 使用五号字
   - 表头根据实验要求确定
   - 表格宽度适配页面（不超出边界）
   - 单元格垂直居中

6. **管理图表题注**
   - 图片/表格后插入题注行
   - 题注格式：`图X-Y 标题` 或 `表X-Y 标题`
   - 图片行和注释行居中
   - 图片留空位

7. **格式统一调整**
   - 正文：小四、宋体、Times New Roman（英文/数字）
   - 首行缩进 2 字符
   - 清理多余空行（标题间最多1个）
   - 最大标题前分页

8. **最终检查**
   - 检查全文格式一致性
   - 验证文件命名规范
   - 确认封面信息正确

### 导出工作流

2. **读取 DOCX**
   - 打开源文件
   - 识别所有表格

3. **提取表格**
   - 遍历每个表格
   - 提取表头和数据

4. **生成 Excel**
   - 每个表格生成独立文件
   - 自动调整列宽
   - 设置表头格式

### PDF工作流

2. **确定操作类型**
   - **合并PDF** → 执行合并操作
   - **拆分PDF** → 执行拆分操作
   - **提取文本** → 执行文本提取

3. **执行对应操作**
   - 合并：按顺序合并多个文件
   - 拆分：按指定页数拆分
   - 提取：提取文本并保存

## 依赖要求

本 Skill 集成了多个文档处理技能：

### 1. DOCX 处理（docx skill）
- **优先使用**: docx skill（如果已安装）
  - 自动检测多个标准路径：
    - `~/.claude/skills/docx`
    - `C:\Users\[用户名]\.claude\skills\docx`
    - `/usr/local/.claude/skills/docx`
  - 提供高级功能：tracked changes、comments、精确的 XML 操作
- **备选方案**: python-docx 库
  - 基本功能：段落、表格、格式设置
  - 无需额外安装，直接可用

### 2. PDF 处理（pdf skill）
- **pypdf**: PDF 读写、合并、拆分
- **pdfplumber**: 文本和表格提取
- **reportlab**: PDF 创建

### 3. XLSX 处理（xlsx skill）
- **pandas**: 数据分析和处理
- **openpyxl**: Excel 公式和格式

### 安装依赖

```bash
# 安装各 skill（推荐）
# 将 docx.skill, pdf.skill, xlsx.skill 安装到 ~/.claude/skills/ 目录
# 支持多个标准路径，系统会自动检测

# 或安装 Python 库（备选）
pip install python-docx pypdf pdfplumber reportlab pandas openpyxl requests
```

## 脚本工具

### 核心模块

- `scripts/docx_handler.py`：**DOCX 处理核心模块**，集成 docx skill 功能
  - `LabReportDocument(docx_path)` - 创建文档对象
  - `.open()` - 打开文档
  - `.fill_cover_info(student_id, name, class_name, date)` - 填写封面信息
  - `.insert_table(rows, cols, headers, data, style)` - 插入表格
  - `.set_paragraph_format(alignment, font_name, font_size, first_line_indent)` - 设置段落格式
  - `.cleanup_empty_paragraphs()` - 清理空行
  - `.close(save_path=None)` - 关闭并保存文档

- `scripts/document_utils.py`：**综合文档处理模块**，集成 pdf 和 xlsx 技能
  - `DocumentProcessor(work_dir)` - 创建处理器对象
  - `.merge_pdfs(pdf_files, output_path)` - 合并多个 PDF
  - `.extract_pdf_text(pdf_path, output_txt)` - 提取 PDF 文本
  - `.split_pdf(pdf_path, output_dir, pages_per_file)` - 拆分 PDF
  - `.create_excel_from_table(table_data, headers, output_path, sheet_name)` - 从表格创建 Excel
  - `.export_docx_tables_to_excel(docx_path, output_dir)` - 导出 DOCX 表格到 Excel
  - `.create_test_case_excel(test_cases, output_path)` - 创建测试用例 Excel

### 功能脚本

- `scripts/analyze_template.py`：**分析模板**，提取报告名称、章节结构、表格位置
- `scripts/fill_cover.py`：填写封面信息（调用 docx_handler）
- `scripts/insert_test_cases.py`：插入测试用例表格（调用 docx_handler）
- `scripts/format_document.py`：统一文档格式（字体、字号、缩进）
- `scripts/manage_captions.py`：管理图表题注
- `scripts/cleanup_spacing.py`：清理多余空行和分页
- `scripts/rewrite_aigc.py`：**降AIGC率文本改写**，支持外部API调用
  - `rewrite_text(content, api_config, prompt_num, model)` - 使用API改写文本
  - `load_api_config(config_path)` - 加载API配置
  - `create_sample_config()` - 创建示例配置文件
  - `get_prompt(prompt_num, content)` - 获取提示词（备选方案）
- `scripts/final_check.py`：最终格式检查

## 资源文件

- `references/format-rules.md`：格式规范详细说明
- `references/aigc-prompts.md`：降AIGC率提示词

## 使用示例

### 示例 1：填写实验报告

**Input:**
```
$lab-report
模板地址：E:\软件测试与质量控制\《软件测试与质量控制》实验报告（一）.doc
学号：20232501306
姓名：欧劲聪
班级：软件23-4
降AIGC率：是
```

**Output:**
- 文件：`欧劲聪-软件23-4-20232501306-实验一 测试用例设计.docx`
- 封面：已填写学号、姓名、班级、日期
- 表格：已插入测试用例表格（9列×19行）
- 格式：已统一为小四/宋体/Times New Roman

### 示例 2：填写小学期实训报告

**Input:**
```
$lab-report
模板地址：E:\大三上小学期\6-新疆大学软件学院小学期实训报告-最终验收参考此格式.docx
学号：20232501306
姓名：欧劲聪
班级：软件23-4
实验内容：图书管理系统
降AIGC率：是
```

**Output:**
- 文件：`欧劲聪-软件23-4-20232501306-小学期实训报告.docx`
- 封面：已填写个人信息
- 内容：已根据"图书管理系统"主题生成各章节
- 图表：已插入必要的表格和图表占位符
- 题注：已管理为图X-Y格式

### 示例 3：导出表格到 Excel

**Input:**
```
$lab-report
操作：导出表格
源文件：E:\软件测试与质量控制\《软件测试与质量控制》实验报告（一）.docx
输出目录：E:\导出表格
```

**Output:**
- 文件：`实验报告（一）_table_1.xlsx`, `实验报告（一）_table_2.xlsx`...
- 内容：DOCX 中的所有表格已导出
- 格式：自动调整列宽，表头加粗

### 示例 4：合并多个 PDF

**Input:**
```
$lab-report
操作：合并PDF
源文件：E:\报告\part1.pdf, E:\报告\part2.pdf, E:\报告\part3.pdf
输出文件：E:\报告\完整报告.pdf
```

**Output:**
- 文件：`完整报告.pdf`
- 内容：按顺序合并的 PDF
- 格式：保持原始格式和书签

### 示例 5：拆分 PDF

**Input:**
```
$lab-report
操作：拆分PDF
源文件：E:\资料\完整文档.pdf
每页数：5
输出目录：E:\拆分结果
```

**Output:**
- 文件：`split_001.pdf`, `split_002.pdf`...
- 内容：每 5 页为一个文件
- 格式：保持原始格式

## 错误处理

| 错误类型 | 处理方式 |
|---------|---------|
| docx skill 未安装 | 自动降级到 python-docx |
| PDF 库未安装 | 提示安装 pypdf 或 PyPDF2 |
| XLSX 库未安装 | 提示安装 openpyxl |
| 文件路径含中文 | 自动复制到临时目录处理 |
| 表格超出页面 | 自动调整列宽和字体大小 |
| 模板格式异常 | 提示用户检查模板文件 |
| 必填信息缺失 | 提示用户补充学号/姓名/班级 |
| API配置缺失 | 降级为提示词模式，提供配置示例 |
| 网络连接失败 | 自动重试，失败后使用备选方案 |

## 注意事项

1. **个人信息必填**：学号、姓名、班级为必填项，不提供则无法生成正确文件名
2. **模板格式**：封面字段位置需符合常规实验报告格式
3. **中文路径**：已自动处理，无需担心
4. **备份原文件**：建议在操作前备份原始模板文件
