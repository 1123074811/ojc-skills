#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
降AIGC率文本改写模块
提供多种提示词和API调用功能
"""

import os
import json
import requests
from typing import Optional, Dict, Any

PROMPT_1 = """
你的角色与目标：

你现在扮演一个专业的"论文（或技术文档）修改助手"。你的核心任务是接收一段中文原文（通常是技术性或学术性的描述），并将其改写成一种特定的风格。这种风格的特点是：比原文稍微啰嗦、更具解释性、措辞上更偏向通俗或口语化（但保持专业底线），并且系统性地使用特定的替代词汇和句式结构。你的目标是精确地模仿分析得出的修改模式，生成"修改后"风格的文本，同时务必保持原文的核心技术信息、逻辑关系和事实准确性，也不要添加过多的字数。

注意不要过于口语化（通常情况下不会过于口语化，有一些比如至于xxx呢，这种的不要有）
注意！你输出的内容不应原多于原文！应时刻记得字数和原文相符！
注意！不要有'xxx呢'这种形式，如'至于vue呢'
不要第一人称

输入与输出：

输入：一段中文原文（标记为"原文"）。
输出：一段严格按照以下规则修改后的中文文本（标记为"修改后"）。

核心修改手法与规则（请严格遵守）：

1. 增加冗余与解释性（Verbose Elaboration）：

动词短语扩展：将简洁的动词或动词短语替换为更长的、带有动作过程描述的短语。
示例："管理" -> "开展...的管理工作" 或 "进行管理"
示例："交互" -> "进行交互" 或 "开展交互"
示例："配置" -> "进行配置"
示例："处理" -> "去处理...工作"
示例："恢复" -> "进行恢复"
示例："实现" -> "得以实现" 或 "来实现"

增加辅助词/结构：在句子中添加语法上允许但非必需的词语，使句子更饱满。
示例：适当增加 "了"、"的"、"地"、"所"、"会"、"可以"、"这个"、"方面"、"当中" 等。
示例："提供功能" -> "有...功能" 或 "拥有...功能"

2. 系统性词汇替换（Systematic Synonym/Phrasing Substitution）：

特定动词/介词/连词替换：
采用 / 使用 -> 运用 / 选用 / 把...当作...来使用
基于 -> 鉴于 / 基于...来开展
利用 -> 借助 / 运用 / 凭借
通过 -> 借助 / 依靠 / 凭借
和 / 及 / 与 -> 以及（尤其是在列举多项时）
并 -> 并且 / 还 / 同时
其 -> 它 / 其 （可根据语境选择）

特定名词/形容词替换：
原因 -> 缘由 / 主要原因囊括...
符合 -> 契合
适合 -> 适宜
特点 -> 特性
提升 / 提高 -> 提高 / 提升 （可互换使用）
极大(地) -> 极大程度(上)
立即 -> 马上

3. 括号内容处理（Bracket Content Integration/Removal）：

解释性括号：对于原文中用于解释、举例或说明缩写的括号 (...) 或 （...）：
优先整合：尝试将括号内的信息自然地融入句子，使用 "也就是"、"即"、"比如"、"像" 等引导词。
示例：ORM（对象关系映射） -> 对象关系映射即ORM 或 ORM也就是对象关系映射
示例：功能（如ORM、Admin） -> 功能，比如ORM、Admin 或 功能，像ORM、Admin等

代码/标识符旁括号：对于紧跟在代码、文件名、类名旁的括号，通常直接移除括号。
示例：视图 (views.py) 中 -> 视图也就是views.py中
示例：权限类 (admin_panel.permissions) -> 权限类 admin_panel.permissions

4. 句式微调与口语化倾向（Sentence Structure & Colloquial Touch）：

使用"把"字句：在合适的场景下，倾向于使用"把"字句。
示例："会将对象移动" -> "会把对象移动"

条件句式转换：将较书面的条件句式改为稍口语化的形式。
示例："若...，则..." -> "要是...，那就..." 或 "如果...，就..."

名词化与动词化转换：根据需要进行调整。
示例："为了将...解耦" -> "为了实现...的解耦"

增加语气词/连接词：如在句首或句中添加"那么"、"这样"、"同时"等。

5. 保持技术准确性（Maintain Technical Accuracy）：

绝对禁止修改：所有的技术术语（如 Django, RESTful API, MySQL, JWT, ORM）、代码片段 (views.py, settings.py)、库名、配置项、API 路径等必须保持原样，不得修改或错误转写。

核心逻辑不变：修改后的句子必须表达与原文完全相同的技术逻辑、因果关系和功能描述。

执行指令：

请根据以上所有规则，对接下来提供的"原文"进行修改，生成符合上述特定风格的"修改后"文本。务必仔细揣摩每个规则的细节和示例，力求在风格上高度一致。

注意不要过于口语化！
注意！你输出的内容不应原多于原文！应时刻记得字数和原文相符！
注意！不要有'xxx呢'这种形式！
不要第一人称！
每段内容最好使用"分-总-分"结构重写。

原文：
{content}

修改后：
"""


def rewrite_text(content: str, api_config: Optional[Dict[str, str]] = None, 
                prompt_num: int = 1, model: str = "gpt-3.5-turbo") -> str:
    """
    使用外部API改写文本以降低AIGC率
    
    Args:
        content: 要改写的原文
        api_config: API配置字典，包含api_key, base_url等
        prompt_num: 提示词编号（1-4）
        model: 使用的模型
        
    Returns:
        str: 改写后的文本
    """
    if not content.strip():
        return content
    
    # 如果没有提供API配置，返回提示词
    if not api_config or not api_config.get('api_key'):
        print("警告: 未提供API配置，返回提示词而非实际改写")
        return get_prompt(prompt_num, content)
    
    try:
        prompt = get_prompt(prompt_num, content)
        
        # 调用OpenAI兼容API
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_config['api_key']}"
        }
        
        data = {
            "model": model,
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.7,
            "max_tokens": min(len(content) * 2, 2000)  # 避免过长
        }
        
        # 使用自定义base_url或默认OpenAI地址
        api_url = api_config.get('base_url', 'https://api.openai.com/v1/chat/completions')
        
        response = requests.post(api_url, headers=headers, json=data, timeout=30)
        response.raise_for_status()
        
        result = response.json()
        
        if 'choices' in result and len(result['choices']) > 0:
            rewritten_text = result['choices'][0]['message']['content'].strip()
            
            # 提取改写后的内容（去除"修改后："等前缀）
            if "修改后：" in rewritten_text:
                rewritten_text = rewritten_text.split("修改后：")[1].strip()
            elif "修改后:" in rewritten_text:
                rewritten_text = rewritten_text.split("修改后:")[1].strip()
            
            return rewritten_text
        else:
            print(f"API返回格式异常: {result}")
            return content
            
    except requests.exceptions.RequestException as e:
        print(f"API请求失败: {e}")
        return content
    except Exception as e:
        print(f"改写过程中出错: {e}")
        return content


def load_api_config(config_path: str = "aigc_config.json") -> Optional[Dict[str, str]]:
    """
    从配置文件加载API配置
    
    Args:
        config_path: 配置文件路径
        
    Returns:
        Dict: API配置字典
    """
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            print(f"配置文件不存在: {config_path}")
            return None
    except Exception as e:
        print(f"加载配置文件失败: {e}")
        return None


def create_sample_config(output_path: str = "aigc_config.json"):
    """
    创建示例配置文件
    
    Args:
        output_path: 输出路径
    """
    sample_config = {
        "api_key": "your-api-key-here",
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-3.5-turbo",
        "prompt_num": 1,
        "description": "AIGC降重率API配置文件"
    }
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(sample_config, f, ensure_ascii=False, indent=2)
        print(f"示例配置文件已创建: {output_path}")
    except Exception as e:
        print(f"创建配置文件失败: {e}")


def get_prompt(prompt_num: int = 1, content: str = "") -> str:
    """
    获取降AIGC率提示词
    
    Args:
        prompt_num: 提示词编号（1-4）
        content: 要改写的原文
        
    Returns:
        完整的提示词文本
    """
    if prompt_num == 1:
        return PROMPT_1.format(content=content)
    else:
        return PROMPT_1.format(content=content)  # 默认使用提示词1


if __name__ == '__main__':
    import sys
    import argparse
    
    parser = argparse.ArgumentParser(description='降AIGC率文本改写工具')
    parser.add_argument('content', nargs='?', help='要改写的文本内容')
    parser.add_argument('--prompt', '-p', type=int, default=1, help='提示词编号（1-4）')
    parser.add_argument('--config', '-c', help='API配置文件路径')
    parser.add_argument('--create-config', action='store_true', help='创建示例配置文件')
    parser.add_argument('--file', '-f', help='从文件读取文本内容')
    parser.add_argument('--output', '-o', help='输出文件路径')
    
    args = parser.parse_args()
    
    # 创建配置文件
    if args.create_config:
        create_sample_config()
        sys.exit(0)
    
    # 获取文本内容
    if args.file:
        try:
            with open(args.file, 'r', encoding='utf-8') as f:
                content = f.read().strip()
        except Exception as e:
            print(f"读取文件失败: {e}")
            sys.exit(1)
    elif args.content:
        content = args.content
    else:
        print("错误: 请提供文本内容或使用 --file 参数")
        parser.print_help()
        sys.exit(1)
    
    # 加载API配置
    api_config = None
    if args.config:
        api_config = load_api_config(args.config)
    
    # 执行改写
    if api_config:
        print(f"正在使用API改写文本（提示词{args.prompt}）...")
        rewritten = rewrite_text(content, api_config, args.prompt)
    else:
        print(f"未提供API配置，返回提示词（提示词{args.prompt}）...")
        rewritten = get_prompt(args.prompt, content)
    
    # 输出结果
    if args.output:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(rewritten)
            print(f"结果已保存到: {args.output}")
        except Exception as e:
            print(f"保存文件失败: {e}")
            print("\n改写结果:")
            print(rewritten)
    else:
        print("\n改写结果:")
        print(rewritten)
