# -*- coding: utf-8 -*-
"""
未发货订单统计 Skill
一键完成未发货订单的数据清洗、透视表生成、统计计算、可视化图表生成

支持三种触发方式：
1. 邮件触发：trigger_by_email(email_subject, email_content)
2. 消息触发：trigger_by_message(message_text)
3. 文件触发：trigger_by_file(file_path)
"""

import os
from .process_data import process_data
from .generate_chart import generate_chart
from .trigger import (
    trigger_by_email,
    trigger_by_message,
    trigger_by_file,
    TRIGGER_KEYWORDS,
    EMAIL_TRIGGER_SUBJECT
)


def run_unshipped_order_statistics(input_excel_path, output_dir="处理结果"):
    """
    执行未发货订单统计完整流程
    
    一站式处理：数据清洗 → 透视表生成 → 统计表计算 → HTML可视化 → PNG截图
    
    Args:
        input_excel_path: 原始Excel文件路径
        output_dir: 输出目录，默认"处理结果"
    
    Returns:
        dict: 包含以下键的结果字典
            - excel_file: 处理后的Excel文件路径
            - html_file: HTML报告文件路径
            - png_file: PNG图片文件路径（如生成成功）
    
    Examples:
        >>> from 未发货订单统计 import run_unshipped_order_statistics
        >>> result = run_unshipped_order_statistics("原始数据.xlsx")
        >>> print(f"Excel: {result['excel_file']}")
        >>> print(f"HTML: {result['html_file']}")
    """
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    print("=" * 70)
    print("📦 未发货订单统计 - 完整流程开始")
    print("=" * 70)
    
    # 第一步：数据处理
    print("\n🔧 步骤 1/2: 数据处理中...")
    excel_result = process_data(input_excel_path, output_dir)
    
    if not excel_result:
        raise RuntimeError("数据处理失败，请检查输入文件")
    
    # 第二步：生成可视化图表
    print("\n🎨 步骤 2/2: 可视化生成中...")
    chart_result = generate_chart(excel_result, output_dir)
    
    print("\n" + "=" * 70)
    print("✅ 未发货订单统计 - 全部完成！")
    print("=" * 70)
    print(f"\n📊 处理结果:")
    print(f"   Excel: {excel_result}")
    print(f"   HTML:  {chart_result['html']}")
    if chart_result.get('png'):
        print(f"   PNG:   {chart_result['png']}")
    print()
    
    # 返回结果
    return {
        "excel_file": excel_result,
        "html_file": chart_result['html'],
        "png_file": chart_result.get('png')
    }


__all__ = [
    'run_unshipped_order_statistics',
    'trigger_by_email',
    'trigger_by_message',
    'trigger_by_file',
    'TRIGGER_KEYWORDS',
    'EMAIL_TRIGGER_SUBJECT'
]
__version__ = '1.0.0'
__author__ = 'Unshipped Order Statistics Skill'
