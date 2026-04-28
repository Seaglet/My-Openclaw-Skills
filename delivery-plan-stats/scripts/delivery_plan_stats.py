# -*- coding: utf-8 -*-
# 任务名：统计发货计划

"""
发货计划统计脚本
处理销售单数据，生成发货计划透视表
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import re
import os

# ============================================
# 配置区域 - 请根据实际需求修改以下变量
# ============================================

# 输入文件路径（Excel文件，支持.xlsx格式）
INPUT_FILE = "用户上传/销售单查询(总时间：4676毫秒,数据获取：3841毫秒)_1777358371610_0_nxpf.xlsx"

# 输出目录
OUTPUT_DIR = "用户上传/"

# ============================================
# 脚本逻辑 - 无需修改
# ============================================

def main():
    # 获取当天日期作为文件后缀
    today = datetime.now().strftime("%Y%m%d")
    
    # 生成输出文件路径（自动添加日期后缀）
    OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"发货计划统计结果{today}.xlsx")
    IMG_PATH = os.path.join(OUTPUT_DIR, f"发货计划透视表{today}.png")
    
    # ========== 步骤0: 读取数据 ==========
    print("=== 步骤0: 读取Excel数据 ===")
    df = pd.read_excel(INPUT_FILE, sheet_name="销售单")
    print(f"原始数据形状: {df.shape}")
    
    # ========== 步骤1: 插入"客审？"列 ==========
    print("\n=== 步骤1: 插入'客审？'列 ===")
    def get_客审(mark):
        if pd.isna(mark):
            return "未客审"
        if "已客审" in str(mark) or "已货审" in str(mark):
            return "已客审"
        return "未客审"
    
    df["客审？"] = df["标记"].apply(get_客审)
    print(f"客审？列统计:\n{df['客审？'].value_counts()}")
    
    # ========== 步骤2: 插入"发货时间类型"列 ==========
    print("\n=== 步骤2: 插入'发货时间类型'列 ===")
    def get_发货时间类型(status):
        if pd.isna(status) or str(status).strip() == "":
            return ""
        status = str(status)
        if status in ["加急发货", "指定时间发货"]:
            return "指定时间"
        if "有货就发" in status or "等通知" in status:
            return "等通知"
        return ""
    
    df["发货时间类型"] = df["订单-发货通知状态"].apply(get_发货时间类型)
    print(f"发货时间类型列统计:\n{df['发货时间类型'].value_counts()}")
    
    # ========== 步骤3: 修改时间格式 ==========
    print("\n=== 步骤3: 修改'订单-人工承诺发货时间'格式为YYYY年MM月 ===")
    def format_time(t):
        if pd.isna(t):
            return ""
        try:
            if isinstance(t, str):
                dt = pd.to_datetime(t)
            else:
                dt = t
            return dt.strftime("%Y年%m月")
        except:
            return ""
    
    df["计划发货月份"] = df["订单-人工承诺发货时间"].apply(format_time)
    print(f"计划发货月份示例:\n{df['计划发货月份'].head(10)}")
    
    # ========== 步骤4: 插入"发货订单类型"列 ==========
    print("\n=== 步骤4: 插入'发货订单类型'列 ===")
    df["发货订单类型"] = df["客审？"] + "-" + df["发货时间类型"]
    # 如果发货时间类型为空，清理格式
    df["发货订单类型"] = df["发货订单类型"].str.replace(r"-$", "", regex=True)
    print(f"发货订单类型列统计:\n{df['发货订单类型'].value_counts()}")
    
    # ========== 步骤5: 插入"发货？"列 ==========
    print("\n=== 步骤5: 插入'发货？'列 ===")
    def get_发货(row):
        发货订单类型 = str(row["发货订单类型"])
        if "未客审" in 发货订单类型:
            return "未客审"
        elif "已客审-指定时间" in 发货订单类型:
            return row["计划发货月份"]
        elif "已客审-等通知" in 发货订单类型:
            return "等通知"
        return "未客审"
    
    df["发货？"] = df.apply(get_发货, axis=1)
    print(f"发货？列统计:\n{df['发货？'].value_counts()}")
    
    # ========== 步骤6: 创建透视表 ==========
    print("\n=== 步骤6: 创建透视表 ===")
    
    # 过滤掉销售渠道为空的数据（虽然这里没有）
    df_filtered = df[df["销售渠道"].notna() & (df["销售渠道"] != "")]
    
    # 创建透视表
    pivot = pd.pivot_table(
        df_filtered,
        values="应收合计",
        index="销售渠道",
        columns="发货？",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="合计"
    )
    
    # 重新排列列顺序
    # 先找出所有月份列并排序
    all_cols = list(pivot.columns)
    month_cols = [c for c in all_cols if c not in ["合计", "等通知", "未客审"]]
    other_cols = [c for c in all_cols if c in ["等通知", "未客审"]]
    
    # 对月份列进行排序
    def sort_key(x):
        if x == "合计":
            return (9999, "")
        if x == "等通知":
            return (1, "等通知")
        if x == "未客审":
            return (2, "未客审")
        try:
            dt = pd.to_datetime(x.replace("年", "-").replace("月", ""))
            return (0, dt)
        except:
            return (3, x)
    
    month_cols_sorted = sorted(month_cols, key=sort_key)
    new_order = month_cols_sorted + other_cols
    if "合计" in all_cols:
        new_order.append("合计")
    
    # 重新排列列
    existing_cols = [c for c in new_order if c in pivot.columns]
    pivot = pivot[existing_cols]
    
    print(f"透视表形状: {pivot.shape}")
    print(f"透视表列顺序: {list(pivot.columns)}")
    print(f"\n透视表预览:\n{pivot}")
    
    # ========== 步骤7: 保存结果 ==========
    print("\n=== 步骤7: 保存处理后的Excel和透视表 ===")
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        # 写入处理后的数据
        df.to_excel(writer, sheet_name='销售单处理后', index=False)
        
        # 写入透视表
        pivot.to_excel(writer, sheet_name='透视表')
        
        # 获取透视表工作表用于格式化
        pivot_ws = writer.sheets['透视表']
        
        # 设置列宽
        pivot_ws.column_dimensions['A'].width = 30  # 销售渠道列
        for i in range(2, len(pivot.columns) + 2):
            pivot_ws.column_dimensions[get_column_letter(i)].width = 15
        
        # 设置金额格式（万元）
        for row in range(2, len(pivot) + 3):  # 包括合计行
            for col in range(2, len(pivot.columns) + 1):
                cell = pivot_ws.cell(row=row, column=col)
                if cell.value is not None and isinstance(cell.value, (int, float)):
                    cell.value = cell.value / 10000
                    cell.number_format = '0.00"万"'
        
        # 设置合计行格式加粗
        last_data_row = len(pivot) + 2  # 合计行
        for col in range(1, len(pivot.columns) + 2):
            cell = pivot_ws.cell(row=last_data_row, column=col)
            cell.font = Font(bold=True)
        
        # 设置表头加粗
        for col in range(1, len(pivot.columns) + 2):
            cell = pivot_ws.cell(row=1, column=col)
            cell.font = Font(bold=True)
    
    print(f"结果已保存到: {OUTPUT_FILE}")
    
    # ========== 步骤8: 生成透视表截图 ==========
    print("\n=== 步骤8: 生成透视表截图 ===")
    generate_pivot_screenshot(pivot, IMG_PATH)
    
    return df, pivot

def generate_pivot_screenshot(pivot_df, img_path):
    """生成透视表的图片截图"""
    # 创建DataFrame展示
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use('Agg')
    
    # 设置中文字体
    plt.rcParams['font.sans-serif'] = ['Noto Sans CJK SC', 'WenQuanYi Micro Hei', 'Noto Sans CJK TC', 'SimHei']  # 中文字体
    plt.rcParams['axes.unicode_minus'] = False  # 负号正常显示
    
    # 转换金额为万元
    pivot_display = pivot_df.copy()
    for col in pivot_display.columns:
        if col != "合计":
            pivot_display[col] = pivot_display[col] / 10000
    
    # 创建图形
    fig, ax = plt.subplots(figsize=(14, max(8, len(pivot_df) * 0.5 + 2)))
    ax.axis('off')
    
    # 获取表格尺寸
    n_rows = len(pivot_display) + 1  # +1 for header
    n_cols = len(pivot_display.columns)
    
    # 创建表格数据
    table_data = []
    col_labels = list(pivot_display.columns)
    
    for idx, row_name in enumerate(pivot_display.index):
        row_data = [f"{v:.2f}万" if isinstance(v, (int, float)) and not pd.isna(v) else str(v) for v in pivot_display.loc[row_name].values]
        table_data.append(row_data)
    
    # 使用matplotlib table
    table = ax.table(
        cellText=table_data,
        colLabels=col_labels,
        rowLabels=list(pivot_display.index),
        cellLoc='center',
        loc='center'
    )
    
    # 设置表格样式
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.3, 2.0)
    
    # 设置表头样式 (第一行)
    for col_idx in range(n_cols + 1):
        if (0, col_idx) in table._cells:
            cell = table._cells[(0, col_idx)]
            cell.set_facecolor('#4472C4')
            cell.set_text_props(color='white', fontweight='bold')
    
    # 设置行标签样式 (第一列)
    for row_idx in range(1, n_rows):
        if (row_idx, 0) in table._cells:
            cell = table._cells[(row_idx, 0)]
            cell.set_facecolor('#D9E2F3')
            cell.set_text_props(fontweight='bold')
    
    # 设置合计行样式 (最后一行)
    last_row = n_rows - 1
    for col_idx in range(n_cols + 1):
        if (last_row, col_idx) in table._cells:
            cell = table._cells[(last_row, col_idx)]
            cell.set_facecolor('#FFC000')
            cell.set_text_props(fontweight='bold')
    
    plt.title('发货计划透视表（金额单位：万元）', fontsize=14, fontweight='bold', pad=20)
    
    # 保存图片
    plt.savefig(img_path, dpi=150, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()
    print(f"透视表截图已保存到: {img_path}")

if __name__ == "__main__":
    df, pivot = main()
    print("\n=== 处理完成 ===")
