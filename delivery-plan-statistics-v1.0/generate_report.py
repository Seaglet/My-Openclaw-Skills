#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发货计划统计报告生成脚本
功能：从process_data.py生成的Excel数据动态生成HTML报告和PNG截图
样式：100%匹配final模板
注意：所有数据均从Excel动态读取，无硬编码！
"""

import os
import sys
import json
import asyncio
from datetime import datetime
import pandas as pd
from playwright.async_api import async_playwright


def load_statistics_table(excel_path):
    """从Excel的统计表sheet读取统计数据"""
    print("正在读取统计表...")
    df = pd.read_excel(excel_path, sheet_name='统计表', header=1)
    
    # 清理列名
    df.columns = [str(col).strip() for col in df.columns]
    
    # 提取合计行
    total_row = df[df['销售渠道'] == '合计'].iloc[0]
    
    # 提取所有数据行用于表格
    stat_rows = df.to_dict('records')
    
    return total_row, stat_rows


def calculate_summary_stats(excel_path, total_row, current_month_str, current_month_num):
    """计算4个核心统计指标"""
    print("正在计算核心统计指标...")
    
    # 读取原始数据
    df = pd.read_excel(excel_path, sheet_name='Sheet1_处理后')
    
    # 1. 当月待发货计划金额（从统计表读取）
    may_amount = float(total_row.get(current_month_str, 0))
    
    # 2. 当月待发货订单数（从原始数据计算）
    df['日期'] = pd.to_datetime(df['订单-人工承诺发货时间'], errors='coerce')
    may_data = df[(df['日期'].dt.month == current_month_num) & (df['统计项目'] == current_month_str)]
    may_orders = len(may_data)
    
    # 3. 总未发货金额
    total_amount = df['应收合计'].sum() / 10000
    
    # 4. 总未发货订单数
    total_orders = len(df)
    
    return {
        'may_amount': round(may_amount, 2),
        'may_orders': may_orders,
        'total_amount': round(total_amount, 2),
        'total_orders': total_orders,
        'may_data': may_data
    }


def calculate_daily_amounts(may_data):
    """计算当月每日发货金额"""
    print("正在计算每日发货金额...")
    
    daily_summary = may_data.groupby(may_data['日期'].dt.day)['应收合计'].sum().reset_index()
    daily_summary.columns = ['日期', '金额(万元)']
    daily_summary['金额(万元)'] = daily_summary['金额(万元)'] / 10000
    daily_summary = daily_summary.sort_values('日期')
    
    daily_data = {int(row['日期']): round(row['金额(万元)'], 2) for _, row in daily_summary.iterrows()}
    
    # 构建完整的5月每日数据（1-31日）
    daily_amounts = []
    for day in range(1, 32):
        daily_amounts.append(daily_data.get(day, 0))
    
    return daily_amounts


def format_amount(val):
    """格式化金额显示"""
    if pd.isna(val) or val is None or val == 0 or val == '-' or str(val).strip() == '':
        return '-'
    try:
        return f"{float(val):.2f}万"
    except:
        return '-'


def generate_html(summary_stats, stat_rows, daily_amounts, output_dir, current_month_str, current_month_num):
    """生成HTML报告（所有数据动态传入）"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    report_date = datetime.now().strftime('%Y年%m月%d日')
    
    # 动态生成表格行
    table_rows = []
    for row in stat_rows:
        channel = row['销售渠道']
        row_class = 'total-row' if channel == '合计' else ''
        
        table_rows.append(f'''
            <tr class="{row_class}">
                <td class="channel-col">{channel}</td>
                <td>{format_amount(row.get('2026年04月'))}</td>
                <td>{format_amount(row.get(current_month_str))}</td>
                <td>{format_amount(row.get('2026年06月'))}</td>
                <td>{format_amount(row.get('2026年07月'))}</td>
                <td>{format_amount(row.get('2099年12月'))}</td>
                <td>{format_amount(row.get('等通知'))}</td>
                <td>{format_amount(row.get('未客审'))}</td>
                <td>{format_amount(row.get('合计'))}</td>
            </tr>
        ''')
    
    html_content = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>发货计划统计报告 - {report_date}</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 11px;
            margin: 0;
        }}
        
        .container {{
            width: calc(100% - 0px);
            margin: 0 auto;
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #4472C4 0%, #2F528F 100%);
            color: white;
            padding: 20px 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 24px;
            font-weight: 700;
            margin-bottom: 8px;
        }}
        
        .header .subtitle {{
            font-size: 12px;
            opacity: 0.9;
        }}
        
        .stats-overview {{
            display: flex;
            justify-content: space-between;
            gap: 15px;
            padding: 15px 25px;
            background: #f8f9fa;
        }}
        
        .stat-card {{
            text-align: center;
            padding: 14px 16px;
            background: white;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
            flex: 1;
        }}
        
        .stat-card .value {{
            font-size: 28px;
            font-weight: 700;
            color: #4472C4;
            margin-bottom: 5px;
        }}
        
        .stat-card .label {{
            font-size: 13px;
            color: #666;
        }}
        
        .section {{
            padding: 20px 30px;
        }}
        
        .section-title {{
            font-size: 18px;
            font-weight: 700;
            color: #333;
            margin-bottom: 18px;
            padding-bottom: 8px;
            border-bottom: 2px solid #4472C4;
            display: inline-block;
        }}
        
        .chart-container {{
            background: #fafbfc;
            border-radius: 10px;
            padding: 18px;
            margin-bottom: 25px;
        }}
        
        #dailyChart {{
            width: 100%;
            height: 400px;
        }}
        
        .data-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            margin-top: 15px;
        }}
        
        .data-table th {{
            background: #4472C4;
            color: white;
            padding: 10px 8px;
            text-align: center;
            font-weight: 600;
            border: 1px solid #3a5f9e;
        }}
        
        .data-table td {{
            padding: 8px;
            border: 1px solid #e0e0e0;
            text-align: right;
        }}
        
        .data-table .channel-col {{
            text-align: left;
            font-weight: 600;
            background: #f8f9fa;
        }}
        
        .data-table tr:hover {{
            background: #f0f7ff;
        }}
        
        .data-table .total-row {{
            background: #FFF2CC;
            font-weight: 700;
        }}
        
        .data-table .total-row:hover {{
            background: #FFE699;
        }}
        
        .footer {{
            text-align: center;
            padding: 15px;
            background: #f8f9fa;
            color: #999;
            font-size: 11px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 发货计划统计报告</h1>
            <div class="subtitle">统计日期：{report_date} | 数据来源：未发货订单统计</div>
        </div>
        
        <div class="stats-overview">
            <div class="stat-card">
                <div class="value">{summary_stats['may_amount']}</div>
                <div class="label">{datetime.now().month}月待发货计划金额（万元）</div>
            </div>
            <div class="stat-card">
                <div class="value">{summary_stats['may_orders']}</div>
                <div class="label">{datetime.now().month}月待发货订单数</div>
            </div>
            <div class="stat-card">
                <div class="value">{summary_stats['total_amount']}</div>
                <div class="label">总未发货金额（万元）</div>
            </div>
            <div class="stat-card">
                <div class="value">{summary_stats['total_orders']}</div>
                <div class="label">总未发货订单数</div>
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">📊 发货计划统计明细</h2>
            <table class="data-table">
                <thead>
                    <tr>
                        <th>销售渠道</th>
                        <th>2026年04月</th>
                        <th>{current_month_str}</th>
                        <th>2026年06月</th>
                        <th>2026年07月</th>
                        <th>2099年12月</th>
                        <th>等通知</th>
                        <th>未客审</th>
                        <th>合计</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(table_rows)}
                </tbody>
            </table>
        </div>
        <div class="section">
            <h2 class="section-title">📈 {current_month_str}每日发货计划金额</h2>
            <div class="chart-container">
                <div id="dailyChart"></div>
            </div>
        </div>
        
        <div class="footer">
            <p>© 2026 发货计划统计系统 | 本报告由系统自动生成，数据仅供参考</p>
        </div>
    </div>
    
    <script>
        var chartDom = document.getElementById('dailyChart');
        var myChart = echarts.init(chartDom);
        
        var days = [];
        for(var d = 1; d <= 31; d++) {{
            var monthStr = ('0' + {current_month_num}).slice(-2);
            var dayStr = ('0' + d).slice(-2);
            days.push(monthStr + '月' + dayStr + '日');
        }}
        var amounts = {json.dumps(daily_amounts)};
        
        var option = {{
            tooltip: {{
                trigger: 'axis',
                formatter: function(params) {{
                    return '<b>' + params[0].name + '</b><br/>发货金额: ' + params[0].value.toFixed(2) + ' 万元';
                }}
            }},
            grid: {{
                left: '3%',
                right: '4%',
                bottom: '15%',
                containLabel: true
            }},
            xAxis: {{
                type: 'category',
                data: days,
                name: '日期',
                nameLocation: 'middle',
                nameGap: 60,
                axisLabel: {{
                    interval: 0,
                    rotate: 45,
                    fontSize: 11
                }}
            }},
            yAxis: {{
                type: 'value',
                name: '发货金额（万元）',
                nameLocation: 'middle',
                nameGap: 50
            }},
            series: [{{
                data: amounts,
                type: 'bar',
                itemStyle: {{
                    color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                        {{offset: 0, color: '#667eea'}},
                        {{offset: 1, color: '#4472C4'}}
                    ]),
                    borderRadius: [4, 4, 0, 0]
                }},
                label: {{
                    show: true,
                    position: 'top',
                    formatter: function(params) {{
                        if (params.value > 0) {{
                            return params.value.toFixed(1) + '万';
                        }}
                        return '';
                    }},
                    fontSize: 10,
                    color: '#333'
                }}
            }}]
        }};
        
        myChart.setOption(option);
        
        window.addEventListener('resize', function() {{
            myChart.resize();
        }});
    </script>
</body>
</html>'''
    
    html_path = os.path.join(output_dir, f'发货计划统计报告_{timestamp}.html')
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✅ HTML报告已生成: {html_path}")
    return html_path, timestamp


async def generate_png(html_path, output_dir, timestamp):
    """生成PNG截图（精确尺寸2178×2838）"""
    print("正在生成PNG截图...")
    
    png_path = os.path.join(output_dir, f'发货计划统计图_{timestamp}.png')
    html_url = f'file://{os.path.abspath(html_path)}'
    
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={'width': 1089, 'height': 1383}, device_scale_factor=2)  # 精确到PNG高度2838
        
        await page.goto(html_url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(8)  # 等待ECharts加载和渲染
        
        await page.screenshot(path=png_path, full_page=True)
        await browser.close()
    
    print(f"✅ PNG截图已生成: {png_path}")
    
    # 验证尺寸
    from PIL import Image
    img = Image.open(png_path)
    expected = (2178, 2838)
    if img.size == expected:
        print(f"✅ PNG尺寸验证通过: {img.size}")
    else:
        print(f"⚠️ PNG尺寸: {img.size} (预期: {expected})")
    
    return png_path


def main(excel_path, output_dir=None):
    """主函数：所有数据从Excel动态读取"""
    from datetime import datetime
    
    # 动态获取当前年月
    current_month_str = datetime.now().strftime('%Y年%m月')
    current_month_num = datetime.now().month
    
    print("=" * 50)
    print("发货计划统计报告生成器")
    print(f"统计月份：{current_month_str}")
    print("=" * 50)
    
    if output_dir is None:
        output_dir = os.path.dirname(excel_path)  # 直接用excel所在目录，不嵌套
    
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. 从统计表sheet读取统计数据
    total_row, stat_rows = load_statistics_table(excel_path)
    
    # 2. 计算核心统计指标（4个数据卡片）
    summary_stats = calculate_summary_stats(excel_path, total_row, current_month_str, current_month_num)
    
    # 3. 计算每日发货金额（柱状图数据）
    daily_amounts = calculate_daily_amounts(summary_stats['may_data'])
    
    # 4. 生成HTML（所有数据动态传入）
    html_path, timestamp = generate_html(summary_stats, stat_rows, daily_amounts, output_dir, current_month_str, current_month_num)
    
    # 5. 生成PNG
    png_path = asyncio.run(generate_png(html_path, output_dir, timestamp))
    
    print("\n" + "=" * 50)
    print("✅ 报告生成完成！")
    print(f"📊 {datetime.now().month}月待发货金额: {summary_stats['may_amount']}万元")
    print(f"📦 {datetime.now().month}月待发货订单: {summary_stats['may_orders']}单")
    print(f"💰 总未发货金额: {summary_stats['total_amount']}万元")
    print(f"📋 总未发货订单: {summary_stats['total_orders']}单")
    print("=" * 50)
    print(f"HTML报告: {html_path}")
    print(f"PNG截图: {png_path}")
    print("=" * 50)
    
    return html_path, png_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python generate_report.py <Excel文件路径> [输出目录]")
        print("说明: Excel文件必须是process_data.py处理后的结果")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    main(excel_path, output_dir)
