# 未发货订单统计 Skill

一键完成未发货订单的数据清洗、透视表生成、统计计算、可视化图表生成的专业Skill。

## Skill 信息

| 项目 | 说明 |
|------|------|
| **Skill名称** | 未发货订单统计 |
| **版本** | v1.0.0 |
| **功能** | 数据清洗、透视表生成、统计表计算、ECharts可视化、PNG导出 |
| **依赖** | pandas, openpyxl, playwright(可选) |

## 快速开始

### 安装依赖

```bash
pip install pandas openpyxl playwright
playwright install chromium  # 可选，用于PNG截图
```

### 一行代码调用

```python
from 未发货订单统计 import run_unshipped_order_statistics

# 执行完整统计流程
result = run_unshipped_order_statistics("原始发货数据.xlsx")

# 获取结果路径
print(f"Excel文件: {result['excel_file']}")
print(f"HTML报告: {result['html_file']}")
print(f"PNG图片: {result['png_file']}")
```

## 功能模块

### 1. 数据处理 (process_data.py)

**核心函数：** `process_data(input_file, output_dir)`

**功能列表：**
- ✅ 自动识别Excel列名（订单、日期、金额、渠道、仓库）
- ✅ 数据清洗与标准化（去除空格、格式转换）
- ✅ 日期格式统一处理
- ✅ 金额单位自动转换（元 → 万元）
- ✅ 生成交叉透视表（销售渠道 x 发货月份）
- ✅ 生成统计表（含统计概览）
- ✅ 输出三工作表Excel文件

**输出工作表：**
1. **Sheet1_处理后** - 清洗后的原始数据
2. **透视表** - 按销售渠道和发货月份的交叉汇总
3. **统计表** - 统计概览 + 详细数据

### 2. 可视化生成 (generate_chart.py)

**核心函数：** `generate_chart(input_file, output_dir)`

**功能列表：**
- ✅ 读取Excel统计表数据
- ✅ 使用template.html模板注入数据
- ✅ 生成ECharts柱状图（5月每日发货金额）
- ✅ 生成数据卡片（总金额、订单数等）
- ✅ 导出交互式HTML报告
- ✅ Playwright转高清PNG截图（可选）

### 3. HTML模板 (template.html)

**动态注入点：**
- `{{TITLE}}` - 报告标题
- `{{SUBTITLE}}` - 报告副标题
- `{{STATS_CARDS}}` - 顶部数据卡片
- `{{DAYS_JSON}}` - X轴日期数据
- `{{AMOUNTS_JSON}}` - Y轴金额数据
- `{{TABLE_HTML}}` - 统计表HTML

**设计特点：**
- 渐变紫色背景
- 精美的卡片式布局
- ECharts可视化图表
- 响应式设计
- 合计行高亮显示

## 模块架构

```
未发货订单统计/
├── __init__.py          # Skill入口，暴露主函数
├── process_data.py      # 数据处理模块
├── generate_chart.py    # 可视化模块
├── template.html        # HTML图表模板
├── SKILL.md             # Skill说明文档（本文件）
├── README.md            # 用户使用说明
└── 处理结果/            # 输出目录
    ├── 发货计划统计结果.xlsx
    ├── 发货计划统计报告.html
    ├── 发货计划统计图.png
    └── chart_data.json
```

## 独立脚本运行

除了作为Skill导入，也可以独立运行脚本：

```bash
# 第一步：数据处理
python process_data.py ../原始数据.xlsx

# 第二步：可视化生成
python generate_chart.py 处理结果/发货计划统计结果.xlsx
```

## 数据列自动识别

Skill会自动识别以下列名（支持模糊匹配）：

| 字段 | 匹配关键词 |
|------|----------|
| 订单号 | 订单号、订单编号、单号 |
| 日期 | 付款时间、承诺发货时间、发货时间、时间 |
| 金额 | 实付金额、应收合计、金额、货款 |
| 销售渠道 | 销售渠道、渠道、来源 |
| 仓库 | 发货仓库、仓库 |

## 常见问题

### Q: Playwright安装失败怎么办？

A: Playwright仅用于PNG截图，不影响核心功能。HTML报告可直接在浏览器打开查看。

```bash
pip install playwright
playwright install chromium
```

### Q: 我的Excel列名不匹配怎么办？

A: Skill支持模糊匹配，只要列名包含关键词即可识别。如果完全不匹配，会使用默认值处理。

### Q: 如何自定义输出目录？

A: 在调用时指定output_dir参数：

```python
result = run_unshipped_order_statistics("数据.xlsx", output_dir="我的输出目录")
```

## 更新日志

### v1.0.0 (2026-05-07)
- 🎉 完成Skill封装
- ✅ 支持一键调用 `run_unshipped_order_statistics()`
- ✅ 双模块架构（数据处理 + 可视化）
- ✅ 自动识别Excel列名
- ✅ ECharts可视化图表
- ✅ Playwright高清PNG导出
- ✅ 完整的SKILL.md文档

## 使用示例

```python
# 导入Skill
from 未发货订单统计 import run_unshipped_order_statistics

# 处理发货数据
result = run_unshipped_order_statistics(
    input_excel_path="发货计划.xlsx",
    output_dir="处理结果"
)

# 使用结果
print(f"✅ 处理完成！")
print(f"📊 Excel统计文件: {result['excel_file']}")
print(f"📈 HTML可视化报告: {result['html_file']}")
if result['png_file']:
    print(f"🖼️  PNG图片文件: {result['png_file']}")
```
