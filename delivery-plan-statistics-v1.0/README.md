# 未发货订单统计 Skill

一款专业的发货计划数据统计与可视化工具，支持邮件/消息/文件三种触发方式，自动生成精美的HTML报告和PNG截图。

## ✨ 功能特性

### 📊 核心功能
- **数据处理**：自动解析发货计划Excel，按月份、渠道统计汇总
- **可视化报告**：生成精美HTML报告，包含数据卡片、统计表格、柱状图
- **PNG截图**：自动生成精确尺寸（2178×2838）的截图，方便分享
- **动态数据**：所有数据从Excel动态读取，无硬编码

### 🎨 样式特点
- **蓝色渐变主题**：专业商务风格，100%匹配final模板
- **数据卡片**：4个核心指标一目了然
- **带边框表格**：黄色合计行，支持hover效果
- **柱状图数值标签**：柱子上方显示具体金额，清晰直观

### 🚀 三种触发方式
1. **邮件触发**：收到主题含"今日发货计划数据"的邮件，自动下载处理
2. **消息触发**：用户发送下载链接 + 关键词，触发统计
3. **文件触发**：直接上传本地Excel文件处理

## 📁 文件结构

```
未发货订单统计/
├── trigger.py           # 触发入口（三种触发方式）
├── process_data.py      # 数据处理脚本（不动）
├── generate_report.py   # 报告生成脚本（固化）
├── README.md            # 说明文档
├── 下载文件/            # 下载的原始数据
└── 处理结果/            # 生成的报告文件
    ├── 发货计划统计结果.xlsx
    ├── 发货计划统计报告_YYYYMMDD_HHMM.html
    └── 发货计划统计图_YYYYMMDD_HHMM.png
```

## 🚀 快速开始

### 方式1：文件触发（测试用）
```python
from trigger import trigger_by_file

result = trigger_by_file("最新发货计划.xlsx")
print(result)
```

命令行测试：
```bash
cd 未发货订单统计
python trigger.py test 最新发货计划.xlsx
```

### 方式2：消息触发
```python
from trigger import trigger_by_message

message = """请处理发货计划统计，数据在这里：
https://example.com/data.xlsx"""

result = trigger_by_message(message)
```

### 方式3：邮件触发
```python
from trigger import trigger_by_email

result = trigger_by_email(
    email_subject="今日发货计划数据",
    email_content="请查收今日数据：https://example.com/data.xlsx"
)
```

## 📋 输出说明

### 返回结果格式
```python
{
    "status": "success",
    "message": "文件触发统计完成",
    "trigger_type": "file",
    "file_path": "最新发货计划.xlsx",
    "result": {
        "excel_file": "处理结果/发货计划统计结果.xlsx",
        "html_file": "处理结果/发货计划统计报告_20260507_2157.html",
        "png_file": "处理结果/发货计划统计图_20260507_2157.png"
    }
}
```

### 输出文件说明

| 文件名 | 说明 |
|--------|------|
| 发货计划统计结果.xlsx | 处理后的统计数据，包含统计表和处理后的数据 |
| 发货计划统计报告_YYYYMMDD_HHMM.html | 完整HTML报告，支持浏览器打开 |
| 发货计划统计图_YYYYMMDD_HHMM.png | PNG截图，精确尺寸2178×2838像素 |

## 🎯 触发关键词

支持以下关键词触发（消息模式）：
- 发货计划统计
- 统计发货计划
- 未发货订单统计
- 未发货统计
- 发货计划
- 发货统计

## 📊 统计指标

### 4个核心数据卡片
1. **5月待发货计划金额**：从统计表读取
2. **5月待发货订单数**：从原始数据统计
3. **总未发货金额**：所有订单应收合计
4. **总未发货订单数**：所有订单数量

### 发货计划统计明细表格
- 按销售渠道分组统计
- 包含各月份（2026年04月、05月、06月、07月等）
- 包含"等通知"、"未客审"特殊状态
- 最后一行合计（黄色背景）

### 5月每日发货计划金额柱状图
- X轴：5月1日至31日
- Y轴：发货金额（万元）
- 数值标签：柱子上方显示具体金额（>0时）
- 渐变色柱子，圆角设计

## 🔧 技术说明

### 数据来源
**所有数据从`process_data.py`生成的Excel文件动态读取，无任何硬编码！**
- 数据卡片数值：从统计表sheet读取/计算
- 表格数据：从统计表sheet逐行读取
- 柱状图数据：从Sheet1_处理后sheet分组统计

### 技术栈
- **数据处理**：pandas
- **可视化**：ECharts 5.4.3（CDN）
- **截图生成**：Playwright + Chromium
- **尺寸控制**：viewport 1089×1419, deviceScaleFactor=2

### 样式规范
- PNG精确尺寸：2178×2838像素
- 四边紫色渐变背景（留白）
- 无边框设计
- 100%匹配final模板样式

## 📝 更新日志

### v2.0 (2026-05-07)
- ✅ 固化generate_report.py脚本
- ✅ 所有数据从Excel动态读取，无硬编码
- ✅ 100%匹配final模板样式
- ✅ 外部CDN引用ECharts
- ✅ 柱状图显示数值标签
- ✅ PNG精确尺寸控制
- ✅ 更新trigger.py调用新脚本
- ✅ 添加完整README文档

### v1.0
- 初始版本，基础功能实现

## 📞 常见问题

**Q: 为什么PNG尺寸不对？**
A: 检查container的宽度和padding设置，确保为1089px宽度

**Q: ECharts加载失败怎么办？**
A: 检查网络连接，确保能访问cdn.jsdelivr.net

**Q: 如何修改样式？**
A: 修改generate_report.py中的CSS部分，不要修改HTML结构和数据逻辑

**Q: 如何添加新的统计指标？**
A: 在generate_report.py的load_statistics_table和calculate_summary_stats函数中扩展

## 📄 许可证

本Skill内部使用，版权所有。
