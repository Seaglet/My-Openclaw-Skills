---
name: april-delivery-data
description: 获取4月发货数据，从WPS在线文档"2604发货"子表E43单元格读取汇总金额。当用户说"给我拉下4月发货数据"、"获取4月发货数据"时使用。
license: MIT
compatibility: Python 3.8+, agent-browser
metadata:
  version: "1.0"
  author: "幕僚长"
  created: "2026-04-28"
---

# 获取4月发货数据

自动从WPS在线文档获取4月发货汇总数据。

## 什么时候使用

- 用户说"给我拉下4月发货数据"
- 用户说"获取4月发货数据"
- 用户说"拉下4月发货"

## 文档信息

- 文档链接：https://365.kdocs.cn/l/ccfgfXOxcW7c?from=docs
- 目标工作表："2604发货"
- 目标单元格：E43（汇总列总计）

## 执行步骤（优化版 - 最快方式）

### 1. 打开文档（使用持久化session）
```bash
agent-browser --session-name wps-doc open "https://365.kdocs.cn/l/ccfgfXOxcW7c?from=docs" && agent-browser --session-name wps-doc tab 0 && agent-browser --session-name wps-doc wait --load networkidle
```

### 2. 使用JS直接读取E43（0.17秒）
```bash
agent-browser --session-name wps-doc eval 'WPSOpenApi.EtApplication().ActiveWorkbook.ActiveSheet.Range("E43").Value'
```

### 3. 只读操作
- 使用WPSOpenApi API，不操作DOM
- 不点击单元格，不进入编辑模式
- 不修改任何数据

## 输出格式

返回两个值：
- **原始值**：单元格显示的原始数字（带千分位分隔符）
- **万元值**：原始值 ÷ 10000，保留2位小数

示例：
```
原始值：4,927,540.6
万元值：492.75万元
```

## 注意事项

1. 数字可能有千分位分隔符，需要去除逗号后计算
2. 数字可能有小数位
3. 不要编辑表格，只读取数据
4. 如需登录，提示用户完成登录后继续

## 使用方式

脚本位于 `scripts/get_april_delivery.py`，可通过以下方式执行：

```bash
# 通过sub-agent执行
sessions_spawn 任务执行浏览器自动化
```
