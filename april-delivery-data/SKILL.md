---
name: april-delivery-data
description: 获取4月发货数据，从WPS在线文档"2604发货"子表E43单元格读取汇总金额。当用户说"给我拉下4月发货数据"、"获取4月发货数据"时使用。
license: MIT
compatibility: kdocs-cli 2.4.3+
metadata:
  version: "2.0"
  author: "幕僚长"
  created: "2026-04-28"
  updated: "2026-04-28"
---

# 获取4月发货数据

自动从WPS在线文档获取4月发货汇总数据，使用 kdocs CLI 直接读取，秒级完成。

## 什么时候使用

- 用户说"给我拉下4月发货数据"
- 用户说"获取4月发货数据"
- 用户说"拉下4月发货"

## 文档信息

- 文档链接：https://365.kdocs.cn/l/ccfgfXOxcW7c?from=docs
- 链接ID：ccfgfXOxcW7c
- 目标工作表："2604发货" (sheetId: 74)
- 目标单元格：E43（汇总列总计）
- 行列索引：rowFrom=42, rowTo=42, colFrom=4, colTo=4（0-based）

## 执行步骤（kdocs CLI 方式 - 最快 ~2秒）

### 前置条件

确保已安装 kdocs-cli 并设置 Token：
```bash
# 安装后验证
kdocs-cli version

# 设置 Token（如未设置）
kdocs-cli auth set-token <TOKEN>
```

### 执行命令

```bash
# 读取 E43 单元格
export PATH="/root/.local/bin:$PATH"
export KINGSOFT_DOCS_TOKEN="<TOKEN>"

kdocs-cli sheet get-range-data '{
  "file_id": "avfCFBGrK1MRpvGy92VK1xXQsZJqEnfAp",
  "sheetId": 74,
  "range": {
    "rowFrom": 42,
    "rowTo": 42,
    "colFrom": 4,
    "colTo": 4
  }
}'
```

### 返回值解析

```json
{
  "code": 0,
  "data": {
    "detail": {
      "rangeData": [{
        "cellText": "4,927,541",           // 显示值（千分位格式）
        "originalCellValue": "4927540.6",  // 原始值
        "fmlaText": "=E19+E28+E32+E42",    // 公式（如有）
        "understandableType": {
          "type": "double",
          "value": 4927540.6               // 数值类型
        }
      }]
    }
  }
}
```

## 输出格式

返回两个值：
- **原始值**：单元格的原始数字
- **万元值**：原始值 ÷ 10000，保留2位小数

示例：
```
原始值：4,927,540.6
万元值：492.75万元
```

## 性能对比

| 方式 | 耗时 |
|------|------|
| 浏览器截图+滚动+识别 | ~5分钟 |
| 持久化session+点击 | 20-30秒 |
| 持久化session+JS API | ~15秒 |
| **kdocs CLI** | **~2秒** ✅ |

## 注意事项

1. 需要有效的 KINGSOFT_DOCS_TOKEN
2. 行列索引为 0-based（E列=4，第43行=42）
3. 只读操作，不修改任何数据
4. Token 过期时需重新获取

## 故障排查

- `400006` 错误：Token 过期，需重新设置
- `400001` 错误：参数格式错误，检查 JSON 格式
- 文件不存在：检查 file_id 是否正确

## 相关文件

- 脚本：`scripts/get_april_delivery.sh`
