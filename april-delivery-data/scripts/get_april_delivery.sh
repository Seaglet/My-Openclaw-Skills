#!/bin/bash
# 获取4月发货数据
# 使用 kdocs CLI 直接读取 WPS 在线表格 E43 单元格

set -e

# 配置
FILE_ID="avfCFBGrK1MRpvGy92VK1xXQsZJqEnfAp"
SHEET_ID=74
ROW_FROM=42
ROW_TO=42
COL_FROM=4
COL_TO=4

# 检查 Token
if [ -z "$KINGSOFT_DOCS_TOKEN" ]; then
    echo "错误: 未设置 KINGSOFT_DOCS_TOKEN 环境变量"
    echo "请先设置: export KINGSOFT_DOCS_TOKEN=<your_token>"
    exit 1
fi

# 执行读取
RESULT=$(kdocs-cli sheet get-range-data "{
  \"file_id\": \"$FILE_ID\",
  \"sheetId\": $SHEET_ID,
  \"range\": {
    \"rowFrom\": $ROW_FROM,
    \"rowTo\": $ROW_TO,
    \"colFrom\": $COL_FROM,
    \"colTo\": $COL_TO
  }
}" --silent 2>/dev/null)

# 解析结果
ORIGINAL=$(echo "$RESULT" | grep -o '"originalCellValue":"[^"]*"' | sed 's/"originalCellValue":"//;s/"//')
CELL_TEXT=$(echo "$RESULT" | grep -o '"cellText":"[^"]*"' | sed 's/"cellText":"//;s/"//')

# 计算万元值
WANYUAN=$(echo "scale=2; $ORIGINAL / 10000" | bc)

# 输出
echo "=========================================="
echo "4月发货数据汇总"
echo "=========================================="
echo "原始值：$ORIGINAL"
echo "显示值：$CELL_TEXT"
echo "万元值：$WANYUAN 万元"
echo "=========================================="
