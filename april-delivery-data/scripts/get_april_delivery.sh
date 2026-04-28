#!/bin/bash
# 任务名：获取4月发货数据
# 从WPS在线文档"2604发货"子表E43单元格获取汇总金额

# 文档链接
DOC_URL="https://365.kdocs.cn/l/ccfgfXOxcW7c?from=docs"

# Session名称（保持登录状态）
SESSION_NAME="wps-doc"

# 截图目录
SCREENSHOT_DIR="/app/data/所有对话/主对话/browser/screenshots/april-delivery"
mkdir -p "$SCREENSHOT_DIR"

echo "=== 开始获取4月发货数据 ==="

# 1. 打开文档（使用持久化session，保持登录状态）
echo "步骤1: 打开WPS文档（使用持久化session）..."
agent-browser --session-name "$SESSION_NAME" open "$DOC_URL" && agent-browser --session-name "$SESSION_NAME" tab 0
agent-browser --session-name "$SESSION_NAME" wait --load networkidle

# 2. 截图查看当前页面
echo "步骤2: 获取页面快照..."
agent-browser screenshot "$SCREENSHOT_DIR/page1.png"

# 3. 获取页面元素
echo "步骤3: 分析页面元素..."
agent-browser snapshot -i

# 4. 切换到"2604发货"子表（如果需要）
# 这一步需要根据snapshot结果判断是否需要切换

# 5. 滚动到表格底部（E43位置）
echo "步骤4: 滚动到表格底部..."
for i in {1..10}; do
    agent-browser scroll down 500
    sleep 1
done

# 6. 截图查看E43位置
echo "步骤5: 截图获取E43单元格..."
agent-browser screenshot "$SCREENSHOT_DIR/e43.png"

# 7. 获取E43单元格值
echo "步骤6: 读取E43单元格值..."
# 需要根据页面元素定位E43，这里需要人工确认或通过snapshot找到对应元素

echo "=== 任务完成 ==="
echo "请查看截图: $SCREENSHOT_DIR/e43.png"

# 注意：完整自动化需要根据页面元素动态获取E43的值
# 当前脚本提供了基本框架，实际执行时需要结合snapshot结果定位元素
