#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
未发货订单统计 - 触发入口脚本
支持三种触发方式：
1. 邮件触发：收到主题包含"今日发货计划数据"的邮件
2. 消息触发：用户发送下载链接 + 关键词
3. 文件触发：用户直接上传Excel文件

流程：接收数据 -> process_data处理 -> generate_report生成HTML+PNG
"""

import os
import re
import requests
import urllib.parse
from datetime import datetime

# 导入主处理函数，避免循环导入
def _run_unshipped_order_statistics(input_excel_path, output_dir="处理结果"):
    """包装主函数，执行完整流程：数据处理 -> 报告生成"""
    try:
        from .process_data import process_data
        from .generate_report import main as generate_report
    except ImportError:
        import sys
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from process_data import process_data
        from generate_report import main as generate_report
    
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
    
    # 第二步：生成报告（HTML + PNG）
    print("\n🎨 步骤 2/2: 生成统计报告中...")
    html_path, png_path = generate_report(excel_result, output_dir)
    
    print("\n" + "=" * 70)
    print("✅ 未发货订单统计 - 全部完成！")
    print("=" * 70)
    print(f"\n📊 处理结果:")
    print(f"   Excel数据: {excel_result}")
    print(f"   HTML报告: {html_path}")
    print(f"   PNG截图:  {png_path}")
    print()
    
    # 返回结果
    return {
        "excel_file": excel_result,
        "html_file": html_path,
        "png_file": png_path
    }


# 触发关键词列表
TRIGGER_KEYWORDS = [
    "发货计划统计",
    "统计发货计划",
    "未发货订单统计",
    "未发货统计",
    "发货计划",
    "发货统计"
]

# 邮件触发关键词
EMAIL_TRIGGER_SUBJECT = "今日发货计划数据"

# 下载目录
DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "下载文件")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


def extract_urls(text):
    """
    从文本中提取所有URL链接
    
    Args:
        text: 文本内容（邮件正文、消息内容等）
    
    Returns:
        list: 提取到的URL列表
    """
    # URL正则表达式模式
    url_pattern = r'https?://[^\s<>"\'）)】]+'
    urls = re.findall(url_pattern, text)
    
    # 清理URL末尾可能带有的标点符号
    cleaned_urls = []
    for url in urls:
        # 移除末尾的标点符号
        url = re.sub(r'[.,;:!?，。；：！？）)】]+$', '', url)
        cleaned_urls.append(url)
    
    return cleaned_urls


def download_file(url, save_dir=DOWNLOAD_DIR):
    """
    从URL下载文件
    
    Args:
        url: 下载链接
        save_dir: 保存目录
    
    Returns:
        str: 下载后的文件路径，失败返回None
    """
    try:
        print(f"正在下载: {url}")
        
        # 设置请求头，模拟浏览器
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        # 发送请求
        response = requests.get(url, headers=headers, timeout=60)
        response.raise_for_status()  # 检查请求是否成功
        
        # 从URL或响应头获取文件名
        filename = None
        
        # 尝试从Content-Disposition获取文件名
        if 'Content-Disposition' in response.headers:
            cd = response.headers['Content-Disposition']
            match = re.search(r'filename[*]?=["\']?([^"\'；]+)["\']?', cd)
            if match:
                filename = match.group(1)
                # 处理URL编码的文件名
                if filename.startswith("UTF-8''"):
                    filename = urllib.parse.unquote(filename[7:])
        
        # 如果没有从响应头获取到，从URL中提取
        if not filename:
            parsed_url = urllib.parse.urlparse(url)
            filename = os.path.basename(parsed_url.path)
            if not filename or '.' not in filename:
                # 使用默认文件名
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"发货计划数据_{timestamp}.xlsx"
        
        # 确保文件名安全
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # 保存文件
        save_path = os.path.join(save_dir, filename)
        
        with open(save_path, 'wb') as f:
            f.write(response.content)
        
        print(f"下载完成: {save_path}")
        print(f"文件大小: {os.path.getsize(save_path) / 1024:.2f} KB")
        
        return save_path
        
    except requests.exceptions.RequestException as e:
        print(f"下载失败: {str(e)}")
        return None
    except Exception as e:
        print(f"保存文件时出错: {str(e)}")
        return None


def trigger_by_email(email_subject, email_content):
    """
    邮件触发入口
    
    检查邮件主题是否包含"今日发货计划数据"，
    从邮件内容提取下载链接，下载Excel文件并执行统计。
    
    Args:
        email_subject: 邮件主题
        email_content: 邮件正文内容
    
    Returns:
        dict: 执行结果，包含状态和结果信息
    """
    print("=" * 70)
    print("📧 邮件触发模式")
    print("=" * 70)
    
    # 检查邮件主题
    if EMAIL_TRIGGER_SUBJECT not in str(email_subject):
        return {
            "status": "skipped",
            "message": f"邮件主题不包含触发关键词 '{EMAIL_TRIGGER_SUBJECT}'",
            "email_subject": email_subject
        }
    
    print(f"\n✅ 检测到触发邮件主题: {email_subject}")
    
    # 提取下载链接
    urls = extract_urls(email_content)
    
    if not urls:
        return {
            "status": "error",
            "message": "邮件中未检测到下载链接",
            "email_subject": email_subject
        }
    
    print(f"\n📎 检测到 {len(urls)} 个下载链接")
    
    # 尝试下载第一个Excel链接
    downloaded_file = None
    for url in urls:
        # 优先尝试看起来像是Excel的链接
        if any(ext in url.lower() for ext in ['.xlsx', '.xls', '.csv']):
            downloaded_file = download_file(url)
            if downloaded_file:
                break
        else:
            # 尝试所有链接
            downloaded_file = download_file(url)
            if downloaded_file:
                break
    
    if not downloaded_file:
        return {
            "status": "error",
            "message": "所有下载链接均下载失败",
            "email_subject": email_subject,
            "urls_found": urls
        }
    
    # 执行统计
    try:
        output_dir = os.path.join(os.path.dirname(__file__), "处理结果")
        result = _run_unshipped_order_statistics(downloaded_file, output_dir)
        
        return {
            "status": "success",
            "message": "邮件触发统计完成",
            "trigger_type": "email",
            "email_subject": email_subject,
            "downloaded_file": downloaded_file,
            "result": result
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"统计执行失败: {str(e)}",
            "trigger_type": "email",
            "email_subject": email_subject,
            "downloaded_file": downloaded_file
        }


def trigger_by_message(message_text):
    """
    消息触发入口
    
    检查消息是否包含关键词，从消息中提取下载链接，
    下载Excel文件并执行统计。
    
    Args:
        message_text: 用户消息内容
    
    Returns:
        dict: 执行结果，包含状态和结果信息
    """
    print("=" * 70)
    print("💬 消息触发模式")
    print("=" * 70)
    
    # 检查是否包含触发关键词
    has_keyword = False
    matched_keyword = None
    for keyword in TRIGGER_KEYWORDS:
        if keyword in str(message_text):
            has_keyword = True
            matched_keyword = keyword
            break
    
    if not has_keyword:
        return {
            "status": "skipped",
            "message": "消息不包含触发关键词",
            "keywords_checked": TRIGGER_KEYWORDS
        }
    
    print(f"\n✅ 检测到触发关键词: {matched_keyword}")
    
    # 提取下载链接
    urls = extract_urls(message_text)
    
    if not urls:
        # 没有链接，提示用户提供链接
        return {
            "status": "need_url",
            "message": "检测到触发关键词，但未找到下载链接。请提供发货计划数据的下载链接。",
            "trigger_type": "message",
            "matched_keyword": matched_keyword
        }
    
    print(f"\n📎 检测到 {len(urls)} 个下载链接")
    
    # 尝试下载第一个Excel链接
    downloaded_file = None
    for url in urls:
        if any(ext in url.lower() for ext in ['.xlsx', '.xls', '.csv']):
            downloaded_file = download_file(url)
            if downloaded_file:
                break
        else:
            downloaded_file = download_file(url)
            if downloaded_file:
                break
    
    if not downloaded_file:
        return {
            "status": "error",
            "message": "所有下载链接均下载失败",
            "trigger_type": "message",
            "matched_keyword": matched_keyword,
            "urls_found": urls
        }
    
    # 执行统计
    try:
        output_dir = os.path.join(os.path.dirname(__file__), "处理结果")
        result = _run_unshipped_order_statistics(downloaded_file, output_dir)
        
        return {
            "status": "success",
            "message": "消息触发统计完成",
            "trigger_type": "message",
            "matched_keyword": matched_keyword,
            "downloaded_file": downloaded_file,
            "result": result
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"统计执行失败: {str(e)}",
            "trigger_type": "message",
            "matched_keyword": matched_keyword,
            "downloaded_file": downloaded_file
        }


def trigger_by_file(file_path):
    """
    文件触发入口
    
    直接处理用户上传的Excel文件。
    
    Args:
        file_path: 用户上传的Excel文件路径
    
    Returns:
        dict: 执行结果，包含状态和结果信息
    """
    print("=" * 70)
    print("📁 文件触发模式")
    print("=" * 70)
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        return {
            "status": "error",
            "message": f"文件不存在: {file_path}",
            "trigger_type": "file"
        }
    
    # 检查文件格式
    file_ext = os.path.splitext(file_path)[1].lower()
    if file_ext not in ['.xlsx', '.xls', '.csv']:
        return {
            "status": "error",
            "message": f"不支持的文件格式: {file_ext}，仅支持 .xlsx, .xls, .csv",
            "trigger_type": "file",
            "file_path": file_path
        }
    
    print(f"\n✅ 接收到数据文件: {file_path}")
    print(f"   文件格式: {file_ext}")
    print(f"   文件大小: {os.path.getsize(file_path) / 1024:.2f} KB")
    
    # 执行统计
    try:
        output_dir = os.path.join(os.path.dirname(__file__), "处理结果")
        result = _run_unshipped_order_statistics(file_path, output_dir)
        
        return {
            "status": "success",
            "message": "文件触发统计完成",
            "trigger_type": "file",
            "file_path": file_path,
            "result": result
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"统计执行失败: {str(e)}",
            "trigger_type": "file",
            "file_path": file_path
        }


if __name__ == '__main__':
    # 测试示例
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == 'test':
        # 测试模式：传入本地Excel文件路径
        test_file = sys.argv[2] if len(sys.argv) > 2 else "最新发货计划.xlsx"
        result = trigger_by_file(test_file)
        print(f"\n测试结果: {result}")
    else:
        print("未发货订单统计 - 触发入口")
        print("-" * 50)
        print("支持三种触发方式：")
        print("  1. trigger_by_email(email_subject, email_content)")
        print("  2. trigger_by_message(message_text)")
        print("  3. trigger_by_file(file_path)")
        print("\n测试命令: python trigger.py test <Excel文件路径>")


def trigger_by_email(email_subject, email_body):
    """
    邮件触发方式
    如果邮件主题或正文包含"统计发货计划"字样，自动提取下载链接并触发任务
    
    Args:
        email_subject: 邮件主题
        email_body: 邮件正文
        
    Returns:
        处理结果字典
    """
    # 检查是否包含触发关键词
    trigger_keywords = ["统计发货计划", "发货计划统计", "未发货订单统计"]
    has_trigger = any(keyword in email_subject or keyword in email_body for keyword in trigger_keywords)
    
    if not has_trigger:
        return {"success": False, "message": "不包含触发关键词，跳过处理"}
    
    # 提取下载链接
    import re
    url_pattern = r'https?://[^\s<>"\']+\.xlsx[^\s<>"\']*'
    urls = re.findall(url_pattern, email_body)
    
    if not urls:
        return {"success": False, "message": "未找到Excel下载链接"}
    
    # 使用第一个链接处理
    download_url = urls[0]
    return trigger_by_message(f"发货计划统计：{download_url}")


# 邮件处理入口函数（供邮件分身调用）
def process_email_trigger(email_subject, email_body):
    """
    邮件分身调用此函数处理邮件
    
    重要：处理邮件时不回复邮件，直接处理数据
    """
    result = trigger_by_email(email_subject, email_body)
    return result
