#!/usr/bin/env python3
"""
Microsoft邮件处理脚本
用于收发Microsoft账号的邮件
"""
import os
import concurrent
from typing import List, Dict
import requests
from loguru import logger
from flask import Flask, request, jsonify

# 自定义日志格式
logger.add("logfile.log", rotation="10 MB", format="{time:YYYY-MM-DD at HH:mm:ss} | {level} | {message}")

# 读取API_KEY
API_KEY = os.getenv('Api-Key','114514')
logger.info(f"API-KEY={API-KEY}")
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'

app = Flask(__name__)

def refresh_access_token(client_id, refresh_token) -> str:
    """刷新访问令牌"""
    refresh_params = {
        'client_id': client_id,
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token',
    }
    try:
        response = requests.post(TOKEN_URL, data=refresh_params)  # 去掉 get_proxy()
        response.raise_for_status()
        tokens = response.json()
        access_token = tokens['access_token']
        return access_token
    except requests.RequestException as e:
        logger.error(f"刷新访问令牌失败: {e}")
        raise

def get_messages(client_id, refresh_token, folder_id: str = 'inbox', top: int = 10) -> List[Dict]:
    """获取指定文件夹的邮件
    Args:
        folder_id: 文件夹ID, 默认为'inbox'
        top: 获取的邮件数量
    """
    access_token = refresh_access_token(client_id, refresh_token)
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json',
        'Prefer': 'outlook.body-content-type="text"'
    }
    query_params = {
        '$top': top,
        '$select': 'subject,receivedDateTime,from,body',
        '$orderby': 'receivedDateTime DESC'
    }
    try:
        response = requests.get(
            f'{GRAPH_API_ENDPOINT}/me/mailFolders/{folder_id}/messages',
            headers=headers,
            params=query_params  # 去掉 get_proxy()
        )
        response.raise_for_status()
        return response.json()['value']
    except requests.RequestException as e:
        logger.error(f"获取邮件失败: {e}")
        raise

def get_junk_messages(client_id, refresh_token, top: int = 10) -> List[Dict]:
    """获取垃圾邮件文件夹中的邮件"""
    return get_messages(client_id, refresh_token, folder_id='junkemail', top=top)

def delete_email(message_id, headers):
    """删除单封邮件的辅助函数"""
    delete_response = requests.delete(
        f'{GRAPH_API_ENDPOINT}/me/messages/{message_id}',
        headers=headers,
    )
    delete_response.raise_for_status()
    return delete_response  # 返回响应以供后续使用

def delete_all_emails(client_id, refresh_token, folder_id: str = 'inbox') -> None:
    """删除指定文件夹中的所有邮件
    Args:
        folder_id: 文件夹ID, 默认为'inbox'
    """
    access_token = refresh_access_token(client_id, refresh_token)
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json'
    }
    # 获取文件夹中的所有邮件
    try:
        response = requests.get(
            f'{GRAPH_API_ENDPOINT}/me/mailFolders/{folder_id}/messages',
            headers=headers,
            params={'$top': 1000},  # 一次获取最多1000封邮件
        )
        response.raise_for_status()
        messages = response.json()['value']
    except requests.RequestException as e:
        logger.error(f"获取邮件失败: {e}")
        raise
    # 逐个删除邮件
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {executor.submit(delete_email, message['id'], headers): message for message in messages}
        for future in concurrent.futures.as_completed(futures):
            message = futures[future]
            try:
                delete_response = future.result()  # 获取结果，触发异常（如果有）
                logger.info(f"已删除邮件: {message['id']}")
            except requests.RequestException as e:
                logger.error(f"删除邮件失败: {e}")
                raise

def delete_all_junk_emails(client_id, refresh_token) -> None:
    """删除垃圾邮件文件夹中的所有邮件"""
    delete_all_emails(client_id, refresh_token, folder_id='junkemail')

def delete_all_inbox_emails(client_id, refresh_token) -> None:
    """删除收件箱中的所有邮件"""
    delete_all_emails(client_id, refresh_token, folder_id='inbox')

def send_email(client_id, refresh_token, to_recipients: List[str], subject: str, content: str, is_html: bool = False) -> bool:
    """发送邮件
    Args:
        to_recipients: 收件人邮箱地址列表
        subject: 邮件主题
        content: 邮件内容
        is_html: 内容是否为HTML格式，默认为False
    Returns:
        bool: 发送是否成功
    """
    access_token = refresh_access_token(client_id, refresh_token)
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    email_msg = {
        'message': {
            'subject': subject,
            'body': {
                'contentType': 'HTML' if is_html else 'Text',
                'content': content
            },
            'toRecipients': [
                {
                    'emailAddress': {
                        'address': recipient
                    }
                } for recipient in to_recipients
            ]
        }
    }
    try:
        response = requests.post(
            f'{GRAPH_API_ENDPOINT}/me/sendMail',
            headers=headers,
            json=email_msg  # 去掉 get_proxy()
        )
        response.raise_for_status()
        logger.info(f"邮件已成功发送给 {', '.join(to_recipients)}")
        return True
    except requests.RequestException as e:
        logger.error(f"发送邮件失败: {e}")
        raise

def get_mail(client_id, refresh_token, subject_pattern=None, sender=None, folder_id: str = 'inbox', top=1):
    """
    获取符合条件的最新邮件内容
    Args:
        email (str): 邮箱地址
        password (str): 密码
        subject_pattern (str, optional): 主题匹配模式
        sender (str, optional): 发件人邮箱
        top (int, optional): 获取最新邮件数量，默认为1
    Returns:
        dict: 最新邮件的内容，包含主题、发件人、时间和正文
    """
    try:
        messages = get_messages(client_id, refresh_token, folder_id=folder_id, top=top)
        for msg in messages:
            # 检查主题是否匹配
            if subject_pattern and subject_pattern not in msg['subject']:
                continue
            # 检查发件人是否匹配
            if sender and sender != msg['from']['emailAddress']['address']:
                continue
            # 返回最新的符合条件的邮件内容
            return {
                'subject': msg['subject'],
                'sender': msg['from']['emailAddress']['address'],
                'time': msg['receivedDateTime'],
                'content': msg['body']['content']
            }
        # 如果没有找到符合条件的邮件，抛出带有具体信息的异常
        error_msg = "未找到符合条件的邮件"
        if subject_pattern:
            error_msg += f"，主题包含：{subject_pattern}"
        if sender:
            error_msg += f"，发件人：{sender}"
        raise ValueError(error_msg)
    except Exception as e:
        logger.error(f"获取邮件失败: {e}")
        raise

def get_code_or_link(client_id, refresh_token, subject_pattern=None, sender=None, folder_id='inbox', pattern=None, between=None):
    """ 从邮件内容中提取验证码、链接或位于特定字符之间的内容 """
    import re
    try:
        # 获取最新邮件
        mail_content = get_mail(client_id=client_id, refresh_token=refresh_token, subject_pattern=subject_pattern, sender=sender,
                                 folder_id=folder_id)
        content = str(mail_content['content'])
        logger.info('成功获取邮件')
        # 尝试提取
        if pattern:
            match = re.search(pattern, content)
            if match:
                return match.group()
        # 尝试提取位于字符a和b之间的内容
        if between:
            logger.info(f"夹中匹配: {between}")
            a, b = between.split(',')
            logger.info(f"{a}, {b}")
            between_pattern = rf'{re.escape(a.strip())}(.*?){re.escape(b.strip())}'  # 去除空格
            logger.info(f"生成的正则表达式: {between_pattern}")
            between_match = re.findall(between_pattern, content, flags=re.DOTALL)
            if between_match:
                return between_match
            else:
                logger.error("未找到匹配内容")
        # 如果都没找到返回False
        logger.error("都没找到")
        return False
    except Exception as e:
        logger.error(f"提取验证码、链接或内容失败: {e}")
        return False

@app.route('/send_email', methods=['POST'])
def api_send_email():
    """发送邮件接口"""
    api_key = request.headers.get('Api-Key')
    if api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json
    client_id = data.get('client_id')
    refresh_token = data.get('refresh_token')
    to_recipients = data.get('to_recipients')
    subject = data.get('subject')
    content = data.get('content')
    is_html = data.get('is_html', False)

    try:
        success = send_email(client_id, refresh_token, to_recipients, subject, content, is_html)
        return jsonify({"success": success}), 200
    except Exception as e:
        logger.error(f"发送邮件失败: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_messages', methods=['GET'])
def api_get_messages():
    """获取邮件接口"""
    api_key = request.headers.get('Api-Key')
    if api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    client_id = request.args.get('client_id')
    refresh_token = request.args.get('refresh_token')
    folder_id = request.args.get('folder_id', 'inbox')
    top = int(request.args.get('top', 10))

    try:
        messages = get_messages(client_id, refresh_token, folder_id, top)
        return jsonify(messages), 200
    except Exception as e:
        logger.error(f"获取邮件失败: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_mail', methods=['GET'])
def api_get_mail():
    """获取符合条件的最新邮件接口"""
    api_key = request.headers.get('Api-Key')
    if api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    client_id = request.args.get('client_id')
    refresh_token = request.args.get('refresh_token')
    subject_pattern = request.args.get('subject_pattern')
    sender = request.args.get('sender')
    folder_id = request.args.get('folder_id', 'inbox')
    top = int(request.args.get('top', 1))

    try:
        mail = get_mail(client_id, refresh_token, subject_pattern, sender, folder_id, top)
        return jsonify(mail), 200
    except Exception as e:
        logger.error(f"获取邮件失败: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_code_or_link', methods=['GET'])
def api_get_code_or_link():
    """从邮件内容中提取验证码、链接或内容接口"""
    api_key = request.headers.get('Api-Key')
    if api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    client_id = request.args.get('client_id')
    refresh_token = request.args.get('refresh_token')
    subject_pattern = request.args.get('subject_pattern')
    sender = request.args.get('sender')
    folder_id = request.args.get('folder_id', 'inbox')
    pattern = request.args.get('pattern')
    between = request.args.get('between')

    try:
        result = get_code_or_link(client_id, refresh_token, subject_pattern, sender, folder_id, pattern, between)
        return jsonify(result), 200
    except Exception as e:
        logger.error(f"提取验证码、链接或内容失败: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/delete_all_inbox_emails', methods=['DELETE'])
def api_delete_all_inbox_emails():
    """删除收件箱中的所有邮件接口"""
    api_key = request.headers.get('Api-Key')
    if api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    client_id = request.args.get('client_id')
    refresh_token = request.args.get('refresh_token')

    try:
        delete_all_inbox_emails(client_id, refresh_token)
        return jsonify({"success": True}), 200
    except Exception as e:
        logger.error(f"删除收件箱邮件失败: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/delete_all_junkemail_emails', methods=['DELETE'])
def api_delete_all_junkemail_emails():
    """删除收件箱中的所有邮件接口"""
    api_key = request.headers.get('Api-Key')
    if api_key != API_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    client_id = request.args.get('client_id')
    refresh_token = request.args.get('refresh_token')

    try:
        delete_all_junk_emails(client_id, refresh_token)
        return jsonify({"success": True}), 200
    except Exception as e:
        logger.error(f"删除收件箱邮件失败: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/', methods=['GET'])
def health_check():
    """健康度检测接口"""
    return jsonify({"status": "healthy"}), 200

if __name__ == '__main__':
    app.run(debug=True)
