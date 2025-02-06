import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

API_KEY = 'your_api_key'

# 设置重试机制
def create_session(retry=3):
    session = requests.Session()
    retries = Retry(total=retry, backoff_factor=2, status_forcelist=list(range(400, 600)))
    adapter = HTTPAdapter(max_retries=retries)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session

def send_email(client_id, refresh_token, to_recipients, subject, content, is_html=False):
    url = '/send_email'
    headers = {'Api-Key': API_KEY}
    data = {
        "client_id": client_id,
        "refresh_token": refresh_token,
        "to_recipients": to_recipients,
        "subject": subject,
        "content": content,
        "is_html": is_html
    }
    session = create_session()
    response = session.post(url, headers=headers, json=data)
    return response.json()

def get_messages(client_id, refresh_token, folder_id='inbox', top=10):
    url = '/get_messages'
    headers = {'Api-Key': API_KEY}
    params = {
        "client_id": client_id,
        "refresh_token": refresh_token,
        "folder_id": folder_id,
        "top": top
    }
    session = create_session()
    response = session.get(url, headers=headers, params=params)
    return response.json()

def get_mail(client_id, refresh_token, subject_pattern=None, sender=None, folder_id='inbox', top=1):
    url = '/get_mail'
    headers = {'Api-Key': API_KEY}
    params = {
        "client_id": client_id,
        "refresh_token": refresh_token,
        "subject_pattern": subject_pattern,
        "sender": sender,
        "folder_id": folder_id,
        "top": top
    }
    session = create_session()
    response = session.get(url, headers=headers, params=params)
    return response.json()

def get_code_or_link(client_id, refresh_token, subject_pattern=None, sender=None, folder_id='inbox', pattern=None, between=None):
    url = '/get_code_or_link'
    headers = {'Api-Key': API_KEY}
    params = {
        "client_id": client_id,
        "refresh_token": refresh_token,
        "subject_pattern": subject_pattern,
        "sender": sender,
        "folder_id": folder_id,
        "pattern": pattern,
        "between": between
    }
    session = create_session()
    response = session.get(url, headers=headers, params=params)
    return response.json()

def delete_all_inbox_emails(client_id, refresh_token):
    url = '/delete_all_inbox_emails'
    headers = {'Api-Key': API_KEY}
    params = {
        "client_id": client_id,
        "refresh_token": refresh_token
    }
    session = create_session()
    response = session.delete(url, headers=headers, params=params)
    return response.json()

def delete_all_junkemail_emails(client_id, refresh_token):
    url = '/delete_all_junkemail_emails'
    headers = {'Api-Key': API_KEY}
    params = {
        "client_id": client_id,
        "refresh_token": refresh_token
    }
    session = create_session()
    response = session.delete(url, headers=headers, params=params)
    return response.json()
