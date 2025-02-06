以下是一个简单的 README 文件示例，您可以根据需要进行调整：
# Microsoft邮件处理脚本

## 简介
该脚本用于收发 Microsoft 账号的邮件，支持获取、发送、删除邮件等功能。使用 Flask 框架构建 RESTful API 接口，方便与其他应用进行集成。

## 功能
- 获取指定文件夹的邮件
- 发送邮件
- 删除指定文件夹中的所有邮件
- 从邮件中提取验证码、链接或特定内容

## 环境要求
- Python 3.x
- Flask
- requests
- loguru

## 安装依赖
在项目目录下运行以下命令安装所需依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

### 配置
在运行脚本之前，请确保设置了环境变量 `Api-Key`，用于 API 访问的身份验证。

### 启动服务
运行以下命令启动 Flask 服务：
```bash
python script.py
```

### API 接口

#### 发送邮件
- **URL**: `/send_email`
- **方法**: `POST`
- **请求体**:
  ```json
  {
      "client_id": "your_client_id",
      "refresh_token": "your_refresh_token",
      "to_recipients": ["recipient@example.com"],
      "subject": "邮件主题",
      "content": "邮件内容",
      "is_html": false
  }
  ```

#### 获取邮件
- **URL**: `/get_messages`
- **方法**: `GET`
- **请求参数**:
  - `client_id`: 客户端ID
  - `refresh_token`: 刷新令牌
  - `folder_id`: 文件夹ID（默认为'inbox'）
  - `top`: 获取的邮件数量（默认为10）

#### 获取符合条件的最新邮件
- **URL**: `/get_mail`
- **方法**: `GET`
- **请求参数**:
  - `client_id`: 客户端ID
  - `refresh_token`: 刷新令牌
  - `subject_pattern`: 主题匹配模式（可选）
  - `sender`: 发件人邮箱（可选）
  - `folder_id`: 文件夹ID（默认为'inbox'）
  - `top`: 获取最新邮件数量（默认为1）

#### 从邮件中提取验证码或链接
- **URL**: `/get_code_or_link`
- **方法**: `GET`
- **请求参数**:
  - `client_id`: 客户端ID
  - `refresh_token`: 刷新令牌
  - `subject_pattern`: 主题匹配模式（可选）
  - `sender`: 发件人邮箱（可选）
  - `folder_id`: 文件夹ID（默认为'inbox'）
  - `pattern`: 提取的正则表达式（可选）
  - `between`: 提取特定字符之间的内容（可选）

#### 删除收件箱中的所有邮件
- **URL**: `/delete_all_inbox_emails`
- **方法**: `DELETE`
- **请求参数**:
  - `client_id`: 客户端ID
  - `refresh_token`: 刷新令牌

#### 删除垃圾邮件中的所有邮件
- **URL**: `/delete_all_junkemail_emails`
- **方法**: `DELETE`
- **请求参数**:
  - `client_id`: 客户端ID
  - `refresh_token`: 刷新令牌

## 日志
所有操作的日志将记录在 `logfile.log` 文件中。

## 许可证
该项目采用 MIT 许可证，详情请查看 LICENSE 文件。
