# 定时邮件收发系统规格说明

## 1. 项目概述

本项目是一个基于 Python 的定时邮件收发系统，提供两个核心功能：
- **定时发送邮件**：支持自定义定时规则、邮件内容、附件、SMTP配置和收件人
- **定时接收邮件附件**：支持 IMAP 和 Outlook(VBS) 两种方式，可自定义筛选条件和附件保存规则

两个功能可独立启用或同时运行。

## 2. 功能需求

### 2.1 定时发送邮件功能

#### 2.1.1 定时配置
- 支持多种循环模式：
  - 小时循环（每隔N小时发送）
  - 天循环（每天指定时间发送）
  - 周循环（每周指定日期和时间发送）
  - 一次性定时（指定具体日期时间发送一次）
- 支持自定义开始时间和结束时间
- 支持启用/禁用定时任务

#### 2.1.2 邮件内容配置
- 邮件主题（支持变量替换）
- 邮件正文（支持纯文本和HTML格式）
- 发件人名称
- 回复地址

#### 2.1.3 附件配置
- 支持添加多个附件
- 支持指定附件路径（绝对路径或相对路径）
- 支持通配符匹配（如 `*.pdf`）
- 支持附件大小限制检查
- 支持附件不存在时的处理策略（跳过/报错/发送无附件邮件）

#### 2.1.4 SMTP配置
- SMTP服务器地址
- SMTP端口（支持SSL/TLS）
- 认证方式（用户名密码/无认证）
- SSL/TLS加密选项
- 连接超时设置
- 支持配置多个SMTP服务器（主备切换）

#### 2.1.5 收件人配置
- 支持多个收件人
- 支持抄送（CC）
- 支持密送（BCC）
- 支持从文件读取收件人列表
- 支持收件人分组

### 2.2 定时接收邮件附件功能

#### 2.2.1 定时配置
- 与发送功能相同的定时规则
- 独立的启用/禁用开关

#### 2.2.2 接收方式

**方式一：IMAP协议**
- IMAP服务器地址
- IMAP端口
- 用户名和密码
- SSL/TLS加密选项
- 邮箱文件夹选择（INBOX、已发送等）

**方式二：Outlook VBS调用**
- Outlook配置文件名称
- 邮箱账户选择
- 文件夹路径

#### 2.2.3 邮件筛选条件
- 发件人地址筛选（支持多个）
- 邮件主题关键词筛选（支持正则表达式）
- 收件时间范围筛选
- 未读/已读状态筛选
- 是否有附件筛选
- 自定义筛选规则

#### 2.2.4 附件保存配置
- 保存路径（支持变量替换，如日期、时间等）
- 文件重名处理策略：
  - 覆盖现有文件
  - 自动重命名（添加序号或时间戳）
  - 跳过已存在文件
- 附件类型过滤（只下载指定扩展名）
- 附件大小限制

#### 2.2.5 邮件处理
- 下载后标记为已读
- 移动到指定文件夹
- 删除原邮件（可选）
- 保留邮件副本

### 2.3 系统管理功能

#### 2.3.1 配置管理
- 配置文件格式：JSON 或 YAML
- 支持配置文件加密（敏感信息）
- 配置文件验证
- 配置热重载

#### 2.3.2 日志记录
- 日志级别：DEBUG、INFO、WARNING、ERROR
- 日志文件轮转（按大小或时间）
- 日志文件路径配置
- 控制台日志输出选项

#### 2.3.3 错误处理
- 发送失败重试机制
- 网络异常处理
- 认证失败处理
- 附件读取失败处理
- 错误通知（邮件或系统通知）

#### 2.3.4 运行模式
- 命令行运行
- 后台服务运行
- 系统托盘运行（可选）

## 3. 技术架构

### 3.1 项目结构
```
定时发邮件和收邮件/
├── config/                 # 配置文件目录
│   ├── config.yaml        # 主配置文件
│   └── config.example.yaml # 配置示例文件
├── logs/                   # 日志文件目录
├── attachments/            # 默认附件目录
│   ├── send/              # 发送附件暂存
│   └── receive/           # 接收附件保存
├── src/                    # 源代码目录
│   ├── __init__.py
│   ├── main.py            # 主程序入口
│   ├── sender.py          # 邮件发送模块
│   ├── receiver.py        # 邮件接收模块
│   ├── scheduler.py       # 定时任务调度模块
│   ├── config_manager.py  # 配置管理模块
│   ├── logger.py          # 日志模块
│   └── utils.py           # 工具函数模块
├── tests/                  # 测试目录
├── requirements.txt        # 依赖包列表
├── README.md              # 项目说明文档
└── run.py                 # 启动脚本
```

### 3.2 技术选型

#### 3.2.1 核心依赖
- **Python**: 3.8+
- **调度器**: schedule 或 APScheduler
- **邮件发送**: smtplib (标准库)
- **邮件接收**: 
  - imaplib (标准库) - IMAP方式
  - pywin32 - Outlook VBS调用
- **配置解析**: PyYAML
- **日志**: logging (标准库)
- **邮件解析**: email (标准库)

#### 3.2.2 可选依赖
- **加密**: cryptography (配置文件加密)
- **GUI**: tkinter 或 PyQt5 (系统托盘)
- **通知**: win10toast (Windows通知)

### 3.3 模块设计

#### 3.3.1 配置管理模块 (config_manager.py)
```python
class ConfigManager:
    - load_config()          # 加载配置文件
    - validate_config()      # 验证配置
    - get_sender_config()    # 获取发送配置
    - get_receiver_config()  # 获取接收配置
    - reload_config()        # 重新加载配置
```

#### 3.3.2 邮件发送模块 (sender.py)
```python
class EmailSender:
    - __init__(smtp_config)  # 初始化SMTP配置
    - connect()              # 连接SMTP服务器
    - send_email()           # 发送邮件
    - add_attachment()       # 添加附件
    - disconnect()           # 断开连接
    - test_connection()      # 测试连接
```

#### 3.3.3 邮件接收模块 (receiver.py)
```python
class EmailReceiver:
    - __init__(receiver_config)  # 初始化接收配置
    - connect()                  # 连接邮件服务器
    - fetch_emails()             # 获取邮件列表
    - filter_emails()            # 筛选邮件
    - save_attachments()         # 保存附件
    - disconnect()               # 断开连接
    
class IMAPReceiver(EmailReceiver):
    - imap协议实现
    
class OutlookReceiver(EmailReceiver):
    - outlook vbs调用实现
```

#### 3.3.4 定时调度模块 (scheduler.py)
```python
class TaskScheduler:
    - __init__()            # 初始化调度器
    - add_sender_task()     # 添加发送任务
    - add_receiver_task()   # 添加接收任务
    - start()               # 启动调度器
    - stop()                # 停止调度器
    - pause_task()          # 暂停任务
    - resume_task()         # 恢复任务
```

## 4. 配置文件格式

### 4.1 主配置文件 (config.yaml)
```yaml
# 系统配置
system:
  log_level: INFO
  log_file: logs/app.log
  log_max_size: 10MB
  log_backup_count: 5

# 发送邮件配置
sender:
  enabled: true
  schedule:
    type: daily          # hourly/daily/weekly/once
    time: "09:00"        # 每天9点
    # interval: 2        # 小时循环时的间隔
    # weekdays: [1,2,3,4,5]  # 周循环时的星期几
  
  smtp:
    host: smtp.example.com
    port: 587
    username: user@example.com
    password: encrypted_password
    use_tls: true
    timeout: 30
  
  email:
    from_name: "发件人名称"
    reply_to: reply@example.com
    subject: "邮件主题 - {date}"
    body_type: html      # text/html
    body: |
      <html>
      <body>
      <h1>邮件内容</h1>
      </body>
      </html>
  
  recipients:
    to:
      - recipient1@example.com
      - recipient2@example.com
    cc:
      - cc@example.com
    bcc:
      - bcc@example.com
  
  attachments:
    - path: attachments/send/report.pdf
      required: true
    - path: attachments/send/*.xlsx
      required: false

# 接收邮件配置
receiver:
  enabled: true
  schedule:
    type: hourly
    interval: 2          # 每2小时
  
  method: imap           # imap/outlook
  
  imap:
    host: imap.example.com
    port: 993
    username: user@example.com
    password: encrypted_password
    use_ssl: true
    folder: INBOX
  
  outlook:
    profile: Outlook
    mailbox: user@example.com
    folder: Inbox
  
  filters:
    from:
      - sender@example.com
    subject_pattern: ".*报告.*"
    has_attachment: true
    unread_only: true
    date_from: "2024-01-01"
    date_to: "2024-12-31"
  
  save:
    path: attachments/receive/{date}/{sender}
    filename_conflict: rename  # overwrite/rename/skip
    allowed_extensions:
      - pdf
      - xlsx
      - docx
    max_size: 50MB
  
  after_receive:
    mark_read: true
    move_to: null
    delete: false
```

## 5. 运行方式

### 5.1 命令行运行
```bash
# 前台运行
python run.py

# 指定配置文件
python run.py --config config/config.yaml

# 只运行发送功能
python run.py --sender-only

# 只运行接收功能
python run.py --receiver-only

# 测试配置
python run.py --test-config

# 后台运行（Linux/Mac）
nohup python run.py &

# 后台运行（Windows）
pythonw run.py
```

### 5.2 系统服务运行
```bash
# 安装为系统服务
python run.py --install-service

# 卸载系统服务
python run.py --uninstall-service

# 启动服务
python run.py --start-service

# 停止服务
python run.py --stop-service
```

## 6. 错误处理策略

### 6.1 发送失败处理
- 网络错误：自动重试3次，间隔递增（1分钟、5分钟、10分钟）
- 认证失败：记录错误日志，发送通知，停止任务
- 附件不存在：根据配置跳过或报错
- SMTP服务器不可用：尝试备用服务器

### 6.2 接收失败处理
- 连接失败：自动重试3次
- 认证失败：记录错误日志，发送通知，停止任务
- 附件保存失败：记录错误，继续处理下一封邮件
- 磁盘空间不足：发送警告通知，暂停接收

### 6.3 通知机制
- 发送邮件通知到管理员邮箱
- Windows系统通知（可选）
- 日志文件记录

## 7. 安全考虑

### 7.1 密码安全
- 配置文件中的密码加密存储
- 支持环境变量读取敏感信息
- 不在日志中记录密码信息

### 7.2 文件安全
- 附件路径验证，防止目录遍历攻击
- 文件大小限制
- 文件类型白名单

### 7.3 网络安全
- 支持SSL/TLS加密连接
- 证书验证
- 连接超时设置

## 8. 扩展性设计

### 8.1 插件系统
- 支持自定义邮件处理器插件
- 支持自定义附件处理器插件
- 支持自定义通知方式插件

### 8.2 API接口
- REST API接口（可选）
- 支持远程配置管理
- 支持任务状态查询

## 9. 测试要求

### 9.1 单元测试
- 配置管理模块测试
- 邮件发送模块测试
- 邮件接收模块测试
- 调度模块测试

### 9.2 集成测试
- 端到端邮件发送测试
- 端到端邮件接收测试
- 定时任务测试

### 9.3 性能测试
- 大量附件发送测试
- 大量邮件接收测试
- 长时间运行稳定性测试

## 10. 文档要求

### 10.1 用户文档
- 安装部署指南
- 配置说明文档
- 使用手册
- 常见问题解答

### 10.2 开发文档
- 架构设计文档
- API文档
- 代码注释规范
