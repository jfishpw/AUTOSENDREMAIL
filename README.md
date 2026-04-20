# 定时邮件收发系统

一个基于 Python 的定时邮件收发系统，支持定时发送邮件和定时接收邮件附件。

## 功能特性

### 发送邮件功能
- ✅ 多种定时规则（小时循环/天循环/周循环/一次性定时）
- ✅ 支持纯文本和HTML格式邮件
- ✅ 支持多个附件和通配符匹配
- ✅ 支持多个收件人、抄送、密送
- ✅ 支持SSL/TLS加密
- ✅ 发送失败自动重试

### 接收邮件功能
- ✅ 支持 IMAP 协议
- ✅ 支持 Outlook VBS 调用（Windows）
- ✅ 多条件筛选（发件人、主题、时间、状态）
- ✅ 文件重名处理（覆盖/重命名/跳过）
- ✅ 附件类型和大小过滤
- ✅ 邮件后处理（标记已读、移动、删除）

### 系统特性
- ✅ 两个功能可独立启用或同时运行
- ✅ 配置文件加密存储密码
- ✅ 完善的日志系统
- ✅ 多种运行模式
- ✅ **图形界面（GUI）和命令行双模式**

## 环境要求

- Python 3.8 或更高版本
- Windows / Linux / macOS

## 安装

### 1. 克隆或下载项目

```bash
cd 定时发邮件和收邮件
```

### 2. 创建虚拟环境（推荐）

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

## 配置

### 1. 复制配置文件模板

```bash
cp config/config.example.yaml config/config.yaml
```

### 2. 编辑配置文件

编辑 `config/config.yaml`，配置以下内容：

#### 系统配置
```yaml
system:
  log_level: INFO          # 日志级别
  log_file: logs/app.log   # 日志文件路径
  log_max_size: 10MB       # 日志文件最大大小
  log_backup_count: 5      # 日志文件备份数量
```

#### 发送邮件配置
```yaml
sender:
  enabled: true            # 是否启用发送功能
  schedule:
    type: daily            # 定时类型: hourly/daily/weekly/once
    time: '09:00'          # 执行时间
  smtp:
    host: smtp.example.com
    port: 587
    username: user@example.com
    password: your_password
    use_tls: true
  email:
    subject: '邮件主题 - {date}'
    body_type: html        # text 或 html
    body: '<html><body>邮件内容</body></html>'
  recipients:
    to:
      - recipient@example.com
  attachments:
    - path: attachments/send/report.pdf
```

#### 接收邮件配置
```yaml
receiver:
  enabled: true            # 是否启用接收功能
  schedule:
    type: hourly
    interval: 2            # 每2小时
  method: imap             # imap 或 outlook
  imap:
    host: imap.example.com
    port: 993
    username: user@example.com
    password: your_password
    use_ssl: true
  filters:
    has_attachment: true
    unread_only: true
    latest_only: false  # 只接收最新的一封邮件
  save:
    path: attachments/receive/{date}/{sender}
    filename_conflict: rename  # overwrite/rename/skip
```

### 3. 密码加密（可选）

为了安全，建议加密配置文件中的密码：

```python
from src.config_manager import ConfigManager

config_manager = ConfigManager()
encrypted = config_manager.encrypt('your_password')
print(f'encrypted:{encrypted}')
```

将输出的加密字符串替换到配置文件中的 `password` 字段。

## 使用方法

### 图形界面模式（推荐）

```bash
python run.py --gui
```

GUI 界面功能：
- 🟢 **启动/停止服务** - 一键控制
- 📊 **实时状态** - 查看运行状态、下次执行时间
- 📝 **运行日志** - 实时查看日志输出
- ⚙️ **设置菜单** - 可视化编辑所有配置
  - 系统设置（日志级别、日志文件等）
  - 发送邮件设置（SMTP、邮件内容、收件人、附件）
  - 接收邮件设置（IMAP/Outlook、筛选条件、保存设置）
  - YAML 编辑器（直接编辑配置文件）
- 🧪 **测试配置** - 一键测试连接

### 命令行模式

#### 前台运行（所有功能）
```bash
python run.py
```

#### 只运行发送功能
```bash
python run.py --sender-only
```

#### 只运行接收功能
```bash
python run.py --receiver-only
```

#### 指定配置文件
```bash
python run.py --config path/to/config.yaml
```

### 测试配置

```bash
python run.py --test-config
```

### 运行一次任务

```bash
# 运行所有任务一次
python run.py --run-once

# 只运行发送任务一次
python run.py --run-once --task-type sender

# 只运行接收任务一次
python run.py --run-once --task-type receiver
```

### 后台运行

#### Linux/Mac
```bash
nohup python run.py > /dev/null 2>&1 &
```

#### Windows
```bash
pythonw run.py
```

## 定时规则说明

### 小时循环 (hourly)
每隔N小时执行一次：
```yaml
schedule:
  type: hourly
  interval: 2  # 每2小时
```

### 天循环 (daily)
每天指定时间执行：
```yaml
schedule:
  type: daily
  time: '09:00'  # 每天9点
```

### 周循环 (weekly)
每周指定日期和时间执行：
```yaml
schedule:
  type: weekly
  time: '09:00'
  weekdays: [1, 2, 3, 4, 5]  # 周一到周五（1=周一，7=周日）
```

### 一次性定时 (once)
指定日期时间执行一次：
```yaml
schedule:
  type: once
  datetime: '2024-12-31 23:59:00'
```

或者：
```yaml
schedule:
  type: once
  date: '2024-12-31'
  time: '23:59'
```

## 变量替换

在邮件主题、正文和保存路径中可以使用以下变量：

- `{date}` - 当前日期 (YYYY-MM-DD)
- `{time}` - 当前时间 (HH-MM-SS)
- `{datetime}` - 当前日期时间 (YYYY-MM-DD_HH-MM-SS)
- `{year}` - 年份
- `{month}` - 月份
- `{day}` - 日期
- `{sender}` - 发件人邮箱（仅接收功能）

示例：
```yaml
email:
  subject: '每日报告 - {date}'

save:
  path: 'attachments/receive/{date}/{sender}'
```

### subject_pattern 主题筛选（正则表达式）

`subject_pattern` 使用**正则表达式**来匹配邮件主题，非常灵活：

**常用正则表达式符号：**

| 符号 | 含义 | 示例 |
|------|------|------|
| `.` | 任意单个字符 | `a.c` 匹配 `abc` |
| `.*` | 任意字符重复任意次 | `KW.*` 匹配 `KW16分发版` |
| `\d` | 任意数字 | `KW\d+` 匹配 `KW16` |
| `^` | 字符串开头 | `^日报` 匹配 `日报xxx` |
| `$` | 字符串结尾 | `周报$` 匹配 `xxx周报` |
| `[]` | 字符集 | `[KW]` 匹配 `K` 或 `W` |
| `\|` | 或 | `KW16\|KW17` 匹配 `KW16` 或 `KW17` |

**实用示例：**

```yaml
# 匹配包含 "日报" 的邮件
subject_pattern: '.*日报.*'

# 匹配 KW16 或 KW17 开头的邮件
subject_pattern: 'KW1[67].*'

# 匹配分发版或周报结尾的邮件
subject_pattern: '.*(分发版|周报)$'

# 匹配多种关键词（用 | 分隔）
subject_pattern: 'KW16.*报告|KW17.*分发版|.*周报.*'

# 匹配回复类邮件（以 Re: 开头）
subject_pattern: 'Re:.*KW16.*'
```

**注意：** 正则表达式会忽略大小写（`re.IGNORECASE`），所以 `kw16` 和 `KW16` 都可以匹配。

## 接收邮件筛选条件

接收邮件功能支持多种筛选条件，所有条件为"与"的关系（需要同时满足）：

### 筛选条件说明

```yaml
filters:
  from:                     # 发件人筛选（支持多个）
    - sender1@example.com
    - sender2@example.com
  subject_pattern: '.*报告.*'  # 主题关键词（支持正则表达式）
  has_attachment: true      # 是否只接收有附件的邮件
  unread_only: true         # 是否只接收未读邮件
  latest_only: false        # 是否只接收最新的一封邮件
  max_emails: 100           # 最多获取的邮件数量（避免邮箱邮件过多导致超时）
```

### max_emails 功能说明

`max_emails` 用于限制获取邮件的数量，避免邮箱中邮件过多导致超时或性能问题：

- **默认值**：100
- **作用**：只获取最近的 N 封邮件（按邮件ID排序，最新的在最后）
- **建议**：根据邮箱邮件数量和网络情况调整，一般设置为 50-200 之间

**示例场景：**

```yaml
# 邮箱中有几千封邮件，只检查最近的50封
filters:
  max_emails: 50
  has_attachment: true
```

### latest_only 功能说明

`latest_only` 是一个实用的开关，当设置为 `true` 时：

- **功能**：在所有筛选条件过滤后，只保留时间最新的一封邮件
- **用途**：适用于只需要获取最新报告或最新通知的场景
- **排序依据**：根据邮件的发送时间进行排序

**示例场景：**

```yaml
# 每天只接收最新的日报
filters:
  from:
    - report@company.com
  subject_pattern: '.*日报.*'
  has_attachment: true
  latest_only: true  # 即使有多封日报，也只保存最新的一封
```

**处理流程：**
1. 根据其他筛选条件过滤邮件
2. 如果有多封邮件符合条件，按时间排序
3. 只保留时间最新的一封邮件
4. 保存该邮件的附件

**日志输出示例：**
```
2026-04-20 18:00:01 - INFO - 筛选邮件: 15 -> 5
2026-04-20 18:00:02 - INFO - 只保留最新邮件: 1 封
2026-04-20 18:00:03 - INFO - 保存附件: attachments/receive/2026-04-20/report/report.pdf
```

## 日志

日志文件默认保存在 `logs/app.log`，支持自动轮转。

查看日志：
```bash
tail -f logs/app.log
```

## 目录结构

```
定时发邮件和收邮件/
├── config/                 # 配置文件目录
│   ├── config.yaml        # 主配置文件
│   └── config.example.yaml # 配置示例
├── logs/                   # 日志文件目录
├── attachments/            # 附件目录
│   ├── send/              # 发送附件
│   └── receive/           # 接收附件
├── src/                    # 源代码目录
│   ├── __init__.py
│   ├── main.py            # 主程序
│   ├── sender.py          # 邮件发送模块
│   ├── receiver.py        # 邮件接收模块
│   ├── scheduler.py       # 定时调度模块
│   ├── config_manager.py  # 配置管理模块
│   ├── logger.py          # 日志模块
│   └── utils.py           # 工具函数模块
├── requirements.txt        # 依赖包列表
├── run.py                 # 启动脚本
└── README.md              # 项目说明
```

## 常见问题

### 1. SMTP连接失败
- 检查SMTP服务器地址和端口
- 确认是否需要SSL/TLS加密
- 检查用户名和密码是否正确
- 某些邮箱需要使用应用专用密码

### 2. IMAP连接失败
- 确认IMAP服务已开启
- 检查服务器地址和端口
- 某些邮箱需要使用应用专用密码

### 3. Outlook连接失败
- 确保已安装 Microsoft Outlook
- 检查Outlook配置文件名称
- 确认邮箱账户已配置

### 4. 附件保存失败
- 检查保存路径是否有写入权限
- 确认磁盘空间充足
- 检查文件名是否包含非法字符

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
