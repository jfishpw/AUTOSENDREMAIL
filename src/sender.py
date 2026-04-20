"""
邮件发送模块
"""
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from typing import List, Dict, Any, Optional
from .logger import get_logger


class EmailSender:
    """邮件发送器"""
    
    def __init__(self, smtp_config: Dict[str, Any]):
        """
        初始化邮件发送器
        
        Args:
            smtp_config: SMTP配置字典
        """
        self.logger = get_logger()
        self.smtp_config = smtp_config
        self.connection = None
        
    def connect(self) -> bool:
        """
        连接SMTP服务器
        
        Returns:
            是否连接成功
        """
        try:
            host = self.smtp_config['host']
            port = self.smtp_config['port']
            use_tls = self.smtp_config.get('use_tls', True)
            timeout = self.smtp_config.get('timeout', 30)
            
            if use_tls:
                self.connection = smtplib.SMTP_SSL(host, port, timeout=timeout)
            else:
                self.connection = smtplib.SMTP(host, port, timeout=timeout)
                if self.smtp_config.get('starttls', False):
                    self.connection.starttls()
            
            username = self.smtp_config.get('username')
            password = self.smtp_config.get('password')
            
            if username and password:
                self.connection.login(username, password)
            
            self.logger.info(f"成功连接到SMTP服务器: {host}:{port}")
            return True
            
        except Exception as e:
            self.logger.error(f"连接SMTP服务器失败: {str(e)}")
            return False
    
    def disconnect(self):
        """断开SMTP连接"""
        if self.connection:
            try:
                self.connection.quit()
                self.logger.info("已断开SMTP连接")
            except Exception as e:
                self.logger.warning(f"断开SMTP连接时出错: {str(e)}")
            finally:
                self.connection = None
    
    def test_connection(self) -> bool:
        """
        测试SMTP连接
        
        Returns:
            是否连接成功
        """
        try:
            success = self.connect()
            if success:
                self.disconnect()
            return success
        except Exception as e:
            self.logger.error(f"测试SMTP连接失败: {str(e)}")
            return False
    
    def send_email(
        self,
        to_addrs: List[str],
        subject: str,
        body: str,
        body_type: str = 'html',
        cc: List[str] = None,
        bcc: List[str] = None,
        attachments: List[str] = None,
        from_name: str = None,
        reply_to: str = None
    ) -> bool:
        """
        发送邮件
        
        Args:
            to_addrs: 收件人列表
            subject: 邮件主题
            body: 邮件正文
            body_type: 正文类型 ('text' 或 'html')
            cc: 抄送列表
            bcc: 密送列表
            attachments: 附件列表
            from_name: 发件人名称
            reply_to: 回复地址
        
        Returns:
            是否发送成功
        """
        try:
            msg = MIMEMultipart()
            
            username = self.smtp_config.get('username', '')
            from_name = from_name or username
            msg['From'] = f"{from_name} <{username}>" if from_name else username
            msg['To'] = ', '.join(to_addrs)
            msg['Subject'] = subject
            
            if reply_to:
                msg['Reply-To'] = reply_to
            
            if cc:
                msg['Cc'] = ', '.join(cc)
            
            body_content = MIMEText(body, body_type, 'utf-8')
            msg.attach(body_content)
            
            if attachments:
                for attachment_path in attachments:
                    if os.path.exists(attachment_path):
                        self._add_attachment(msg, attachment_path)
                    else:
                        self.logger.warning(f"附件不存在: {attachment_path}")
            
            all_recipients = to_addrs.copy()
            if cc:
                all_recipients.extend(cc)
            if bcc:
                all_recipients.extend(bcc)
            
            if not self.connection:
                if not self.connect():
                    return False
            
            self.connection.sendmail(username, all_recipients, msg.as_string())
            self.logger.info(f"邮件发送成功: {subject} -> {', '.join(to_addrs)}")
            return True
            
        except Exception as e:
            self.logger.error(f"邮件发送失败: {str(e)}")
            return False
    
    def _add_attachment(self, msg: MIMEMultipart, filepath: str):
        """
        添加附件到邮件
        
        Args:
            msg: 邮件消息对象
            filepath: 附件文件路径
        """
        try:
            filename = os.path.basename(filepath)
            
            with open(filepath, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {filename}'
            )
            msg.attach(part)
            
            self.logger.debug(f"添加附件: {filename}")
            
        except Exception as e:
            self.logger.error(f"添加附件失败 {filepath}: {str(e)}")
    
    def send_with_config(
        self,
        email_config: Dict[str, Any],
        recipients_config: Dict[str, Any],
        attachments_config: List[Dict[str, Any]] = None
    ) -> bool:
        """
        使用配置发送邮件
        
        Args:
            email_config: 邮件内容配置
            recipients_config: 收件人配置
            attachments_config: 附件配置
        
        Returns:
            是否发送成功
        """
        subject = email_config.get('subject', '')
        subject = self._replace_variables(subject)
        
        body = email_config.get('body', '')
        body = self._replace_variables(body)
        
        body_type = email_config.get('body_type', 'html')
        from_name = email_config.get('from_name')
        reply_to = email_config.get('reply_to')
        
        to_addrs = recipients_config.get('to', [])
        cc = recipients_config.get('cc', [])
        bcc = recipients_config.get('bcc', [])
        
        attachments = []
        if attachments_config:
            for att_config in attachments_config:
                att_path = att_config.get('path')
                if att_path:
                    from .utils import find_files
                    files = find_files(att_path)
                    attachments.extend(files)
        
        return self.send_email(
            to_addrs=to_addrs,
            subject=subject,
            body=body,
            body_type=body_type,
            cc=cc,
            bcc=bcc,
            attachments=attachments,
            from_name=from_name,
            reply_to=reply_to
        )
    
    def _replace_variables(self, text: str) -> str:
        """
        替换文本中的变量
        
        Args:
            text: 包含变量的文本
        
        Returns:
            替换后的文本
        """
        now = datetime.now()
        
        variables = {
            'date': now.strftime('%Y-%m-%d'),
            'time': now.strftime('%H:%M:%S'),
            'datetime': now.strftime('%Y-%m-%d %H:%M:%S'),
            'year': str(now.year),
            'month': f'{now.month:02d}',
            'day': f'{now.day:02d}',
        }
        
        result = text
        for key, value in variables.items():
            result = result.replace(f'{{{key}}}', value)
        
        return result
