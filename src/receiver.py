"""
邮件接收模块
"""
import os
import re
import ssl
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from .logger import get_logger
from .utils import (
    ensure_dir, expand_path, handle_file_conflict, 
    sanitize_filename, parse_size_to_bytes
)


class EmailReceiver:
    """邮件接收器基类"""
    
    def __init__(self, config: Dict[str, Any]):
        """
        初始化邮件接收器
        
        Args:
            config: 接收配置字典
        """
        self.logger = get_logger()
        self.config = config
        self.filters = config.get('filters', {})
        self.save_config = config.get('save', {})
        self.after_receive = config.get('after_receive', {})
    
    def connect(self) -> bool:
        """连接邮件服务器"""
        raise NotImplementedError
    
    def disconnect(self):
        """断开连接"""
        raise NotImplementedError
    
    def fetch_emails(self) -> List[Dict[str, Any]]:
        """获取邮件列表"""
        raise NotImplementedError
    
    def filter_emails(self, emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        过滤邮件
        
        Args:
            emails: 邮件列表
        
        Returns:
            筛选后的邮件列表
        """
        filtered = []
        
        for email_data in emails:
            subject = email_data.get('subject', '')
            sender = email_data.get('from', '')
            self.logger.info(f"邮件: 主题='{subject[:50]}', 发件人='{sender[:30]}'")
            
            if self._match_filters(email_data):
                filtered.append(email_data)
        
        self.logger.info(f"筛选邮件: {len(emails)} -> {len(filtered)}")
        
        if self.filters.get('latest_only', False) and filtered:
            filtered = self._get_latest_email(filtered)
            self.logger.info(f"只保留最新邮件: {len(filtered)} 封")
        
        return filtered
    
    def _get_latest_email(self, emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        获取最新的邮件
        
        Args:
            emails: 邮件列表
        
        Returns:
            包含最新邮件的列表
        """
        def parse_email_date(email_data: Dict[str, Any]) -> datetime:
            """解析邮件日期"""
            try:
                date_str = email_data.get('date', '')
                if date_str:
                    dt = parsedate_to_datetime(date_str)
                    if dt.tzinfo is not None:
                        dt = dt.replace(tzinfo=None)
                    return dt
            except Exception as e:
                self.logger.debug(f"解析邮件日期失败: {str(e)}")
            return datetime.min
        
        sorted_emails = sorted(emails, key=parse_email_date, reverse=True)
        return [sorted_emails[0]] if sorted_emails else []
    
    def _match_filters(self, email_data: Dict[str, Any]) -> bool:
        """
        检查邮件是否匹配筛选条件
        
        Args:
            email_data: 邮件数据
        
        Returns:
            是否匹配
        """
        subject = email_data.get('subject', '')
        sender = email_data.get('from', '')
        
        self.logger.debug(f"检查邮件: 主题='{subject}', 发件人='{sender}'")
        
        if self.filters.get('from'):
            if not any(f.lower() in sender.lower() for f in self.filters['from']):
                self.logger.debug(f"  -> 发件人不匹配: {sender}")
                return False
        
        if self.filters.get('subject_pattern'):
            pattern = self.filters['subject_pattern']
            match = re.search(pattern, subject, re.IGNORECASE)
            self.logger.debug(f"  -> 主题='{subject}', 模式='{pattern}', 匹配={match is not None}")
            if not match:
                return False
        
        if self.filters.get('has_attachment'):
            if not email_data.get('has_attachment', False):
                self.logger.debug(f"  -> 没有附件")
                return False
        
        if self.filters.get('unread_only'):
            if email_data.get('read', False):
                self.logger.debug(f"  -> 邮件已读")
                return False
        
        self.logger.debug(f"  -> 匹配成功")
        return True
    
    def save_attachments(self, email_data: Dict[str, Any]) -> List[str]:
        """
        保存邮件附件
        
        Args:
            email_data: 邮件数据
        
        Returns:
            保存的附件路径列表
        """
        saved_files = []
        attachments = email_data.get('attachments', [])
        
        if not attachments:
            return saved_files
        
        base_path = self.save_config.get('path', 'attachments/receive')
        sender = email_data.get('from', 'unknown')
        sender = sender.split('<')[-1].split('>')[0] if '<' in sender else sender
        sender = sanitize_filename(sender.split('@')[0])
        
        save_path = expand_path(base_path, {'sender': sender})
        ensure_dir(save_path)
        
        allowed_extensions = self.save_config.get('allowed_extensions', [])
        max_size = parse_size_to_bytes(self.save_config.get('max_size', '50MB'))
        conflict_strategy = self.save_config.get('filename_conflict', 'rename')
        
        for attachment in attachments:
            filename = attachment['filename']
            content = attachment['content']
            
            ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
            if allowed_extensions and ext not in allowed_extensions:
                self.logger.debug(f"跳过不允许的文件类型: {filename}")
                continue
            
            if len(content) > max_size:
                self.logger.warning(f"附件过大，跳过: {filename}")
                continue
            
            filename = sanitize_filename(filename)
            filepath = os.path.join(save_path, filename)
            
            filepath, should_save = handle_file_conflict(filepath, conflict_strategy)
            
            if should_save:
                try:
                    with open(filepath, 'wb') as f:
                        f.write(content)
                    saved_files.append(filepath)
                    self.logger.info(f"保存附件: {filepath}")
                except Exception as e:
                    self.logger.error(f"保存附件失败 {filename}: {str(e)}")
        
        return saved_files
    
    def process_emails(self) -> Tuple[int, int]:
        """
        处理邮件（获取、筛选、保存附件）
        
        Returns:
            (处理的邮件数, 保存的附件数)
        """
        raise NotImplementedError


class IMAPReceiver(EmailReceiver):
    """IMAP邮件接收器"""
    
    def __init__(self, config: Dict[str, Any]):
        """
        初始化IMAP接收器
        
        Args:
            config: 接收配置字典
        """
        super().__init__(config)
        self.imap_config = config.get('imap', {})
        self.connection = None
    
    def connect(self) -> bool:
        """
        连接IMAP服务器
        
        Returns:
            是否连接成功
        """
        try:
            host = self.imap_config['host']
            port = self.imap_config['port']
            use_ssl = self.imap_config.get('use_ssl', True)
            
            if use_ssl:
                ssl_context = ssl.create_default_context()
                ssl_context.check_hostname = False
                ssl_context.verify_mode = ssl.CERT_NONE
                ssl_context.set_ciphers('DEFAULT@SECLEVEL=1')
                self.connection = imaplib.IMAP4_SSL(host, port, ssl_context=ssl_context)
            else:
                self.connection = imaplib.IMAP4(host, port)
            
            username = self.imap_config['username']
            password = self.imap_config['password']
            self.connection.login(username, password)
            
            self.logger.info(f"成功连接到IMAP服务器: {host}:{port}")
            return True
            
        except Exception as e:
            self.logger.error(f"连接IMAP服务器失败: {str(e)}")
            return False
    
    def disconnect(self):
        """断开IMAP连接"""
        if self.connection:
            try:
                try:
                    self.connection.close()
                except:
                    pass
                self.connection.logout()
                self.logger.info("已断开IMAP连接")
            except Exception as e:
                self.logger.warning(f"断开IMAP连接时出错: {str(e)}")
            finally:
                self.connection = None
    
    def fetch_emails(self) -> List[Dict[str, Any]]:
        """
        获取邮件列表
        
        Returns:
            邮件数据列表
        """
        if not self.connection:
            if not self.connect():
                return []
        
        try:
            folder = self.imap_config.get('folder', 'INBOX')
            self.connection.select(folder)
            
            status, messages = self.connection.search(None, 'ALL')
            if status != 'OK':
                self.logger.error("搜索邮件失败")
                return []
            
            email_ids = messages[0].split()
            total_emails = len(email_ids)
            self.logger.info(f"邮箱中共有 {total_emails} 封邮件")
            
            max_emails = self.filters.get('max_emails', 100)
            if len(email_ids) > max_emails:
                email_ids = email_ids[-max_emails:]
                self.logger.info(f"只获取最近的 {max_emails} 封邮件")
            
            emails = []
            failed_count = 0
            
            for i, email_id in enumerate(email_ids, 1):
                try:
                    email_data = self._fetch_email(email_id)
                    if email_data:
                        emails.append(email_data)
                    
                    if i % 10 == 0:
                        self.logger.debug(f"已处理 {i}/{len(email_ids)} 封邮件")
                except Exception as e:
                    failed_count += 1
                    self.logger.warning(f"获取邮件 {email_id} 失败: {str(e)}")
                    if failed_count > 10:
                        self.logger.error("失败次数过多，停止获取")
                        break
            
            self.logger.info(f"成功获取 {len(emails)} 封邮件")
            return emails
            
        except Exception as e:
            self.logger.error(f"获取邮件失败: {str(e)}")
            return []
    
    def _fetch_email(self, email_id: bytes) -> Optional[Dict[str, Any]]:
        """
        获取单封邮件
        
        Args:
            email_id: 邮件ID
        
        Returns:
            邮件数据字典
        """
        try:
            status, msg_data = self.connection.fetch(email_id, '(RFC822)')
            if status != 'OK':
                return None
            
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            subject = self._decode_header(msg.get('Subject', ''))
            from_addr = self._decode_header(msg.get('From', ''))
            to_addr = self._decode_header(msg.get('To', ''))
            date_str = msg.get('Date', '')
            
            attachments = []
            has_attachment = False
            
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                
                filename = part.get_filename()
                if filename:
                    filename = self._decode_header(filename)
                    content = part.get_payload(decode=True)
                    attachments.append({
                        'filename': filename,
                        'content': content
                    })
                    has_attachment = True
            
            return {
                'id': email_id,
                'subject': subject,
                'from': from_addr,
                'to': to_addr,
                'date': date_str,
                'has_attachment': has_attachment,
                'attachments': attachments,
                'read': False,
                'message': msg
            }
            
        except Exception as e:
            self.logger.error(f"获取邮件失败 {email_id}: {str(e)}")
            return None
    
    def _decode_header(self, header: str) -> str:
        """
        解码邮件头
        
        Args:
            header: 邮件头字符串
        
        Returns:
            解码后的字符串
        """
        if not header:
            return ''
        
        decoded_parts = []
        for part, encoding in decode_header(header):
            if isinstance(part, bytes):
                try:
                    decoded_parts.append(part.decode(encoding or 'utf-8', errors='ignore'))
                except:
                    decoded_parts.append(part.decode('utf-8', errors='ignore'))
            else:
                decoded_parts.append(part)
        
        return ''.join(decoded_parts)
    
    def mark_as_read(self, email_id: bytes) -> bool:
        """
        标记邮件为已读
        
        Args:
            email_id: 邮件ID
        
        Returns:
            是否成功
        """
        try:
            self.connection.store(email_id, '+FLAGS', '\\Seen')
            self.logger.debug(f"标记邮件为已读: {email_id}")
            return True
        except Exception as e:
            self.logger.error(f"标记邮件为已读失败: {str(e)}")
            return False
    
    def move_email(self, email_id: bytes, dest_folder: str) -> bool:
        """
        移动邮件到指定文件夹
        
        Args:
            email_id: 邮件ID
            dest_folder: 目标文件夹
        
        Returns:
            是否成功
        """
        try:
            self.connection.copy(email_id, dest_folder)
            self.connection.store(email_id, '+FLAGS', '\\Deleted')
            self.logger.info(f"移动邮件到 {dest_folder}: {email_id}")
            return True
        except Exception as e:
            self.logger.error(f"移动邮件失败: {str(e)}")
            return False
    
    def delete_email(self, email_id: bytes) -> bool:
        """
        删除邮件
        
        Args:
            email_id: 邮件ID
        
        Returns:
            是否成功
        """
        try:
            self.connection.store(email_id, '+FLAGS', '\\Deleted')
            self.logger.info(f"标记删除邮件: {email_id}")
            return True
        except Exception as e:
            self.logger.error(f"删除邮件失败: {str(e)}")
            return False
    
    def process_emails(self) -> Tuple[int, int]:
        """
        处理邮件
        
        Returns:
            (处理的邮件数, 保存的附件数)
        """
        emails = self.fetch_emails()
        filtered_emails = self.filter_emails(emails)
        
        processed_count = 0
        attachment_count = 0
        
        for email_data in filtered_emails:
            saved_files = self.save_attachments(email_data)
            
            if saved_files:
                processed_count += 1
                attachment_count += len(saved_files)
                
                if self.after_receive.get('mark_read'):
                    self.mark_as_read(email_data['id'])
                
                move_to = self.after_receive.get('move_to')
                if move_to:
                    self.move_email(email_data['id'], move_to)
                
                if self.after_receive.get('delete'):
                    self.delete_email(email_data['id'])
        
        return processed_count, attachment_count


class OutlookReceiver(EmailReceiver):
    """Outlook邮件接收器（通过VBS调用）"""
    
    def __init__(self, config: Dict[str, Any]):
        """
        初始化Outlook接收器
        
        Args:
            config: 接收配置字典
        """
        super().__init__(config)
        self.outlook_config = config.get('outlook', {})
        self.outlook = None
        self.namespace = None
    
    def connect(self) -> bool:
        """
        连接Outlook
        
        Returns:
            是否连接成功
        """
        try:
            import win32com.client
            
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            self.logger.info("成功连接到Outlook")
            return True
            
        except Exception as e:
            self.logger.error(f"连接Outlook失败: {str(e)}")
            return False
    
    def disconnect(self):
        """断开Outlook连接"""
        self.outlook = None
        self.namespace = None
        self.logger.info("已断开Outlook连接")
    
    def fetch_emails(self) -> List[Dict[str, Any]]:
        """
        获取邮件列表
        
        Returns:
            邮件数据列表
        """
        if not self.namespace:
            if not self.connect():
                return []
        
        try:
            folder_name = self.outlook_config.get('folder', 'Inbox')
            folder = self._get_folder(folder_name)
            
            if not folder:
                self.logger.error(f"找不到文件夹: {folder_name}")
                return []
            
            messages = folder.Items
            messages.Sort("[ReceivedTime]", True)
            
            emails = []
            for message in messages:
                email_data = self._parse_message(message)
                if email_data:
                    emails.append(email_data)
            
            self.logger.info(f"获取到 {len(emails)} 封邮件")
            return emails
            
        except Exception as e:
            self.logger.error(f"获取邮件失败: {str(e)}")
            return []
    
    def _get_folder(self, folder_name: str):
        """
        获取文件夹对象
        
        Args:
            folder_name: 文件夹名称
        
        Returns:
            文件夹对象
        """
        try:
            for folder in self.namespace.Folders:
                for subfolder in folder.Folders:
                    if subfolder.Name == folder_name:
                        return subfolder
            return None
        except Exception as e:
            self.logger.error(f"获取文件夹失败: {str(e)}")
            return None
    
    def _parse_message(self, message) -> Optional[Dict[str, Any]]:
        """
        解析邮件消息
        
        Args:
            message: Outlook邮件对象
        
        Returns:
            邮件数据字典
        """
        try:
            attachments = []
            has_attachment = False
            
            if message.Attachments.Count > 0:
                has_attachment = True
                for i in range(1, message.Attachments.Count + 1):
                    attachment = message.Attachments.Item(i)
                    attachments.append({
                        'filename': attachment.FileName,
                        'attachment': attachment
                    })
            
            return {
                'id': message.EntryID,
                'subject': message.Subject or '',
                'from': message.SenderEmailAddress or '',
                'to': message.To or '',
                'date': str(message.ReceivedTime),
                'has_attachment': has_attachment,
                'attachments': attachments,
                'read': not message.UnRead,
                'message': message
            }
            
        except Exception as e:
            self.logger.error(f"解析邮件失败: {str(e)}")
            return None
    
    def save_attachments(self, email_data: Dict[str, Any]) -> List[str]:
        """
        保存邮件附件
        
        Args:
            email_data: 邮件数据
        
        Returns:
            保存的附件路径列表
        """
        saved_files = []
        attachments = email_data.get('attachments', [])
        
        if not attachments:
            return saved_files
        
        base_path = self.save_config.get('path', 'attachments/receive')
        sender = email_data.get('from', 'unknown')
        sender = sanitize_filename(sender.split('@')[0] if '@' in sender else sender)
        
        save_path = expand_path(base_path, {'sender': sender})
        ensure_dir(save_path)
        
        allowed_extensions = self.save_config.get('allowed_extensions', [])
        conflict_strategy = self.save_config.get('filename_conflict', 'rename')
        
        for attachment_data in attachments:
            filename = attachment_data['filename']
            attachment = attachment_data['attachment']
            
            ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''
            if allowed_extensions and ext not in allowed_extensions:
                self.logger.debug(f"跳过不允许的文件类型: {filename}")
                continue
            
            filename = sanitize_filename(filename)
            filepath = os.path.join(save_path, filename)
            
            filepath, should_save = handle_file_conflict(filepath, conflict_strategy)
            
            if should_save:
                try:
                    attachment.SaveAsFile(os.path.abspath(filepath))
                    saved_files.append(filepath)
                    self.logger.info(f"保存附件: {filepath}")
                except Exception as e:
                    self.logger.error(f"保存附件失败 {filename}: {str(e)}")
        
        return saved_files
    
    def mark_as_read(self, message) -> bool:
        """
        标记邮件为已读
        
        Args:
            message: Outlook邮件对象
        
        Returns:
            是否成功
        """
        try:
            message.UnRead = False
            message.Save()
            self.logger.debug(f"标记邮件为已读: {message.Subject}")
            return True
        except Exception as e:
            self.logger.error(f"标记邮件为已读失败: {str(e)}")
            return False
    
    def move_email(self, message, dest_folder: str) -> bool:
        """
        移动邮件到指定文件夹
        
        Args:
            message: Outlook邮件对象
            dest_folder: 目标文件夹名称
        
        Returns:
            是否成功
        """
        try:
            folder = self._get_folder(dest_folder)
            if folder:
                message.Move(folder)
                self.logger.info(f"移动邮件到 {dest_folder}: {message.Subject}")
                return True
            return False
        except Exception as e:
            self.logger.error(f"移动邮件失败: {str(e)}")
            return False
    
    def delete_email(self, message) -> bool:
        """
        删除邮件
        
        Args:
            message: Outlook邮件对象
        
        Returns:
            是否成功
        """
        try:
            message.Delete()
            self.logger.info(f"删除邮件: {message.Subject}")
            return True
        except Exception as e:
            self.logger.error(f"删除邮件失败: {str(e)}")
            return False
    
    def process_emails(self) -> Tuple[int, int]:
        """
        处理邮件
        
        Returns:
            (处理的邮件数, 保存的附件数)
        """
        emails = self.fetch_emails()
        filtered_emails = self.filter_emails(emails)
        
        processed_count = 0
        attachment_count = 0
        
        for email_data in filtered_emails:
            saved_files = self.save_attachments(email_data)
            
            if saved_files:
                processed_count += 1
                attachment_count += len(saved_files)
                
                message = email_data['message']
                
                if self.after_receive.get('mark_read'):
                    self.mark_as_read(message)
                
                move_to = self.after_receive.get('move_to')
                if move_to:
                    self.move_email(message, move_to)
                
                if self.after_receive.get('delete'):
                    self.delete_email(message)
        
        return processed_count, attachment_count


def create_receiver(config: Dict[str, Any]) -> EmailReceiver:
    """
    创建邮件接收器
    
    Args:
        config: 接收配置字典
    
    Returns:
        邮件接收器实例
    """
    method = config.get('method', 'imap')
    
    if method == 'imap':
        return IMAPReceiver(config)
    elif method == 'outlook':
        return OutlookReceiver(config)
    else:
        raise ValueError(f"不支持的接收方式: {method}")
