"""
配置管理模块
"""
import os
import json
import yaml
from pathlib import Path
from typing import Dict, Any, Optional
from cryptography.fernet import Fernet
import base64


class ConfigManager:
    """配置管理器"""
    
    def __init__(self, config_path: str = None):
        self.config_path = config_path or os.path.join('config', 'config.yaml')
        self.config = {}
        self.key = self._get_or_create_key()
        self.cipher = Fernet(self.key)
        
    def _get_or_create_key(self) -> bytes:
        """获取或创建加密密钥"""
        key_file = os.path.join('config', '.key')
        if os.path.exists(key_file):
            with open(key_file, 'rb') as f:
                return f.read()
        else:
            key = Fernet.generate_key()
            os.makedirs(os.path.dirname(key_file), exist_ok=True)
            with open(key_file, 'wb') as f:
                f.write(key)
            return key
    
    def load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")
        
        with open(self.config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
        
        self._decrypt_passwords()
        self.validate_config()
        return self.config
    
    def _decrypt_passwords(self):
        """解密配置中的密码"""
        if 'sender' in self.config and 'smtp' in self.config['sender']:
            password = self.config['sender']['smtp'].get('password')
            if password and password.startswith('encrypted:'):
                encrypted = password[10:]
                self.config['sender']['smtp']['password'] = self.decrypt(encrypted)
        
        if 'receiver' in self.config:
            if 'imap' in self.config['receiver']:
                password = self.config['receiver']['imap'].get('password')
                if password and password.startswith('encrypted:'):
                    encrypted = password[10:]
                    self.config['receiver']['imap']['password'] = self.decrypt(encrypted)
    
    def encrypt(self, plain_text: str) -> str:
        """加密文本"""
        return self.cipher.encrypt(plain_text.encode()).decode()
    
    def decrypt(self, encrypted_text: str) -> str:
        """解密文本"""
        try:
            return self.cipher.decrypt(encrypted_text.encode()).decode()
        except Exception:
            return encrypted_text
    
    def validate_config(self) -> bool:
        """验证配置文件"""
        if not self.config:
            raise ValueError("配置为空")
        
        if 'system' not in self.config:
            raise ValueError("缺少系统配置")
        
        sender_enabled = self.config.get('sender', {}).get('enabled', False)
        receiver_enabled = self.config.get('receiver', {}).get('enabled', False)
        
        if not sender_enabled and not receiver_enabled:
            raise ValueError("发送和接收功能至少需要启用一个")
        
        if sender_enabled:
            self._validate_sender_config()
        
        if receiver_enabled:
            self._validate_receiver_config()
        
        return True
    
    def _validate_sender_config(self):
        """验证发送配置"""
        sender = self.config.get('sender', {})
        
        if 'smtp' not in sender:
            raise ValueError("缺少SMTP配置")
        
        smtp = sender['smtp']
        required_fields = ['host', 'port', 'username']
        for field in required_fields:
            if field not in smtp:
                raise ValueError(f"SMTP配置缺少必需字段: {field}")
        
        if 'recipients' not in sender or 'to' not in sender['recipients']:
            raise ValueError("缺少收件人配置")
        
        if not sender['recipients']['to']:
            raise ValueError("收件人列表不能为空")
    
    def _validate_receiver_config(self):
        """验证接收配置"""
        receiver = self.config.get('receiver', {})
        
        method = receiver.get('method', 'imap')
        if method not in ['imap', 'outlook']:
            raise ValueError(f"不支持的接收方式: {method}")
        
        if method == 'imap':
            if 'imap' not in receiver:
                raise ValueError("缺少IMAP配置")
            imap = receiver['imap']
            required_fields = ['host', 'port', 'username', 'password']
            for field in required_fields:
                if field not in imap:
                    raise ValueError(f"IMAP配置缺少必需字段: {field}")
        
        if 'save' not in receiver:
            raise ValueError("缺少附件保存配置")
        
        if 'path' not in receiver['save']:
            raise ValueError("缺少附件保存路径")
    
    def get_sender_config(self) -> Dict[str, Any]:
        """获取发送配置"""
        return self.config.get('sender', {})
    
    def get_receiver_config(self) -> Dict[str, Any]:
        """获取接收配置"""
        return self.config.get('receiver', {})
    
    def get_system_config(self) -> Dict[str, Any]:
        """获取系统配置"""
        return self.config.get('system', {})
    
    def reload_config(self) -> Dict[str, Any]:
        """重新加载配置"""
        return self.load_config()
    
    @staticmethod
    def create_example_config(output_path: str):
        """创建示例配置文件"""
        example_config = {
            'system': {
                'log_level': 'INFO',
                'log_file': 'logs/app.log',
                'log_max_size': '10MB',
                'log_backup_count': 5
            },
            'sender': {
                'enabled': True,
                'schedule': {
                    'type': 'daily',
                    'time': '09:00'
                },
                'smtp': {
                    'host': 'smtp.example.com',
                    'port': 587,
                    'username': 'user@example.com',
                    'password': 'your_password',
                    'use_tls': True,
                    'timeout': 30
                },
                'email': {
                    'from_name': '发件人名称',
                    'reply_to': 'reply@example.com',
                    'subject': '邮件主题 - {date}',
                    'body_type': 'html',
                    'body': '<html><body><h1>邮件内容</h1></body></html>'
                },
                'recipients': {
                    'to': ['recipient1@example.com', 'recipient2@example.com'],
                    'cc': [],
                    'bcc': []
                },
                'attachments': []
            },
            'receiver': {
                'enabled': True,
                'schedule': {
                    'type': 'hourly',
                    'interval': 2
                },
                'method': 'imap',
                'imap': {
                    'host': 'imap.example.com',
                    'port': 993,
                    'username': 'user@example.com',
                    'password': 'your_password',
                    'use_ssl': True,
                    'folder': 'INBOX'
                },
                'filters': {
                    'from': [],
                    'subject_pattern': '',
                    'has_attachment': True,
                    'unread_only': True
                },
                'save': {
                    'path': 'attachments/receive/{date}/{sender}',
                    'filename_conflict': 'rename',
                    'allowed_extensions': ['pdf', 'xlsx', 'docx'],
                    'max_size': '50MB'
                },
                'after_receive': {
                    'mark_read': True,
                    'move_to': None,
                    'delete': False
                }
            }
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            yaml.dump(example_config, f, allow_unicode=True, default_flow_style=False)
