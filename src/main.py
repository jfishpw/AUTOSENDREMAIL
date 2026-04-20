"""
主程序模块
"""
import os
import sys
import signal
import argparse
from typing import Optional
from .config_manager import ConfigManager
from .logger import setup_logger, get_logger
from .sender import EmailSender
from .receiver import create_receiver
from .scheduler import TaskScheduler


class EmailSchedulerApp:
    """邮件调度应用"""
    
    def __init__(self, config_path: str = None):
        """
        初始化应用
        
        Args:
            config_path: 配置文件路径
        """
        self.config_manager = ConfigManager(config_path)
        self.config = None
        self.logger = None
        self.scheduler = None
        self.sender = None
        self.receiver = None
        self.running = False
    
    def initialize(self) -> bool:
        """
        初始化应用
        
        Returns:
            是否初始化成功
        """
        try:
            self.config = self.config_manager.load_config()
            
            system_config = self.config_manager.get_system_config()
            self.logger = setup_logger(
                log_level=system_config.get('log_level', 'INFO'),
                log_file=system_config.get('log_file'),
                log_max_size=system_config.get('log_max_size', '10MB'),
                log_backup_count=system_config.get('log_backup_count', 5)
            )
            
            self.scheduler = TaskScheduler()
            
            sender_config = self.config_manager.get_sender_config()
            if sender_config.get('enabled', False):
                self.sender = EmailSender(sender_config.get('smtp', {}))
            
            receiver_config = self.config_manager.get_receiver_config()
            if receiver_config.get('enabled', False):
                self.receiver = create_receiver(receiver_config)
            
            self.logger.info("应用初始化成功")
            return True
            
        except Exception as e:
            print(f"初始化失败: {str(e)}")
            return False
    
    def setup_tasks(self, sender_only: bool = False, receiver_only: bool = False):
        """
        设置定时任务
        
        Args:
            sender_only: 只启用发送功能
            receiver_only: 只启用接收功能
        """
        if not sender_only and not receiver_only:
            sender_only = True
            receiver_only = True
        
        if sender_only and self.sender:
            sender_config = self.config_manager.get_sender_config()
            schedule_config = sender_config.get('schedule', {})
            
            self.scheduler.add_sender_task(
                task_func=self._send_email_task,
                schedule_config=schedule_config,
                task_id='sender_task'
            )
        
        if receiver_only and self.receiver:
            receiver_config = self.config_manager.get_receiver_config()
            schedule_config = receiver_config.get('schedule', {})
            
            self.scheduler.add_receiver_task(
                task_func=self._receive_email_task,
                schedule_config=schedule_config,
                task_id='receiver_task'
            )
    
    def _send_email_task(self):
        """发送邮件任务"""
        try:
            self.logger.info("开始执行发送邮件任务")
            
            sender_config = self.config_manager.get_sender_config()
            email_config = sender_config.get('email', {})
            recipients_config = sender_config.get('recipients', {})
            attachments_config = sender_config.get('attachments', [])
            
            success = self.sender.send_with_config(
                email_config=email_config,
                recipients_config=recipients_config,
                attachments_config=attachments_config
            )
            
            if success:
                self.logger.info("发送邮件任务执行成功")
            else:
                self.logger.error("发送邮件任务执行失败")
            
        except Exception as e:
            self.logger.error(f"发送邮件任务异常: {str(e)}")
    
    def _receive_email_task(self):
        """接收邮件任务"""
        try:
            self.logger.info("开始执行接收邮件任务")
            
            processed, saved = self.receiver.process_emails()
            
            self.logger.info(
                f"接收邮件任务执行完成: 处理 {processed} 封邮件, "
                f"保存 {saved} 个附件"
            )
            
        except Exception as e:
            self.logger.error(f"接收邮件任务异常: {str(e)}")
    
    def run(self, sender_only: bool = False, receiver_only: bool = False):
        """
        运行应用
        
        Args:
            sender_only: 只运行发送功能
            receiver_only: 只运行接收功能
        """
        if not self.initialize():
            sys.exit(1)
        
        self.setup_tasks(sender_only, receiver_only)
        
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)
        
        self.scheduler.start()
        self.running = True
        
        self.logger.info("应用开始运行")
        self._print_task_info()
        
        try:
            while self.running:
                import time
                time.sleep(1)
        except KeyboardInterrupt:
            self.logger.info("收到中断信号")
        finally:
            self.shutdown()
    
    def _signal_handler(self, signum, frame):
        """信号处理函数"""
        self.logger.info(f"收到信号 {signum}")
        self.running = False
    
    def _print_task_info(self):
        """打印任务信息"""
        tasks_info = self.scheduler.get_all_tasks()
        for task_id, info in tasks_info.items():
            if info:
                self.logger.info(
                    f"任务 [{task_id}]: 下次运行时间 {info['next_run_time']}"
                )
    
    def shutdown(self):
        """关闭应用"""
        self.logger.info("正在关闭应用...")
        
        if self.scheduler:
            self.scheduler.stop()
        
        if self.sender:
            self.sender.disconnect()
        
        if self.receiver:
            self.receiver.disconnect()
        
        self.logger.info("应用已关闭")
    
    def test_config(self) -> bool:
        """
        测试配置
        
        Returns:
            配置是否有效
        """
        try:
            if not self.initialize():
                return False
            
            self.logger.info("配置验证成功")
            
            if self.sender:
                self.logger.info("测试SMTP连接...")
                if self.sender.test_connection():
                    self.logger.info("SMTP连接测试成功")
                else:
                    self.logger.error("SMTP连接测试失败")
                    return False
            
            if self.receiver:
                self.logger.info("测试邮件接收连接...")
                if self.receiver.connect():
                    self.logger.info("邮件接收连接测试成功")
                    self.receiver.disconnect()
                else:
                    self.logger.error("邮件接收连接测试失败")
                    return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"配置测试失败: {str(e)}")
            return False
    
    def run_once(self, task_type: str = 'all'):
        """
        运行一次任务
        
        Args:
            task_type: 任务类型 ('sender', 'receiver', 'all')
        """
        if not self.initialize():
            sys.exit(1)
        
        if task_type in ['sender', 'all'] and self.sender:
            self._send_email_task()
        
        if task_type in ['receiver', 'all'] and self.receiver:
            self._receive_email_task()
        
        self.shutdown()


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='定时邮件收发系统',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        '--config',
        type=str,
        default=None,
        help='配置文件路径 (默认: config/config.yaml)'
    )
    
    parser.add_argument(
        '--sender-only',
        action='store_true',
        help='只运行发送功能'
    )
    
    parser.add_argument(
        '--receiver-only',
        action='store_true',
        help='只运行接收功能'
    )
    
    parser.add_argument(
        '--test-config',
        action='store_true',
        help='测试配置文件'
    )
    
    parser.add_argument(
        '--run-once',
        action='store_true',
        help='只运行一次任务'
    )
    
    parser.add_argument(
        '--task-type',
        type=str,
        choices=['sender', 'receiver', 'all'],
        default='all',
        help='运行的任务类型 (与 --run-once 一起使用)'
    )
    
    parser.add_argument(
        '--gui',
        action='store_true',
        help='启动图形界面'
    )
    
    args = parser.parse_args()
    
    if args.gui:
        from .gui import EmailSchedulerGUI
        gui = EmailSchedulerGUI(args.config)
        gui.run()
        return
    
    app = EmailSchedulerApp(args.config)
    
    if args.test_config:
        success = app.test_config()
        sys.exit(0 if success else 1)
    
    if args.run_once:
        app.run_once(args.task_type)
        sys.exit(0)
    
    app.run(
        sender_only=args.sender_only,
        receiver_only=args.receiver_only
    )


if __name__ == '__main__':
    main()
