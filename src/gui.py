"""
GUI界面模块
"""
import os
import sys
import yaml
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import datetime
from typing import Optional


class EmailSchedulerGUI:
    """邮件调度系统GUI界面"""
    
    def __init__(self, config_path: str = None):
        self.config_path = config_path or os.path.join('config', 'config.yaml')
        self.app = None
        self.running = False
        self.log_thread = None
        self.log_file = None
        self.log_pos = 0
        
        self.root = tk.Tk()
        self.root.title("定时邮件收发系统")
        self.root.geometry("900x650")
        self.root.minsize(800, 550)
        
        self._build_menu()
        self._build_main_frame()
        self._build_status_bar()
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _build_menu(self):
        """构建菜单栏"""
        menubar = tk.Menu(self.root)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="启动服务", command=self._start_service)
        file_menu.add_command(label="停止服务", command=self._stop_service)
        file_menu.add_separator()
        file_menu.add_command(label="运行一次(发送)", command=lambda: self._run_once('sender'))
        file_menu.add_command(label="运行一次(接收)", command=lambda: self._run_once('receiver'))
        file_menu.add_command(label="运行一次(全部)", command=lambda: self._run_once('all'))
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self._on_close)
        menubar.add_cascade(label="操作", menu=file_menu)
        
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="系统设置", command=self._open_system_settings)
        settings_menu.add_command(label="发送邮件设置", command=self._open_sender_settings)
        settings_menu.add_command(label="接收邮件设置", command=self._open_receiver_settings)
        settings_menu.add_separator()
        settings_menu.add_command(label="编辑配置文件(YAML)", command=self._open_yaml_editor)
        menubar.add_cascade(label="设置", menu=settings_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="测试配置", command=self._test_config)
        help_menu.add_command(label="关于", command=self._show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def _build_main_frame(self):
        """构建主界面"""
        main_frame = ttk.Frame(self.root, padding=5)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        status_frame = ttk.LabelFrame(main_frame, text="运行状态", padding=10)
        status_frame.pack(fill=tk.X, pady=(0, 5))
        
        info_grid = ttk.Frame(status_frame)
        info_grid.pack(fill=tk.X)
        
        ttk.Label(info_grid, text="服务状态:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.status_label = ttk.Label(info_grid, text="未启动", foreground="gray")
        self.status_label.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(info_grid, text="发送功能:").grid(row=0, column=2, sticky=tk.W, padx=(20, 10))
        self.sender_status = ttk.Label(info_grid, text="未启用", foreground="gray")
        self.sender_status.grid(row=0, column=3, sticky=tk.W)
        
        ttk.Label(info_grid, text="接收功能:").grid(row=0, column=4, sticky=tk.W, padx=(20, 10))
        self.receiver_status = ttk.Label(info_grid, text="未启用", foreground="gray")
        self.receiver_status.grid(row=0, column=5, sticky=tk.W)
        
        ttk.Label(info_grid, text="发送任务:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.sender_next_run = ttk.Label(info_grid, text="-", foreground="gray")
        self.sender_next_run.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        ttk.Label(info_grid, text="接收任务:").grid(row=1, column=3, sticky=tk.W, padx=(20, 10), pady=(5, 0))
        self.receiver_next_run = ttk.Label(info_grid, text="-", foreground="gray")
        self.receiver_next_run.grid(row=1, column=4, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        btn_frame = ttk.Frame(status_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.start_btn = ttk.Button(btn_frame, text="▶ 启动", command=self._start_service, width=12)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.stop_btn = ttk.Button(btn_frame, text="■ 停止", command=self._stop_service, width=12, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(btn_frame, text="运行一次", command=lambda: self._run_once('all'), width=12).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="测试配置", command=self._test_config, width=12).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="设置", command=self._open_sender_settings, width=12).pack(side=tk.RIGHT)
        
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=("Consolas", 9),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        log_btn_frame = ttk.Frame(log_frame)
        log_btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(log_btn_frame, text="清空日志", command=self._clear_log, width=10).pack(side=tk.LEFT)
        ttk.Button(log_btn_frame, text="打开日志文件", command=self._open_log_file, width=12).pack(side=tk.LEFT, padx=(5, 0))
    
    def _build_status_bar(self):
        """构建状态栏"""
        self.statusbar = ttk.Frame(self.root, relief=tk.SUNKEN)
        self.statusbar.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.statusbar_label = ttk.Label(self.statusbar, text="就绪", padding=(5, 2))
        self.statusbar_label.pack(side=tk.LEFT)
        
        self.time_label = ttk.Label(self.statusbar, text="", padding=(5, 2))
        self.time_label.pack(side=tk.RIGHT)
        
        self._update_clock()
    
    def _update_clock(self):
        """更新时钟"""
        self.time_label.config(text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.root.after(1000, self._update_clock)
    
    def _append_log(self, message: str):
        """追加日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _clear_log(self):
        """清空日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def _open_log_file(self):
        """打开日志文件"""
        log_path = os.path.join('logs', 'app.log')
        if os.path.exists(log_path):
            os.startfile(log_path)
        else:
            messagebox.showinfo("提示", "日志文件不存在")
    
    def _load_config(self) -> dict:
        """加载配置文件"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
            return {}
    
    def _save_config(self, config: dict):
        """保存配置文件"""
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
            self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置文件失败: {str(e)}")
    
    def _start_service(self):
        """启动服务"""
        if self.running:
            return
        
        self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 正在启动服务...")
        self.statusbar_label.config(text="正在启动...")
        self.root.update()
        
        def do_start():
            try:
                from .main import EmailSchedulerApp
                self.app = EmailSchedulerApp(self.config_path)
                
                if not self.app.initialize():
                    self.root.after(0, lambda: self._append_log(
                        f"[{datetime.now().strftime('%H:%M:%S')}] 初始化失败"
                    ))
                    self.root.after(0, lambda: self.statusbar_label.config(text="启动失败"))
                    return
                
                self.app.setup_tasks()
                self.app.scheduler.start()
                self.running = True
                
                self.root.after(0, self._on_service_started)
                
            except Exception as e:
                self.root.after(0, lambda: self._append_log(
                    f"[{datetime.now().strftime('%H:%M:%S')}] 启动异常: {str(e)}"
                ))
                self.root.after(0, lambda: self.statusbar_label.config(text="启动异常"))
        
        threading.Thread(target=do_start, daemon=True).start()
    
    def _on_service_started(self):
        """服务启动后的UI更新"""
        self.status_label.config(text="运行中", foreground="green")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.statusbar_label.config(text="服务运行中")
        
        sender_config = self.app.config_manager.get_sender_config()
        receiver_config = self.app.config_manager.get_receiver_config()
        
        if sender_config.get('enabled', False):
            self.sender_status.config(text="已启用", foreground="green")
        else:
            self.sender_status.config(text="未启用", foreground="gray")
        
        if receiver_config.get('enabled', False):
            self.receiver_status.config(text="已启用", foreground="green")
        else:
            self.receiver_status.config(text="未启用", foreground="gray")
        
        self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 服务已启动")
        self._start_log_monitor()
        self._update_task_info()
    
    def _stop_service(self):
        """停止服务"""
        if not self.running:
            return
        
        self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 正在停止服务...")
        
        try:
            if self.app:
                self.app.shutdown()
            self.running = False
            
            self.status_label.config(text="已停止", foreground="orange")
            self.sender_status.config(text="未启用", foreground="gray")
            self.receiver_status.config(text="未启用", foreground="gray")
            self.sender_next_run.config(text="-", foreground="gray")
            self.receiver_next_run.config(text="-", foreground="gray")
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.statusbar_label.config(text="服务已停止")
            
            self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 服务已停止")
        except Exception as e:
            self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 停止异常: {str(e)}")
    
    def _run_once(self, task_type: str = 'all'):
        """运行一次"""
        self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 正在运行一次任务({task_type})...")
        self.statusbar_label.config(text="正在执行...")
        self.root.update()
        
        def do_run():
            try:
                from .main import EmailSchedulerApp
                app = EmailSchedulerApp(self.config_path)
                app.run_once(task_type)
                self.root.after(0, lambda: self._append_log(
                    f"[{datetime.now().strftime('%H:%M:%S')}] 任务执行完成"
                ))
            except Exception as e:
                self.root.after(0, lambda: self._append_log(
                    f"[{datetime.now().strftime('%H:%M:%S')}] 任务执行异常: {str(e)}"
                ))
            finally:
                self.root.after(0, lambda: self.statusbar_label.config(text="就绪"))
        
        threading.Thread(target=do_run, daemon=True).start()
    
    def _test_config(self):
        """测试配置"""
        self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 正在测试配置...")
        self.statusbar_label.config(text="测试中...")
        self.root.update()
        
        def do_test():
            try:
                from .main import EmailSchedulerApp
                app = EmailSchedulerApp(self.config_path)
                success = app.test_config()
                if success:
                    self.root.after(0, lambda: self._append_log(
                        f"[{datetime.now().strftime('%H:%M:%S')}] ✅ 配置测试通过"
                    ))
                    self.root.after(0, lambda: messagebox.showinfo("测试结果", "配置测试通过！"))
                else:
                    self.root.after(0, lambda: self._append_log(
                        f"[{datetime.now().strftime('%H:%M:%S')}] ❌ 配置测试失败"
                    ))
                    self.root.after(0, lambda: messagebox.showerror("测试结果", "配置测试失败，请检查日志"))
            except Exception as e:
                self.root.after(0, lambda: self._append_log(
                    f"[{datetime.now().strftime('%H:%M:%S')}] ❌ 测试异常: {str(e)}"
                ))
                self.root.after(0, lambda: messagebox.showerror("测试结果", f"测试异常: {str(e)}"))
            finally:
                self.root.after(0, lambda: self.statusbar_label.config(text="就绪"))
        
        threading.Thread(target=do_test, daemon=True).start()
    
    def _start_log_monitor(self):
        """启动日志文件监控"""
        log_path = os.path.join('logs', 'app.log')
        if not os.path.exists(log_path):
            return
        
        try:
            self.log_file = open(log_path, 'r', encoding='utf-8', errors='replace')
            self.log_file.seek(0, 2)
            self.log_pos = self.log_file.tell()
        except Exception:
            self.log_file = None
        
        self._check_log_updates()
    
    def _check_log_updates(self):
        """检查日志更新"""
        if not self.running or not self.log_file:
            return
        
        try:
            self.log_file.seek(self.log_pos)
            new_content = self.log_file.read()
            if new_content:
                self.log_pos = self.log_file.tell()
                for line in new_content.strip().split('\n'):
                    if line.strip():
                        self._append_log(line)
        except Exception:
            pass
        
        self.root.after(2000, self._check_log_updates)
    
    def _update_task_info(self):
        """更新任务信息"""
        if not self.running or not self.app or not self.app.scheduler:
            return
        
        try:
            tasks_info = self.app.scheduler.get_all_tasks()
            
            sender_info = tasks_info.get('sender_task')
            if sender_info and sender_info.get('next_run_time'):
                next_time = sender_info['next_run_time'].strftime('%Y-%m-%d %H:%M:%S')
                self.sender_next_run.config(text=next_time, foreground="blue")
            else:
                self.sender_next_run.config(text="未调度", foreground="gray")
            
            receiver_info = tasks_info.get('receiver_task')
            if receiver_info and receiver_info.get('next_run_time'):
                next_time = receiver_info['next_run_time'].strftime('%Y-%m-%d %H:%M:%S')
                self.receiver_next_run.config(text=next_time, foreground="blue")
            else:
                self.receiver_next_run.config(text="未调度", foreground="gray")
        except Exception:
            pass
        
        self.root.after(10000, self._update_task_info)
    
    def _open_system_settings(self):
        """打开系统设置"""
        config = self._load_config()
        if not config:
            return
        
        system = config.get('system', {})
        
        win = tk.Toplevel(self.root)
        win.title("系统设置")
        win.geometry("450x280")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        
        frame = ttk.Frame(win, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)
        
        fields = {}
        row = 0
        
        for label, key, default in [
            ("日志级别:", "log_level", "INFO"),
            ("日志文件:", "log_file", "logs/app.log"),
            ("日志最大大小:", "log_max_size", "10MB"),
            ("日志备份数:", "log_backup_count", 5),
        ]:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky=tk.W, pady=3)
            var = tk.StringVar(value=str(system.get(key, default)))
            if key == "log_level":
                combo = ttk.Combobox(frame, textvariable=var, values=["DEBUG", "INFO", "WARNING", "ERROR"], width=25, state="readonly")
                combo.grid(row=row, column=1, pady=3, padx=(10, 0))
            elif key == "log_backup_count":
                spin = ttk.Spinbox(frame, from_=1, to=20, textvariable=var, width=25)
                spin.grid(row=row, column=1, pady=3, padx=(10, 0))
            else:
                ttk.Entry(frame, textvariable=var, width=28).grid(row=row, column=1, pady=3, padx=(10, 0))
            fields[key] = var
            row += 1
        
        def save():
            for key, var in fields.items():
                val = var.get()
                if key == "log_backup_count":
                    val = int(val)
                system[key] = val
            config['system'] = system
            self._save_config(config)
            win.destroy()
        
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=(15, 0))
        ttk.Button(btn_frame, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=win.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    def _open_sender_settings(self):
        """打开发送邮件设置"""
        config = self._load_config()
        if not config:
            return
        
        sender = config.setdefault('sender', {})
        sender.setdefault('enabled', False)
        sender.setdefault('schedule', {'type': 'hourly', 'interval': 1, 'time': '09:00'})
        sender.setdefault('smtp', {'host': '', 'port': 465, 'username': '', 'password': '', 'use_tls': True, 'timeout': 30})
        sender.setdefault('email', {'from_name': '', 'reply_to': '', 'subject': '', 'body_type': 'html', 'body': ''})
        sender.setdefault('recipients', {'to': [], 'cc': [], 'bcc': []})
        sender.setdefault('attachments', [])
        
        win = tk.Toplevel(self.root)
        win.title("发送邮件设置")
        win.geometry("600x620")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        
        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        basic_frame = ttk.Frame(notebook, padding=10)
        notebook.add(basic_frame, text="基本设置")
        
        row = 0
        enabled_var = tk.BooleanVar(value=sender.get('enabled', False))
        ttk.Checkbutton(basic_frame, text="启用发送功能", variable=enabled_var).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=5)
        row += 1
        
        schedule = sender.get('schedule', {})
        ttk.Label(basic_frame, text="── 定时设置 ──", font=("", 9, "bold")).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        row += 1
        
        ttk.Label(basic_frame, text="循环类型:").grid(row=row, column=0, sticky=tk.W, pady=3)
        schedule_type_var = tk.StringVar(value=schedule.get('type', 'hourly'))
        type_combo = ttk.Combobox(basic_frame, textvariable=schedule_type_var,
                                   values=["hourly", "daily", "weekly", "once"], width=20, state="readonly")
        type_combo.grid(row=row, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        row += 1
        
        ttk.Label(basic_frame, text="间隔(小时):").grid(row=row, column=0, sticky=tk.W, pady=3)
        interval_var = tk.StringVar(value=str(schedule.get('interval', 1)))
        ttk.Entry(basic_frame, textvariable=interval_var, width=22).grid(row=row, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        row += 1
        
        ttk.Label(basic_frame, text="执行时间:").grid(row=row, column=0, sticky=tk.W, pady=3)
        time_var = tk.StringVar(value=str(schedule.get('time', '09:00')))
        ttk.Entry(basic_frame, textvariable=time_var, width=22).grid(row=row, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        row += 1
        
        ttk.Label(basic_frame, text="说明: hourly=每小时, daily=每天, weekly=每周, once=一次",
                  foreground="gray").grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=3)
        row += 1
        
        smtp_frame = ttk.Frame(notebook, padding=10)
        notebook.add(smtp_frame, text="SMTP设置")
        
        smtp = sender.get('smtp', {})
        smtp_vars = {}
        srow = 0
        
        for label, key, default, show in [
            ("SMTP服务器:", "host", "", ""),
            ("端口:", "port", 465, ""),
            ("用户名:", "username", "", ""),
            ("密码:", "password", "", "*"),
            ("超时(秒):", "timeout", 30, ""),
        ]:
            ttk.Label(smtp_frame, text=label).grid(row=srow, column=0, sticky=tk.W, pady=3)
            var = tk.StringVar(value=str(smtp.get(key, default)))
            entry = ttk.Entry(smtp_frame, textvariable=var, width=30, show=show if show else "")
            entry.grid(row=srow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
            smtp_vars[key] = var
            srow += 1
        
        tls_var = tk.BooleanVar(value=smtp.get('use_tls', True))
        ttk.Checkbutton(smtp_frame, text="使用TLS加密", variable=tls_var).grid(row=srow, column=0, columnspan=2, sticky=tk.W, pady=5)
        smtp_vars['use_tls'] = tls_var
        
        email_frame = ttk.Frame(notebook, padding=10)
        notebook.add(email_frame, text="邮件内容")
        
        email = sender.get('email', {})
        email_vars = {}
        erow = 0
        
        for label, key, default in [
            ("发件人名称:", "from_name", ""),
            ("回复地址:", "reply_to", ""),
            ("邮件主题:", "subject", ""),
        ]:
            ttk.Label(email_frame, text=label).grid(row=erow, column=0, sticky=tk.W, pady=3)
            var = tk.StringVar(value=str(email.get(key, default)))
            ttk.Entry(email_frame, textvariable=var, width=35).grid(row=erow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
            email_vars[key] = var
            erow += 1
        
        ttk.Label(email_frame, text="正文格式:").grid(row=erow, column=0, sticky=tk.W, pady=3)
        body_type_var = tk.StringVar(value=email.get('body_type', 'html'))
        ttk.Combobox(email_frame, textvariable=body_type_var, values=["html", "text"], width=10, state="readonly").grid(
            row=erow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        email_vars['body_type'] = body_type_var
        erow += 1
        
        ttk.Label(email_frame, text="邮件正文:").grid(row=erow, column=0, sticky=tk.NW, pady=3)
        body_var = tk.StringVar(value=str(email.get('body', '')))
        body_text = scrolledtext.ScrolledText(email_frame, width=35, height=8, wrap=tk.WORD)
        body_text.grid(row=erow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        body_text.insert(1.0, str(email.get('body', '')))
        erow += 1
        
        recip_frame = ttk.Frame(notebook, padding=10)
        notebook.add(recip_frame, text="收件人")
        
        recipients = sender.get('recipients', {})
        rrow = 0
        
        ttk.Label(recip_frame, text="收件人(每行一个):").grid(row=rrow, column=0, sticky=tk.NW, pady=3)
        to_text = scrolledtext.ScrolledText(recip_frame, width=35, height=4, wrap=tk.WORD)
        to_text.grid(row=rrow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        to_list = recipients.get('to', [])
        to_text.insert(1.0, '\n'.join(to_list if isinstance(to_list, list) else [to_list]))
        rrow += 1
        
        ttk.Label(recip_frame, text="抄送(每行一个):").grid(row=rrow, column=0, sticky=tk.NW, pady=3)
        cc_text = scrolledtext.ScrolledText(recip_frame, width=35, height=3, wrap=tk.WORD)
        cc_text.grid(row=rrow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        cc_list = recipients.get('cc', [])
        cc_text.insert(1.0, '\n'.join(cc_list if isinstance(cc_list, list) else [cc_list]))
        rrow += 1
        
        ttk.Label(recip_frame, text="密送(每行一个):").grid(row=rrow, column=0, sticky=tk.NW, pady=3)
        bcc_text = scrolledtext.ScrolledText(recip_frame, width=35, height=3, wrap=tk.WORD)
        bcc_text.grid(row=rrow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        bcc_list = recipients.get('bcc', [])
        bcc_text.insert(1.0, '\n'.join(bcc_list if isinstance(bcc_list, list) else [bcc_list]))
        rrow += 1
        
        attach_frame = ttk.Frame(notebook, padding=10)
        notebook.add(attach_frame, text="附件")
        
        ttk.Label(attach_frame, text="附件路径(每行一个，支持通配符如 *.pdf):").pack(anchor=tk.W)
        attach_text = scrolledtext.ScrolledText(attach_frame, width=50, height=10, wrap=tk.WORD)
        attach_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        attachments = sender.get('attachments', [])
        attach_lines = []
        for att in attachments:
            if isinstance(att, dict):
                path = att.get('path', '')
                required = att.get('required', False)
                attach_lines.append(f"{path} (required={required})")
            else:
                attach_lines.append(str(att))
        attach_text.insert(1.0, '\n'.join(attach_lines))
        
        ttk.Label(attach_frame, text="格式: 路径 (required=true/false)", foreground="gray").pack(anchor=tk.W)
        
        def save():
            sender['enabled'] = enabled_var.get()
            sender['schedule'] = {
                'type': schedule_type_var.get(),
                'interval': int(interval_var.get()) if interval_var.get().isdigit() else 1,
                'time': time_var.get()
            }
            
            for key, var in smtp_vars.items():
                if key == 'use_tls':
                    sender['smtp'][key] = var.get()
                elif key == 'port':
                    sender['smtp'][key] = int(var.get()) if var.get().isdigit() else 465
                elif key == 'timeout':
                    sender['smtp'][key] = int(var.get()) if var.get().isdigit() else 30
                else:
                    sender['smtp'][key] = var.get()
            
            for key, var in email_vars.items():
                if key == 'body_type':
                    sender['email'][key] = var.get()
                else:
                    sender['email'][key] = var.get()
            sender['email']['body'] = body_text.get(1.0, tk.END).strip()
            
            to_content = [x.strip() for x in to_text.get(1.0, tk.END).strip().split('\n') if x.strip()]
            cc_content = [x.strip() for x in cc_text.get(1.0, tk.END).strip().split('\n') if x.strip()]
            bcc_content = [x.strip() for x in bcc_text.get(1.0, tk.END).strip().split('\n') if x.strip()]
            sender['recipients'] = {'to': to_content, 'cc': cc_content, 'bcc': bcc_content}
            
            attach_content = []
            for line in attach_text.get(1.0, tk.END).strip().split('\n'):
                line = line.strip()
                if not line:
                    continue
                if '(required=' in line:
                    parts = line.split('(required=')
                    path = parts[0].strip()
                    req = parts[1].replace(')', '').strip().lower() == 'true'
                    attach_content.append({'path': path, 'required': req})
                else:
                    attach_content.append({'path': line, 'required': False})
            sender['attachments'] = attach_content
            
            config['sender'] = sender
            self._save_config(config)
            win.destroy()
        
        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=win.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    def _open_receiver_settings(self):
        """打开接收邮件设置"""
        config = self._load_config()
        if not config:
            return
        
        receiver = config.setdefault('receiver', {})
        receiver.setdefault('enabled', False)
        receiver.setdefault('schedule', {'type': 'hourly', 'interval': 1})
        receiver.setdefault('method', 'imap')
        receiver.setdefault('imap', {'host': '', 'port': 993, 'username': '', 'password': '', 'use_ssl': True, 'folder': 'INBOX'})
        receiver.setdefault('outlook', {'profile': 'Outlook', 'mailbox': '', 'folder': 'Inbox'})
        receiver.setdefault('filters', {'from': [], 'subject_pattern': '', 'has_attachment': True, 'unread_only': True, 'latest_only': False, 'max_emails': 100})
        receiver.setdefault('save', {'path': 'attachments/receive/{date}/{sender}', 'filename_conflict': 'rename', 'allowed_extensions': ['pdf', 'xlsx', 'docx'], 'max_size': '50MB'})
        receiver.setdefault('after_receive', {'mark_read': True, 'move_to': None, 'delete': False})
        
        win = tk.Toplevel(self.root)
        win.title("接收邮件设置")
        win.geometry("600x620")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        
        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        basic_frame = ttk.Frame(notebook, padding=10)
        notebook.add(basic_frame, text="基本设置")
        
        row = 0
        enabled_var = tk.BooleanVar(value=receiver.get('enabled', False))
        ttk.Checkbutton(basic_frame, text="启用接收功能", variable=enabled_var).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=5)
        row += 1
        
        schedule = receiver.get('schedule', {})
        ttk.Label(basic_frame, text="── 定时设置 ──", font=("", 9, "bold")).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        row += 1
        
        ttk.Label(basic_frame, text="循环类型:").grid(row=row, column=0, sticky=tk.W, pady=3)
        schedule_type_var = tk.StringVar(value=schedule.get('type', 'hourly'))
        ttk.Combobox(basic_frame, textvariable=schedule_type_var,
                     values=["hourly", "daily", "weekly", "once"], width=20, state="readonly").grid(
            row=row, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        row += 1
        
        ttk.Label(basic_frame, text="间隔(小时):").grid(row=row, column=0, sticky=tk.W, pady=3)
        interval_var = tk.StringVar(value=str(schedule.get('interval', 1)))
        ttk.Entry(basic_frame, textvariable=interval_var, width=22).grid(row=row, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        row += 1
        
        ttk.Label(basic_frame, text="执行时间:").grid(row=row, column=0, sticky=tk.W, pady=3)
        time_var = tk.StringVar(value=str(schedule.get('time', '09:00')))
        ttk.Entry(basic_frame, textvariable=time_var, width=22).grid(row=row, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        row += 1
        
        ttk.Label(basic_frame, text="── 接收方式 ──", font=("", 9, "bold")).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        row += 1
        
        ttk.Label(basic_frame, text="接收方式:").grid(row=row, column=0, sticky=tk.W, pady=3)
        method_var = tk.StringVar(value=receiver.get('method', 'imap'))
        ttk.Combobox(basic_frame, textvariable=method_var, values=["imap", "outlook"], width=20, state="readonly").grid(
            row=row, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        row += 1
        
        imap_frame = ttk.Frame(notebook, padding=10)
        notebook.add(imap_frame, text="IMAP设置")
        
        imap = receiver.get('imap', {})
        imap_vars = {}
        irow = 0
        
        for label, key, default in [
            ("IMAP服务器:", "host", ""),
            ("端口:", "port", 993),
            ("用户名:", "username", ""),
            ("密码:", "password", ""),
            ("文件夹:", "folder", "INBOX"),
        ]:
            ttk.Label(imap_frame, text=label).grid(row=irow, column=0, sticky=tk.W, pady=3)
            show = "*" if key == "password" else ""
            var = tk.StringVar(value=str(imap.get(key, default)))
            ttk.Entry(imap_frame, textvariable=var, width=30, show=show).grid(row=irow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
            imap_vars[key] = var
            irow += 1
        
        ssl_var = tk.BooleanVar(value=imap.get('use_ssl', True))
        ttk.Checkbutton(imap_frame, text="使用SSL加密", variable=ssl_var).grid(row=irow, column=0, columnspan=2, sticky=tk.W, pady=5)
        imap_vars['use_ssl'] = ssl_var
        
        outlook_frame = ttk.Frame(notebook, padding=10)
        notebook.add(outlook_frame, text="Outlook设置")
        
        outlook = receiver.get('outlook', {})
        outlook_vars = {}
        orow = 0
        
        for label, key, default in [
            ("配置文件名:", "profile", "Outlook"),
            ("邮箱地址:", "mailbox", ""),
            ("文件夹:", "folder", "Inbox"),
        ]:
            ttk.Label(outlook_frame, text=label).grid(row=orow, column=0, sticky=tk.W, pady=3)
            var = tk.StringVar(value=str(outlook.get(key, default)))
            ttk.Entry(outlook_frame, textvariable=var, width=30).grid(row=orow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
            outlook_vars[key] = var
            orow += 1
        
        ttk.Label(outlook_frame, text="注意: Outlook方式需要本地安装Outlook客户端",
                  foreground="gray").grid(row=orow, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        filter_frame = ttk.Frame(notebook, padding=10)
        notebook.add(filter_frame, text="筛选条件")
        
        filters = receiver.get('filters', {})
        frow = 0
        
        ttk.Label(filter_frame, text="发件人(每行一个):").grid(row=frow, column=0, sticky=tk.NW, pady=3)
        from_text = scrolledtext.ScrolledText(filter_frame, width=35, height=3, wrap=tk.WORD)
        from_text.grid(row=frow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        from_list = filters.get('from', [])
        from_text.insert(1.0, '\n'.join(from_list if isinstance(from_list, list) else [from_list]))
        frow += 1
        
        ttk.Label(filter_frame, text="主题正则:").grid(row=frow, column=0, sticky=tk.W, pady=3)
        subject_var = tk.StringVar(value=str(filters.get('subject_pattern', '')))
        ttk.Entry(filter_frame, textvariable=subject_var, width=35).grid(row=frow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        frow += 1
        
        ttk.Label(filter_frame, text="  提示: .*关键词.* 匹配包含关键词的主题",
                  foreground="gray").grid(row=frow, column=0, columnspan=2, sticky=tk.W, pady=2)
        frow += 1
        
        has_attach_var = tk.BooleanVar(value=filters.get('has_attachment', True))
        ttk.Checkbutton(filter_frame, text="只接收有附件的邮件", variable=has_attach_var).grid(row=frow, column=0, columnspan=2, sticky=tk.W, pady=3)
        frow += 1
        
        unread_var = tk.BooleanVar(value=filters.get('unread_only', True))
        ttk.Checkbutton(filter_frame, text="只接收未读邮件", variable=unread_var).grid(row=frow, column=0, columnspan=2, sticky=tk.W, pady=3)
        frow += 1
        
        latest_var = tk.BooleanVar(value=filters.get('latest_only', False))
        ttk.Checkbutton(filter_frame, text="只保留最新的一封邮件", variable=latest_var).grid(row=frow, column=0, columnspan=2, sticky=tk.W, pady=3)
        frow += 1
        
        ttk.Label(filter_frame, text="最多获取邮件数:").grid(row=frow, column=0, sticky=tk.W, pady=3)
        max_emails_var = tk.StringVar(value=str(filters.get('max_emails', 100)))
        ttk.Spinbox(filter_frame, from_=1, to=1000, textvariable=max_emails_var, width=10).grid(
            row=frow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        frow += 1
        
        save_frame = ttk.Frame(notebook, padding=10)
        notebook.add(save_frame, text="保存设置")
        
        save_config = receiver.get('save', {})
        srow = 0
        
        ttk.Label(save_frame, text="保存路径:").grid(row=srow, column=0, sticky=tk.W, pady=3)
        save_path_var = tk.StringVar(value=str(save_config.get('path', 'attachments/receive/{date}/{sender}')))
        ttk.Entry(save_frame, textvariable=save_path_var, width=35).grid(row=srow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        srow += 1
        
        ttk.Label(save_frame, text="  支持变量: {date} {time} {sender} {year} {month} {day}",
                  foreground="gray").grid(row=srow, column=0, columnspan=2, sticky=tk.W, pady=2)
        srow += 1
        
        ttk.Label(save_frame, text="重名处理:").grid(row=srow, column=0, sticky=tk.W, pady=3)
        conflict_var = tk.StringVar(value=str(save_config.get('filename_conflict', 'rename')))
        ttk.Combobox(save_frame, textvariable=conflict_var, values=["rename", "overwrite", "skip"], width=15, state="readonly").grid(
            row=srow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        srow += 1
        
        ttk.Label(save_frame, text="允许的扩展名(每行一个):").grid(row=srow, column=0, sticky=tk.NW, pady=3)
        ext_text = scrolledtext.ScrolledText(save_frame, width=35, height=4, wrap=tk.WORD)
        ext_text.grid(row=srow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        ext_list = save_config.get('allowed_extensions', [])
        ext_text.insert(1.0, '\n'.join(ext_list if isinstance(ext_list, list) else [ext_list]))
        srow += 1
        
        ttk.Label(save_frame, text="最大附件大小:").grid(row=srow, column=0, sticky=tk.W, pady=3)
        max_size_var = tk.StringVar(value=str(save_config.get('max_size', '50MB')))
        ttk.Entry(save_frame, textvariable=max_size_var, width=15).grid(row=srow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        srow += 1
        
        after_frame = ttk.Frame(notebook, padding=10)
        notebook.add(after_frame, text="接收后处理")
        
        after_config = receiver.get('after_receive', {})
        arow = 0
        
        mark_read_var = tk.BooleanVar(value=after_config.get('mark_read', True))
        ttk.Checkbutton(after_frame, text="标记为已读", variable=mark_read_var).grid(row=arow, column=0, columnspan=2, sticky=tk.W, pady=5)
        arow += 1
        
        ttk.Label(after_frame, text="移动到文件夹:").grid(row=arow, column=0, sticky=tk.W, pady=3)
        move_var = tk.StringVar(value=str(after_config.get('move_to', '') or ''))
        ttk.Entry(after_frame, textvariable=move_var, width=25).grid(row=arow, column=1, sticky=tk.W, pady=3, padx=(10, 0))
        arow += 1
        
        delete_var = tk.BooleanVar(value=after_config.get('delete', False))
        ttk.Checkbutton(after_frame, text="接收后删除原邮件", variable=delete_var).grid(row=arow, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        def save():
            receiver['enabled'] = enabled_var.get()
            receiver['schedule'] = {
                'type': schedule_type_var.get(),
                'interval': int(interval_var.get()) if interval_var.get().isdigit() else 1,
                'time': time_var.get()
            }
            receiver['method'] = method_var.get()
            
            for key, var in imap_vars.items():
                if key == 'use_ssl':
                    receiver['imap'][key] = var.get()
                elif key == 'port':
                    receiver['imap'][key] = int(var.get()) if var.get().isdigit() else 993
                else:
                    receiver['imap'][key] = var.get()
            
            for key, var in outlook_vars.items():
                receiver['outlook'][key] = var.get()
            
            from_content = [x.strip() for x in from_text.get(1.0, tk.END).strip().split('\n') if x.strip()]
            receiver['filters'] = {
                'from': from_content,
                'subject_pattern': subject_var.get(),
                'has_attachment': has_attach_var.get(),
                'unread_only': unread_var.get(),
                'latest_only': latest_var.get(),
                'max_emails': int(max_emails_var.get()) if max_emails_var.get().isdigit() else 100,
            }
            
            ext_content = [x.strip() for x in ext_text.get(1.0, tk.END).strip().split('\n') if x.strip()]
            receiver['save'] = {
                'path': save_path_var.get(),
                'filename_conflict': conflict_var.get(),
                'allowed_extensions': ext_content,
                'max_size': max_size_var.get(),
            }
            
            move_to_val = move_var.get().strip() if move_var.get().strip() else None
            receiver['after_receive'] = {
                'mark_read': mark_read_var.get(),
                'move_to': move_to_val,
                'delete': delete_var.get(),
            }
            
            config['receiver'] = receiver
            self._save_config(config)
            win.destroy()
        
        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=win.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    def _open_yaml_editor(self):
        """打开YAML编辑器"""
        config = self._load_config()
        if not config:
            config = {}
        
        win = tk.Toplevel(self.root)
        win.title("编辑配置文件 (YAML)")
        win.geometry("700x550")
        win.transient(self.root)
        win.grab_set()
        
        text = scrolledtext.ScrolledText(win, wrap=tk.WORD, font=("Consolas", 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        yaml_content = yaml.dump(config, allow_unicode=True, default_flow_style=False)
        text.insert(1.0, yaml_content)
        
        def save():
            try:
                new_config = yaml.safe_load(text.get(1.0, tk.END))
                if new_config is None:
                    messagebox.showerror("错误", "配置内容为空")
                    return
                self._save_config(new_config)
                win.destroy()
            except yaml.YAMLError as e:
                messagebox.showerror("YAML格式错误", str(e))
        
        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="保存", command=save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=win.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    def _show_about(self):
        """显示关于"""
        messagebox.showinfo(
            "关于",
            "定时邮件收发系统 v1.0\n\n"
            "功能：\n"
            "• 定时发送邮件\n"
            "• 定时接收邮件附件\n"
            "• 支持 IMAP / Outlook\n"
            "• 支持多种定时规则\n\n"
            "GitHub: https://github.com/jfishpw/AUTOSENDREMAIL"
        )
    
    def _on_close(self):
        """关闭窗口"""
        if self.running:
            if not messagebox.askyesno("确认", "服务正在运行，确定要退出吗？"):
                return
            self._stop_service()
        
        if self.log_file:
            try:
                self.log_file.close()
            except Exception:
                pass
        
        self.root.destroy()
    
    def run(self):
        """运行GUI"""
        config = self._load_config()
        if config:
            self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 配置文件已加载: {self.config_path}")
            sender_enabled = config.get('sender', {}).get('enabled', False)
            receiver_enabled = config.get('receiver', {}).get('enabled', False)
            self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] 发送功能: {'已启用' if sender_enabled else '未启用'}, 接收功能: {'已启用' if receiver_enabled else '未启用'}")
        else:
            self._append_log(f"[{datetime.now().strftime('%H:%M:%S')}] ⚠ 配置文件加载失败，请检查设置")
        
        self.root.mainloop()
