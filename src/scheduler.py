"""
定时调度模块
"""
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.interval import IntervalTrigger
from apscheduler.triggers.date import DateTrigger
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, Callable
from .logger import get_logger


class TaskScheduler:
    """定时任务调度器"""
    
    def __init__(self):
        """初始化调度器"""
        self.logger = get_logger()
        self.scheduler = BackgroundScheduler()
        self.tasks = {}
    
    def start(self):
        """启动调度器"""
        if not self.scheduler.running:
            self.scheduler.start()
            self.logger.info("调度器已启动")
    
    def stop(self):
        """停止调度器"""
        if self.scheduler.running:
            self.scheduler.shutdown()
            self.logger.info("调度器已停止")
    
    def add_sender_task(
        self,
        task_func: Callable,
        schedule_config: Dict[str, Any],
        task_id: str = 'sender_task'
    ) -> bool:
        """
        添加发送任务
        
        Args:
            task_func: 任务函数
            schedule_config: 定时配置
            task_id: 任务ID
        
        Returns:
            是否添加成功
        """
        try:
            trigger = self._create_trigger(schedule_config)
            
            self.scheduler.add_job(
                task_func,
                trigger=trigger,
                id=task_id,
                name='邮件发送任务',
                misfire_grace_time=3600
            )
            
            self.tasks[task_id] = {
                'type': 'sender',
                'config': schedule_config,
                'func': task_func
            }
            
            self.logger.info(f"添加发送任务: {task_id}")
            return True
            
        except Exception as e:
            self.logger.error(f"添加发送任务失败: {str(e)}")
            return False
    
    def add_receiver_task(
        self,
        task_func: Callable,
        schedule_config: Dict[str, Any],
        task_id: str = 'receiver_task'
    ) -> bool:
        """
        添加接收任务
        
        Args:
            task_func: 任务函数
            schedule_config: 定时配置
            task_id: 任务ID
        
        Returns:
            是否添加成功
        """
        try:
            trigger = self._create_trigger(schedule_config)
            
            self.scheduler.add_job(
                task_func,
                trigger=trigger,
                id=task_id,
                name='邮件接收任务',
                misfire_grace_time=3600
            )
            
            self.tasks[task_id] = {
                'type': 'receiver',
                'config': schedule_config,
                'func': task_func
            }
            
            self.logger.info(f"添加接收任务: {task_id}")
            return True
            
        except Exception as e:
            self.logger.error(f"添加接收任务失败: {str(e)}")
            return False
    
    def _create_trigger(self, schedule_config: Dict[str, Any]):
        """
        创建触发器
        
        Args:
            schedule_config: 定时配置
        
        Returns:
            触发器对象
        """
        schedule_type = schedule_config.get('type', 'daily')
        
        if schedule_type == 'hourly':
            interval = schedule_config.get('interval', 1)
            return IntervalTrigger(hours=interval)
        
        elif schedule_type == 'daily':
            time_str = schedule_config.get('time', '09:00')
            hour, minute = map(int, time_str.split(':'))
            return CronTrigger(hour=hour, minute=minute)
        
        elif schedule_type == 'weekly':
            time_str = schedule_config.get('time', '09:00')
            weekdays = schedule_config.get('weekdays', [1, 2, 3, 4, 5])
            hour, minute = map(int, time_str.split(':'))
            
            day_of_week = ','.join([str(d - 1) for d in weekdays])
            return CronTrigger(
                day_of_week=day_of_week,
                hour=hour,
                minute=minute
            )
        
        elif schedule_type == 'once':
            datetime_str = schedule_config.get('datetime')
            if datetime_str:
                run_date = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')
            else:
                time_str = schedule_config.get('time', '09:00')
                date_str = schedule_config.get('date', datetime.now().strftime('%Y-%m-%d'))
                run_date = datetime.strptime(f"{date_str} {time_str}", '%Y-%m-%d %H:%M')
            
            return DateTrigger(run_date=run_date)
        
        else:
            raise ValueError(f"不支持的定时类型: {schedule_type}")
    
    def remove_task(self, task_id: str) -> bool:
        """
        移除任务
        
        Args:
            task_id: 任务ID
        
        Returns:
            是否移除成功
        """
        try:
            self.scheduler.remove_job(task_id)
            if task_id in self.tasks:
                del self.tasks[task_id]
            self.logger.info(f"移除任务: {task_id}")
            return True
        except Exception as e:
            self.logger.error(f"移除任务失败: {str(e)}")
            return False
    
    def pause_task(self, task_id: str) -> bool:
        """
        暂停任务
        
        Args:
            task_id: 任务ID
        
        Returns:
            是否暂停成功
        """
        try:
            self.scheduler.pause_job(task_id)
            self.logger.info(f"暂停任务: {task_id}")
            return True
        except Exception as e:
            self.logger.error(f"暂停任务失败: {str(e)}")
            return False
    
    def resume_task(self, task_id: str) -> bool:
        """
        恢复任务
        
        Args:
            task_id: 任务ID
        
        Returns:
            是否恢复成功
        """
        try:
            self.scheduler.resume_job(task_id)
            self.logger.info(f"恢复任务: {task_id}")
            return True
        except Exception as e:
            self.logger.error(f"恢复任务失败: {str(e)}")
            return False
    
    def get_task_info(self, task_id: str) -> Optional[Dict[str, Any]]:
        """
        获取任务信息
        
        Args:
            task_id: 任务ID
        
        Returns:
            任务信息字典
        """
        job = self.scheduler.get_job(task_id)
        if job:
            return {
                'id': job.id,
                'name': job.name,
                'next_run_time': job.next_run_time,
                'trigger': str(job.trigger)
            }
        return None
    
    def get_all_tasks(self) -> Dict[str, Dict[str, Any]]:
        """
        获取所有任务信息
        
        Returns:
            任务信息字典
        """
        tasks_info = {}
        for task_id in self.tasks:
            tasks_info[task_id] = self.get_task_info(task_id)
        return tasks_info
    
    def run_task_now(self, task_id: str) -> bool:
        """
        立即运行任务
        
        Args:
            task_id: 任务ID
        
        Returns:
            是否成功触发
        """
        try:
            if task_id in self.tasks:
                task_info = self.tasks[task_id]
                task_info['func']()
                self.logger.info(f"立即执行任务: {task_id}")
                return True
            else:
                self.logger.warning(f"任务不存在: {task_id}")
                return False
        except Exception as e:
            self.logger.error(f"立即执行任务失败: {str(e)}")
            return False
