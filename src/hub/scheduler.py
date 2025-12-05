"""
Task Scheduler for Automation Hub.
Handles scheduled and recurring task execution using APScheduler.
"""

import logging
import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from datetime import datetime

from PyQt5.QtCore import QObject, pyqtSignal

try:
    from apscheduler.schedulers.background import BackgroundScheduler
    from apscheduler.triggers.cron import CronTrigger
    from apscheduler.triggers.date import DateTrigger
    from apscheduler.triggers.interval import IntervalTrigger
    from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_ERROR
    APSCHEDULER_AVAILABLE = True
except ImportError:
    APSCHEDULER_AVAILABLE = False


class ScheduledTask:
    """Represents a scheduled task."""

    def __init__(
            self,
            id: str,
            name: str,
            script_id: str,
            parameters: Dict[str, Any],
            schedule_type: str,
            schedule_config: Dict[str, Any],
            enabled: bool = True
    ):
        self.id = id
        self.name = name
        self.script_id = script_id
        self.parameters = parameters
        self.schedule_type = schedule_type  # 'once', 'interval', 'cron'
        self.schedule_config = schedule_config
        self.enabled = enabled
        self.created_at = datetime.now()
        self.last_run: Optional[datetime] = None
        self.next_run: Optional[datetime] = None
        self.run_count = 0

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for persistence."""
        return {
            'id': self.id,
            'name': self.name,
            'script_id': self.script_id,
            'parameters': self.parameters,
            'schedule_type': self.schedule_type,
            'schedule_config': self.schedule_config,
            'enabled': self.enabled,
            'created_at': self.created_at.isoformat(),
            'last_run': self.last_run.isoformat() if self.last_run else None,
            'run_count': self.run_count
        }

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> 'ScheduledTask':
        """Create from dictionary."""
        task = ScheduledTask(
            id=data['id'],
            name=data['name'],
            script_id=data['script_id'],
            parameters=data['parameters'],
            schedule_type=data['schedule_type'],
            schedule_config=data['schedule_config'],
            enabled=data.get('enabled', True)
        )
        if data.get('created_at'):
            task.created_at = datetime.fromisoformat(data['created_at'])
        if data.get('last_run'):
            task.last_run = datetime.fromisoformat(data['last_run'])
        task.run_count = data.get('run_count', 0)
        return task


class SchedulerManager(QObject):
    """
    Manages scheduled and recurring tasks using APScheduler.
    """

    task_executed = pyqtSignal(str, bool)  # task_id, success
    task_added = pyqtSignal(ScheduledTask)
    task_removed = pyqtSignal(str)  # task_id
    task_updated = pyqtSignal(ScheduledTask)

    def __init__(self, script_manager: Any, config_manager: Any) -> None:
        super().__init__()
        self.logger = logging.getLogger(__name__)
        self.script_manager = script_manager
        self.config_manager = config_manager

        self.tasks: Dict[str, ScheduledTask] = {}
        self.scheduler: Optional[Any] = None

        if not APSCHEDULER_AVAILABLE:
            self.logger.warning("APScheduler not available - task scheduling disabled")
            return

        # Initialize scheduler
        self.scheduler = BackgroundScheduler()
        self.scheduler.add_listener(
            self._on_job_executed,
            EVENT_JOB_EXECUTED | EVENT_JOB_ERROR
        )

        # Load saved tasks
        self._load_tasks()

        # Start scheduler
        self.scheduler.start()
        self.logger.info("Scheduler started")

    def add_task(self, task: ScheduledTask) -> bool:
        """Add a scheduled task."""
        if not APSCHEDULER_AVAILABLE:
            self.logger.error("APScheduler not available")
            return False

        try:
            # Create APScheduler job
            trigger = self._create_trigger(task.schedule_type, task.schedule_config)

            if trigger is None:
                return False

            self.scheduler.add_job(
                func=self._execute_task,
                trigger=trigger,
                id=task.id,
                name=task.name,
                args=[task.id],
                replace_existing=True
            )

            # Store task
            self.tasks[task.id] = task

            # Update next run time
            job = self.scheduler.get_job(task.id)
            if job:
                task.next_run = job.next_run_time

            # Save and emit signal
            self._save_tasks()
            self.task_added.emit(task)

            self.logger.info(f"Task added: {task.name}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to add task: {e}")
            return False

    def remove_task(self, task_id: str) -> bool:
        """Remove a scheduled task."""
        if task_id not in self.tasks:
            return False

        try:
            # Remove from scheduler
            if self.scheduler:
                self.scheduler.remove_job(task_id)

            # Remove from tasks
            del self.tasks[task_id]

            # Save and emit signal
            self._save_tasks()
            self.task_removed.emit(task_id)

            self.logger.info(f"Task removed: {task_id}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to remove task: {e}")
            return False

    def enable_task(self, task_id: str, enabled: bool) -> bool:
        """Enable or disable a task."""
        if task_id not in self.tasks:
            return False

        task = self.tasks[task_id]
        task.enabled = enabled

        if enabled:
            # Resume job
            if self.scheduler:
                self.scheduler.resume_job(task_id)
        else:
            # Pause job
            if self.scheduler:
                self.scheduler.pause_job(task_id)

        self._save_tasks()
        self.task_updated.emit(task)
        return True

    def get_tasks(self) -> List[ScheduledTask]:
        """Get all scheduled tasks."""
        return list(self.tasks.values())

    def get_task(self, task_id: str) -> Optional[ScheduledTask]:
        """Get a specific task."""
        return self.tasks.get(task_id)

    def _execute_task(self, task_id: str) -> None:
        """Execute a scheduled task."""
        if task_id not in self.tasks:
            self.logger.error(f"Task not found: {task_id}")
            return

        task = self.tasks[task_id]

        if not task.enabled:
            return

        try:
            self.logger.info(f"Executing scheduled task: {task.name}")

            # Get script
            script = self.script_manager.get_script(task.script_id)
            if not script:
                raise ValueError(f"Script not found: {task.script_id}")

            # Execute script
            self.script_manager.execute_script(script, task.parameters)

            # Update task info
            task.last_run = datetime.now()
            task.run_count += 1
            self._save_tasks()

            self.task_executed.emit(task_id, True)

        except Exception as e:
            self.logger.error(f"Task execution failed: {e}")
            self.task_executed.emit(task_id, False)

    def _create_trigger(self, schedule_type: str, config: Dict[str, Any]) -> Optional[Any]:
        """Create APScheduler trigger from configuration."""
        try:
            if schedule_type == 'once':
                # One-time execution
                run_date = datetime.fromisoformat(config['run_date'])
                return DateTrigger(run_date=run_date)

            elif schedule_type == 'interval':
                # Interval-based (every N hours/days)
                return IntervalTrigger(**config)

            elif schedule_type == 'cron':
                # Cron-style schedule
                return CronTrigger(**config)

            else:
                self.logger.error(f"Unknown schedule type: {schedule_type}")
                return None

        except Exception as e:
            self.logger.error(f"Failed to create trigger: {e}")
            return None

    def _on_job_executed(self, event: Any) -> None:
        """Handle job execution event."""
        job_id = event.job_id
        if job_id in self.tasks:
            task = self.tasks[job_id]
            job = self.scheduler.get_job(job_id)
            if job:
                task.next_run = job.next_run_time

    def _load_tasks(self) -> None:
        """Load tasks from configuration."""
        try:
            tasks_file = Path(self.config_manager.config_dir) / 'scheduled_tasks.json'

            if tasks_file.exists():
                with open(tasks_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                for task_data in data:
                    task = ScheduledTask.from_dict(task_data)
                    self.add_task(task)

                self.logger.info(f"Loaded {len(self.tasks)} scheduled tasks")

        except Exception as e:
            self.logger.error(f"Failed to load tasks: {e}")

    def _save_tasks(self) -> None:
        """Save tasks to configuration."""
        try:
            tasks_file = Path(self.config_manager.config_dir) / 'scheduled_tasks.json'

            data = [task.to_dict() for task in self.tasks.values()]

            with open(tasks_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)

        except Exception as e:
            self.logger.error(f"Failed to save tasks: {e}")

    def shutdown(self) -> None:
        """Shutdown the scheduler."""
        if self.scheduler:
            self.scheduler.shutdown()
            self.logger.info("Scheduler stopped")
