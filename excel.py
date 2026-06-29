import win32com.client as win32
import yaml
import os
import time
import schedule
import requests
import base64
import hashlib
from datetime import datetime, timedelta
import logging
from logging.handlers import TimedRotatingFileHandler
import argparse
import pythoncom
import io
import threading
import shutil
import json

# ---------------------------- 颜色输出常量 ----------------------------
COLOR_GREEN = "\033[32m"
COLOR_RED = "\033[91m"
COLOR_GRAY = "\033[90m"
COLOR_BLUE = "\033[34m"
COLOR_YELLOW = "\033[33m"
COLOR_CYAN = "\033[36m"
COLOR_MAGENTA = "\033[35m"
COLOR_BOLD = "\033[1m"
COLOR_RESET = "\033[0m"

# ---------------------------- 执行记录管理 ----------------------------
def get_execution_log_path():
    """获取执行记录文件路径"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, "execution_log.json")

def load_execution_log():
    """加载执行记录"""
    log_path = get_execution_log_path()
    if os.path.exists(log_path):
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_execution_log(log_data):
    """保存执行记录"""
    log_path = get_execution_log_path()
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log_data, f, indent=2, ensure_ascii=False)

def record_execution(task_id: int, task_name: str, trigger_time: str, success: bool, manual=False):
    """记录任务执行结果"""
    log_data = load_execution_log()
    today = datetime.now().strftime("%Y-%m-%d")
    
    if today not in log_data:
        log_data[today] = {"tasks": {}}
    
    if str(task_id) not in log_data[today]["tasks"]:
        log_data[today]["tasks"][str(task_id)] = {"name": task_name, "executions": {}}
    
    status = "success" if success else "failed"
    log_data[today]["tasks"][str(task_id)]["executions"][trigger_time] = {
        "status": status,
        "manual": manual
    }
    
    save_execution_log(log_data)

# ---------------------------- 自动安装依赖 ----------------------------
def install_missing_dependencies():
    """自动安装未安装的依赖"""
    import subprocess
    import sys
    
    dependencies = {
        'PIL': 'pillow',
        'pyautogui': 'pyautogui',
        'pymsgbox': 'pymsgbox',
        'pyscreeze': 'pyscreeze',
        'pygetwindow': 'pygetwindow',
        'pyrect': 'pyrect',
        'mouseinfo': 'mouseinfo',
        'pytweening': 'pytweening',
        'portalocker': 'portalocker'
    }
    
    installed = []
    failed = []
    
    for module_name, package_name in dependencies.items():
        try:
            __import__(module_name)
            installed.append(module_name)
        except ImportError:
            print(f"检测到缺失依赖：{module_name}，正在安装...")
            try:
                subprocess.check_call([
                    sys.executable, '-m', 'pip', 'install', package_name,
                    '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple',
                    '--timeout', '60'
                ])
                try:
                    __import__(module_name)
                    print(f"成功安装：{module_name}")
                    installed.append(module_name)
                except ImportError:
                    print(f"安装成功但导入失败：{module_name}")
                    failed.append(module_name)
            except subprocess.CalledProcessError:
                print(f"安装失败：{module_name}")
                failed.append(module_name)
            except Exception as e:
                print(f"安装 {module_name} 时发生异常：{e}")
                failed.append(module_name)
    
    if installed:
        print(f"已安装/确认依赖：{', '.join(installed)}")
    if failed:
        print(f"安装失败的依赖：{', '.join(failed)}")
    
    return failed

install_missing_dependencies()

from PIL import Image
import portalocker

# ---------------------------- 日志配置 ----------------------------
class TaskIDFilter(logging.Filter):
    """日志过滤器：确保task_id始终存在"""
    def filter(self, record):
        if not hasattr(record, 'task_id'):
            record.task_id = '-'
        return True

class DateRotatingFileHandler(logging.FileHandler):
    """自定义日志处理器：按日期自动切换日志文件，支持年/月分级目录"""
    
    def __init__(self, base_dir, encoding='utf-8'):
        self.base_dir = base_dir
        self.current_date = datetime.now().date()
        log_path = self._get_log_path()
        super().__init__(log_path, mode='a', encoding=encoding)
    
    def _get_log_path(self):
        now = datetime.now()
        log_dir = os.path.join(self.base_dir, "logs", str(now.year), f"{now.month:02d}")
        os.makedirs(log_dir, exist_ok=True)
        log_filename = now.strftime("%Y-%m-%d.log")
        return os.path.join(log_dir, log_filename)
    
    def emit(self, record):
        today = datetime.now().date()
        if today != self.current_date:
            self.close()
            self.current_date = today
            self.baseFilename = self._get_log_path()
            self.stream = self._open()
        super().emit(record)

def setup_logging():
    """配置日志系统：按年/月分级目录，每日独立日志文件（支持跨天自动切换）"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    log_format = "%(asctime)s [%(levelname)s] [任务%(task_id)s] %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    
    formatter = logging.Formatter(log_format, datefmt=date_format)
    
    handlers = [
        logging.StreamHandler(),
        DateRotatingFileHandler(base_dir, encoding='utf-8')
    ]
    
    for handler in handlers:
        handler.setFormatter(formatter)
        handler.addFilter(TaskIDFilter())
    
    logging.basicConfig(
        level=logging.DEBUG,
        handlers=handlers,
        force=True
    )
    
    logging.getLogger("PIL").setLevel(logging.WARNING)
    logging.getLogger("PIL.PngImagePlugin").setLevel(logging.WARNING)
    
    logger = logging.getLogger("ExcelBot")
    logger = logging.LoggerAdapter(logger, {"task_id": "-"})
    
    current_log_path = handlers[1].baseFilename
    logger.info(f"日志系统已初始化，日志文件：{current_log_path}")
    return logger

logger = setup_logging()

def get_task_logger(task_id: int):
    """获取带任务序号的日志记录器"""
    base_logger = logging.getLogger("ExcelBot")
    return logging.LoggerAdapter(base_logger, {"task_id": task_id})

def is_in_monthly_range(start_day: int, end_day: int, config_name: str = "") -> bool:
    """
    检查今天是否在指定的每月日期范围内
    
    :param start_day: 开始日期（1-31，负数表示倒数第几天）
    :param end_day: 结束日期（1-31，负数表示倒数第几天，0表示最后一天）
    :param config_name: 配置名称（用于日志输出）
    :return: True表示在范围内，False表示不在范围内
    """
    today = datetime.now().day
    current_month = datetime.now().month
    current_year = datetime.now().year
    
    if current_month == 12:
        next_month = datetime(current_year + 1, 1, 1)
    else:
        next_month = datetime(current_year, current_month + 1, 1)
    days_in_month = (next_month - datetime(current_year, current_month, 1)).days
    
    original_start_day = start_day
    original_end_day = end_day
    
    if start_day <= 0:
        start_day = days_in_month + start_day + 1
    if end_day <= 0:
        end_day = days_in_month + end_day
    
    adjusted = False
    if start_day > days_in_month:
        start_day = days_in_month
        adjusted = True
    elif start_day < 1:
        start_day = 1
        adjusted = True
    
    if end_day > days_in_month:
        end_day = days_in_month
        adjusted = True
    elif end_day < 1:
        end_day = 1
        adjusted = True
    
    if adjusted and config_name:
        logger.warning(
            f"[{config_name}] 日期范围配置被调整："
            f"原始({original_start_day}-{original_end_day}) -> "
            f"调整后({start_day}-{end_day})，当月天数：{days_in_month}"
        )
    
    if start_day <= end_day:
        return start_day <= today <= end_day
    else:
        return today >= start_day or today <= end_day

# ---------------------------- 任务日志分隔工具 ----------------------------
def _get_display_length(s):
    """计算字符串的显示长度（去除ANSI颜色转义码）"""
    import re
    return len(re.sub(r'\x1b\[[0-9;]*m', '', s))

def print_task_header(task_id, task_name, target_logger=None):
    """打印任务开始分隔框"""
    log = target_logger or logger
    box_width = 80
    content_width = box_width - 4
    
    title = f"[任务{task_id}] {task_name} - 开始执行"
    display_len = _get_display_length(title)
    if display_len > content_width:
        title = title[:content_width]
        display_len = content_width
    padding = " " * ((content_width - display_len) // 2)
    title_line = f"{padding}{title}{padding}"
    if _get_display_length(title_line) < content_width:
        title_line += " " * (content_width - _get_display_length(title_line))
    
    log.info(f"{COLOR_CYAN}{COLOR_BOLD}╔{'═' * content_width}╗{COLOR_RESET}")
    log.info(f"{COLOR_CYAN}{COLOR_BOLD}║{COLOR_RESET} {title_line} {COLOR_CYAN}{COLOR_BOLD}║{COLOR_RESET}")
    log.info(f"{COLOR_CYAN}{COLOR_BOLD}╚{'═' * content_width}╝{COLOR_RESET}")

def print_task_footer(task_id, task_name, success, elapsed_time, target_logger=None):
    """打印任务结束分隔框"""
    log = target_logger or logger
    box_width = 80
    content_width = box_width - 4
    
    status_icon = f"{COLOR_GREEN}✓{COLOR_RESET}" if success else f"{COLOR_RED}✗{COLOR_RESET}"
    status_text = "完成" if success else "失败"
    time_text = f"耗时: {elapsed_time:.2f}s"
    
    title = f"[任务{task_id}] {task_name} - {status_text} {status_icon} | {time_text}"
    display_len = _get_display_length(title)
    if display_len > content_width:
        title = title[:content_width]
        display_len = content_width
    padding = " " * ((content_width - display_len) // 2)
    title_line = f"{padding}{title}{padding}"
    if _get_display_length(title_line) < content_width:
        title_line += " " * (content_width - _get_display_length(title_line))
    
    border_color = COLOR_GREEN if success else COLOR_RED
    log.info(f"{border_color}{COLOR_BOLD}╔{'═' * content_width}╗{COLOR_RESET}")
    log.info(f"{border_color}{COLOR_BOLD}║{COLOR_RESET} {title_line} {border_color}{COLOR_BOLD}║{COLOR_RESET}")
    log.info(f"{border_color}{COLOR_BOLD}╚{'═' * content_width}╝{COLOR_RESET}")

def print_subtask_header(task_id, subtask_idx, subtask_name="", target_logger=None):
    """打印子任务开始分隔线"""
    log = target_logger or logger
    name_part = f" {subtask_name}" if subtask_name else ""
    line = f"  ── [子任务{task_id}-{subtask_idx}]{name_part} ──"
    log.info(f"{COLOR_YELLOW}{line}{COLOR_RESET}")

def print_subtask_footer(task_id, subtask_idx, success, target_logger=None):
    """打印子任务结束分隔线"""
    log = target_logger or logger
    status_icon = f"{COLOR_GREEN}✓{COLOR_RESET}" if success else f"{COLOR_RED}✗{COLOR_RESET}"
    status_text = "完成" if success else "失败"
    line = f"  ── [子任务{task_id}-{subtask_idx}] {status_text} {status_icon} ──"
    log.info(f"{COLOR_YELLOW}{line}{COLOR_RESET}")

# ---------------------------- 文件锁工具类 ----------------------------
class FileLock:
    """文件锁工具类，用于防止并发访问"""
    
    def __init__(self, lock_file_path, lock_name=""):
        self.lock_file_path = lock_file_path
        self.lock_file = None
        self.lock_name = lock_name or os.path.basename(lock_file_path)
        self._first_wait_logged = False
    
    def acquire(self, timeout=300, poll_interval=2):
        """获取文件锁
        
        :param timeout: 最大等待时间（秒），默认300秒
        :param poll_interval: 轮询间隔（秒），默认2秒
        :return: True表示获取锁成功，False表示超时
        """
        start_time = time.time()
        self._first_wait_logged = False
        
        while time.time() - start_time < timeout:
            try:
                self.lock_file = open(self.lock_file_path, 'w')
                portalocker.lock(self.lock_file, portalocker.LOCK_EX | portalocker.LOCK_NB)
                waited_ms = int((time.time() - start_time) * 1000)
                if waited_ms > 0:
                    logger.info(f"[{self.lock_name}] 成功获取文件锁，等待耗时 {waited_ms}ms")
                else:
                    logger.info(f"[{self.lock_name}] 成功获取文件锁")
                return True
            except (IOError, BlockingIOError, portalocker.LockException):
                if self.lock_file:
                    try:
                        self.lock_file.close()
                    except Exception:
                        pass
                    self.lock_file = None
                
                elapsed_ms = int((time.time() - start_time) * 1000)
                if not self._first_wait_logged:
                    self._first_wait_logged = True
                    remaining = int(timeout - (time.time() - start_time))
                    logger.warning(f"[{self.lock_name}] 锁被占用，开始等待（超时时间 {timeout}秒）")
                
                time.sleep(poll_interval)
        
        waited_ms = int((time.time() - start_time) * 1000)
        logger.error(f"[{self.lock_name}] 获取文件锁超时，等待耗时 {waited_ms}ms")
        return False
    
    def release(self):
        """释放文件锁并删除lock文件"""
        if self.lock_file:
            try:
                portalocker.unlock(self.lock_file)
                self.lock_file.close()
                logger.info(f"成功释放文件锁：{self.lock_file_path}")
            except Exception as e:
                logger.warning(f"释放文件锁异常：{str(e)}")
            finally:
                self.lock_file = None
        
        # 删除lock文件，避免文件累积
        if self.lock_file_path and os.path.exists(self.lock_file_path):
            try:
                os.remove(self.lock_file_path)
                logger.debug(f"已删除lock文件：{self.lock_file_path}")
            except Exception as e:
                logger.warning(f"删除lock文件失败：{str(e)}")
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.release()

try:
    import pyautogui
    logger.info("成功导入 pyautogui")
except Exception as e:
    logger.warning(f"导入 pyautogui 失败：{e}")
    pyautogui = None

# ---------------------------- Excel 处理器 ----------------------------
class ExcelProcessor:
    """Excel 操作引擎"""
    
    def __init__(self, file_path: str, visible=True, task_logger=None):
        self.file_path = os.path.abspath(file_path)
        self.visible = True
        self.excel = None
        self.workbook = None
        self._refresh_timeout = 500
        self._dialog_watchdog_stop = threading.Event()
        self._file_lock = None
        self.logger = task_logger or logger

    def __enter__(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        locks_dir = os.path.join(base_dir, "locks")
        os.makedirs(locks_dir, exist_ok=True)
        
        file_name = os.path.basename(self.file_path)
        lock_file_path = os.path.join(locks_dir, file_name + ".lock")
        self._file_lock = FileLock(lock_file_path)
        
        if not self._file_lock.acquire(timeout=300):
            raise RuntimeError(f"无法获取文件锁，超时退出：{self.file_path}")
        
        max_retries = 3
        retry_delay = 2
        
        pythoncom.CoInitialize()
        
        for attempt in range(max_retries + 1):
            try:
                self.excel = win32.DispatchEx("Excel.Application")
                self.logger.debug("成功创建 Excel 实例")
                try:
                    self.excel.Visible = self.visible
                    self.logger.debug(f"成功设置 Excel 可见性: {self.visible}")
                except Exception as e:
                    self.logger.warning(f"设置 Excel 可见性失败: {e}")
                try:
                    self.excel.DisplayAlerts = False
                    self.logger.debug("成功设置 DisplayAlerts = False")
                except Exception as e:
                    self.logger.warning(f"设置 DisplayAlerts 失败: {e}")
                self.workbook = self.excel.Workbooks.Open(self.file_path)

                self.logger.debug(f"等待5秒")
                time.sleep(5)
                
                if pyautogui:
                    self.logger.debug(f"按下Esc键")
                    pyautogui.press('esc')

                self.logger.debug(f"等待3秒")
                time.sleep(3)
                
                for sheet in self._iter_worksheets():
                    try:
                        sheet.Activate()
                        sheet.Application.ActiveWindow.Zoom = 220
                    except Exception as e:
                        self.logger.debug(f"设置缩放失败：{str(e)}")
                self.logger.debug(f"成功打开文件：{os.path.basename(self.file_path)}")
                return self
            except Exception as e:
                error_str = str(e)
                if "消息筛选器显示应用程序正在使用中" in error_str and attempt < max_retries:
                    self.logger.warning(f"Excel 忙，第 {attempt + 1}/{max_retries} 次重试...")
                    self._safe_shutdown()
                    time.sleep(retry_delay)
                    retry_delay *= 2
                else:
                    self._safe_shutdown()
                    raise RuntimeError(f"Excel 启动失败：{str(e)}")

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._safe_shutdown()

    def _iter_worksheets(self):
        if not self.workbook:
            return []
        sheets = self.workbook.Worksheets
        return [sheets.Item(i) for i in range(1, sheets.Count + 1)]

    def _start_dialog_watchdog(self, timeout_s: float = 90.0):
        if pyautogui is None:
            return

        self._dialog_watchdog_stop.clear()
        self.logger.info(f"启动弹窗守护线程，超时时间：{timeout_s}秒")

        def _run():
            end_at = time.time() + float(timeout_s)
            while not self._dialog_watchdog_stop.is_set() and time.time() < end_at:
                self._dismiss_other_people_editing_dialog()
                time.sleep(0.1)
            self.logger.info("弹窗守护线程结束")

        threading.Thread(target=_run, daemon=True).start()

    def _stop_dialog_watchdog(self):
        try:
            self._dialog_watchdog_stop.set()
        except Exception:
            pass

    def _dismiss_other_people_editing_dialog(self) -> bool:
        if pyautogui is None:
            return False

        button_images_dir = os.path.join(os.path.dirname(__file__), 'button_images')
        os.makedirs(button_images_dir, exist_ok=True)
        
        # 自动扫描目录下所有图片文件
        button_images = []
        for filename in os.listdir(button_images_dir):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                button_images.append(filename)
        
        for button_image in button_images:
            button_path = os.path.join(button_images_dir, button_image)
            try:
                location = pyautogui.locateCenterOnScreen(button_path)
                if location:
                    pyautogui.click(location)
                    self.logger.info(f"检测到弹窗，已点击按钮：{button_image}")
                    return True
            except Exception as e:
                self.logger.debug(f"未检测到按钮 {button_image} ：{e}")
        
        return False

    def _safe_shutdown(self):
        self._stop_dialog_watchdog()

        if self.workbook is not None:
            try:
                self.workbook.Worksheets(1).Activate()
            except Exception as e:
                self.logger.debug(f"激活第一个工作表失败：{str(e)}")
            
            try:
                self.workbook.Close(SaveChanges=True)
            except Exception as e:
                self.logger.warning(f"关闭工作簿异常：{str(e)}")
            finally:
                self.workbook = None

        if self.excel is not None:
            try:
                self.excel.Quit()
            except Exception as e:
                self.logger.warning(f"关闭 Excel 进程异常：{str(e)}")
            finally:
                self.excel = None

        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            self.logger.debug(f"COM 反初始化异常：{str(e)}")

        if self._file_lock:
            self._file_lock.release()
            self._file_lock = None

        self.logger.debug("Excel 进程已释放")

    def force_terminate(self):
        """强制终止Excel进程（用于任务超时场景）"""
        self.logger.warning("强制终止Excel进程开始")
        
        self._stop_dialog_watchdog()

        # 先尝试安全关闭
        if self.workbook is not None:
            try:
                self.workbook.Close(SaveChanges=True)
                self.logger.info("成功关闭工作簿")
            except Exception as e:
                self.logger.warning(f"安全关闭工作簿失败：{str(e)}")
            finally:
                self.workbook = None

        # 尝试正常退出Excel
        if self.excel is not None:
            try:
                self.excel.Quit()
                self.logger.info("成功退出Excel进程")
            except Exception as e:
                self.logger.warning(f"正常退出Excel失败：{str(e)}")
            finally:
                self.excel = None

        # 强制终止所有Excel进程（作为最后的手段）
        try:
            import subprocess
            subprocess.run(
                ["taskkill", "/f", "/im", "EXCEL.EXE"],
                capture_output=True,
                timeout=10
            )
            self.logger.info("已强制终止所有Excel进程")
        except Exception as e:
            self.logger.warning(f"强制终止Excel进程异常：{str(e)}")

        # 反初始化COM
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            self.logger.debug(f"COM 反初始化异常：{str(e)}")

        # 强制释放文件锁
        if self._file_lock:
            try:
                # 先尝试正常释放
                self._file_lock.release()
            except Exception as e:
                self.logger.warning(f"正常释放锁失败，尝试强制删除锁文件：{str(e)}")
                # 强制删除锁文件
                if self._file_lock.lock_file_path and os.path.exists(self._file_lock.lock_file_path):
                    try:
                        os.remove(self._file_lock.lock_file_path)
                        self.logger.info("已强制删除锁文件")
                    except Exception as ex:
                        self.logger.error(f"强制删除锁文件失败：{str(ex)}")
            finally:
                self._file_lock = None

        self.logger.warning("Excel进程强制终止完成")

    def refresh_data(self) -> bool:
        self.logger.info("开始刷新数据...")
        start_time = time.time()
        
        try:
            linked_tables = []
            for sheet in self._iter_worksheets():
                try:
                    list_objects = sheet.ListObjects
                    for i in range(1, list_objects.Count + 1):
                        table = list_objects.Item(i)
                        try:
                            if hasattr(table, 'QueryTable') and table.QueryTable is not None:
                                table_range = table.Range.Address if hasattr(table, 'Range') else "未知"
                                linked_tables.append({
                                    'sheet': sheet.Name,
                                    'table': table.Name,
                                    'range': table_range,
                                    'table_obj': table,
                                    'query_table': table.QueryTable
                                })
                                self.logger.debug(f"发现链接数据源的表格：工作表 [{sheet.Name}] - 查询 [{table.Name}] - 范围 [{table_range}]")
                        except Exception as e:
                            self.logger.debug(f"工作表 [{sheet.Name}] 的表格 [{table.Name if hasattr(table, 'Name') else 'Unknown'}] 未连接数据源")
                except Exception as e:
                    self.logger.debug(f"检查工作表 [{sheet.Name}] 时出错：{e}")
            
            if linked_tables:
                self.logger.info(f"共发现 {len(linked_tables)} 个链接了数据源的表格")
                for item in linked_tables:
                    self.logger.info(f"  - 工作表 [{item['sheet']}] - 查询 [{item['table']}] - 范围 [{item['range']}]")
            else:
                self.logger.info("未发现链接了数据源的表格")

            refresh_tables = []
            if linked_tables:
                self.logger.info("检查数据连接属性：")
                for item in linked_tables:
                    try:
                        query_table = item['query_table']
                        
                        if query_table is not None:
                            will_refresh_on_refresh_all = False
                            try:
                                workbook_conn = query_table.WorkbookConnection
                                if workbook_conn:
                                    try:
                                        if hasattr(workbook_conn, 'RefreshWithRefreshAll'):
                                            will_refresh_on_refresh_all = workbook_conn.RefreshWithRefreshAll
                                    except:
                                        pass
                            except Exception as conn_e:
                                self.logger.debug(f"检查连接对象时出错: {conn_e}")
                            
                            status = "✓ 已设置" if will_refresh_on_refresh_all else "✗ 未设置"
                            self.logger.info(f"  工作表 [{item['sheet']}] - 查询 [{item['table']}]:")
                            self.logger.info(f"    - 全部刷新时刷新此连接: {status}")
                            
                            if will_refresh_on_refresh_all:
                                self.logger.info(f"    - 数据源表格范围: {item['range']}")
                                refresh_tables.append(item)
                                
                    except Exception as e:
                        self.logger.warning(f"检查表格 [{item['sheet']}] - [{item['table']}] 的连接属性时出错: {e}")
            
            if refresh_tables:
                self.logger.info(f"在 {len(refresh_tables)} 个工作表中设置左上角单元格值为1")
                for item in refresh_tables:
                    range_start = item['range'].split(':')[0]
                    try:
                        sheet = self.workbook.Worksheets(item['sheet'])
                        sheet.Range(range_start).Value = 1
                        self.logger.info(f"  工作表 [{item['sheet']}] - 已将 {range_start} 单元格值设置为 1")
                    except Exception as e:
                        self.logger.warning(f"设置工作表 [{item['sheet']}] 的 {range_start} 单元格值时出错: {e}")

            time.sleep(10)
            
            max_retries = 3
            failed_tables = []
            
            for retry_count in range(max_retries + 1):
                if retry_count == 0:
                    self.logger.info("执行全部刷新...")
                    self.workbook.RefreshAll()
                    self.excel.CalculateUntilAsyncQueriesDone()
                else:
                    if not failed_tables:
                        break
                    
                    self.logger.warning(f"第 {retry_count} 次重试刷新，发现 {len(failed_tables)} 个表格刷新失败")
                    for item in failed_tables:
                        self.logger.warning(f"  重试刷新：工作表 [{item['sheet']}] - 查询 [{item['table']}]")
                        try:
                            if item['query_table']:
                                item['query_table'].Refresh()
                        except Exception as e:
                            self.logger.error(f"重试刷新表格 [{item['sheet']}] - [{item['table']}] 时出错: {e}")
                    
                    self.excel.CalculateUntilAsyncQueriesDone()
                
                calculation_timeout = 300
                calculation_start = time.time()
                while time.time() - calculation_start < calculation_timeout:
                    if self.excel.CalculationState == 0:
                        break
                    time.sleep(5)
                else:
                    self.logger.warning("计算状态检查超时，继续验证单元格值")
                
                if refresh_tables:
                    failed_tables = []
                    for item in refresh_tables:
                        range_start = item['range'].split(':')[0]
                        try:
                            sheet = self.workbook.Worksheets(item['sheet'])
                            cell_value = sheet.Range(range_start).Value
                            if str(cell_value).strip() != '1':
                                self.logger.info(f"工作表 [{item['sheet']}] - 查询 [{item['table']}] 的 {range_start} 单元格值已更新，刷新成功")
                            else:
                                failed_tables.append(item)
                                self.logger.warning(f"工作表 [{item['sheet']}] - 查询 [{item['table']}] 的 {range_start} 单元格值仍为1，刷新失败")
                        except Exception as e:
                            self.logger.warning(f"检查工作表 [{item['sheet']}] 的 {range_start} 单元格值时出错: {e}")
                            failed_tables.append(item)
                    
                    if not failed_tables:
                        self.logger.info("所有表格刷新成功！")
                        self._start_dialog_watchdog(timeout_s=90)
                        for sheet in self._iter_worksheets():
                            try:
                                if sheet.AutoFilter is not None:
                                    sheet.AutoFilter.ApplyFilter()
                                    self.logger.info(f"重新应用筛选：{sheet.Name}")
                            except Exception as e:
                                self.logger.warning(f"应用筛选/排序失败：{sheet.Name} - {e}")
                        self._stop_dialog_watchdog()
                        for sheet in self._iter_worksheets():
                            try:
                                if hasattr(sheet, "PivotTables"):
                                    for i in range(1, sheet.PivotTables().Count + 1):
                                        pt = sheet.PivotTables(i)
                                        pt.RefreshTable()
                                        self.logger.info(f"刷新数据透视表：{sheet.Name} - {pt.Name}")
                                        time.sleep(1)
                            except Exception as e:
                                self.logger.warning(f"刷新数据透视表失败：{sheet.Name} - {e}")

                        return True
                    else:
                        if retry_count < max_retries:
                            self.logger.warning(f"发现 {len(failed_tables)} 个表格刷新失败，准备重试（剩余 {max_retries - retry_count} 次）")
                            time.sleep(5)
                        else:
                            self.logger.error(f"达到最大重试次数（{max_retries}次），仍有 {len(failed_tables)} 个表格刷新失败")
                            for item in failed_tables:
                                self.logger.error(f"  失败表格：工作表 [{item['sheet']}] - 查询 [{item['table']}]")
                            return False
                else:
                    self.logger.info("没有需要验证的表格，刷新完成")
                    self._start_dialog_watchdog(timeout_s=90)
                    for sheet in self._iter_worksheets():
                        try:
                            if sheet.AutoFilter is not None:
                                sheet.AutoFilter.ApplyFilter()
                                self.logger.info(f"重新应用筛选：{sheet.Name}")
                        except Exception as e:
                            self.logger.warning(f"应用筛选/排序失败：{sheet.Name} - {e}")
                    self._stop_dialog_watchdog()
                    for sheet in self._iter_worksheets():
                        try:
                            if hasattr(sheet, "PivotTables"):
                                for i in range(1, sheet.PivotTables().Count + 1):
                                    pt = sheet.PivotTables(i)
                                    pt.RefreshTable()
                                    self.logger.info(f"刷新数据透视表：{sheet.Name} - {pt.Name}")
                                    time.sleep(1)
                        except Exception as e:
                            self.logger.warning(f"刷新数据透视表失败：{sheet.Name} - {e}")
                    return True
            
            self.logger.warning("刷新循环异常结束")
            return False
        except Exception as e:
            self.logger.error(f"刷新异常：{str(e)}")
            return False
        finally:
            self._stop_dialog_watchdog()

    def validate_date(self, check_sheet, check_range, check_frequency) -> bool:
        for attempt in range(1, check_frequency+1):
            try:
                self.logger.debug(f"校验数据：工作表 [{check_sheet}] 区域 [{check_range}]")
                sheet = self.workbook.Worksheets(check_sheet)
                valid = sheet.Range(check_range).Value != 0
                self.logger.info(f"数据校验 {'通过' if valid else '失败'}（第 {attempt} 次尝试）共{check_frequency}次")
    
                if valid:
                    return True
                if attempt < check_frequency:
                    time.sleep(10)
                    self.refresh_data()
            except Exception as e:
                self.logger.error(f"校验异常：{str(e)}")
        return False

    def capture_screenshots(self, configs: list, retry_times: int = 3):
        screenshots_dict = {}
        pending_configs = list(configs)
        total_attempts = retry_times + 1

        try:
            for attempt in range(1, total_attempts + 1):
                if not pending_configs:
                    break

                if attempt == 1:
                    self.logger.info(f"开始截图，共 {len(pending_configs)} 个区域")
                else:
                    self.logger.warning(
                        f"截图重试第 {attempt - 1}/{retry_times} 次，待重试区域 {len(pending_configs)} 个"
                    )

                next_pending = []
                for cfg in pending_configs:
                    start_day = cfg.get("start_day", 1)
                    end_day = cfg.get("end_day", 31)
                    config_name = cfg.get("name", "")
                    
                    if not is_in_monthly_range(start_day, end_day, config_name):
                        self.logger.info(
                            f"跳过截图 [{config_name}]：当前日期不在播报范围内（{start_day}-{end_day}）"
                        )
                        continue
                    
                    try:
                        sheet = self.workbook.Worksheets(cfg["sheet_name"])
                        output_path = self._generate_path(cfg["name"])

                        if self._capture_range(sheet, cfg["range"], output_path):
                            cfg_idx = configs.index(cfg)
                            screenshots_dict[cfg_idx] = output_path
                            self.logger.info(
                                f"截图成功：[{cfg['name']}] 工作表[{cfg['sheet_name']}] 区域[{cfg['range']}]"
                            )
                        else:
                            next_pending.append(cfg)
                            self.logger.warning(
                                f"截图失败：[{cfg['name']}] 工作表[{cfg['sheet_name']}] 区域[{cfg['range']}]"
                            )
                    except Exception as e:
                        next_pending.append(cfg)
                        self.logger.error(f"截图异常 [{cfg['name']}]：{str(e)}")

                pending_configs = next_pending
                if pending_configs and attempt < total_attempts:
                    time.sleep(2)
        finally:
            try:
                for sheet in self._iter_worksheets():
                    sheet.Activate()
                    sheet.Application.ActiveWindow.Zoom = 100
                self.logger.debug("已将所有工作表缩放比例恢复为100%")
            except Exception as e:
                self.logger.warning(f"恢复缩放比例失败：{str(e)}")

        screenshots = [screenshots_dict[i] for i in range(len(configs)) if i in screenshots_dict]
        return screenshots, pending_configs
    

    def _capture_range(self, sheet, range_addr: str, output_path: str) -> bool:
        try:
            if ":" in range_addr:
                range_obj = sheet.Range(range_addr)
            else:
                start_cell = sheet.Range(range_addr.split(":")[0])
                range_obj = start_cell.CurrentRegion

            self.logger.debug(f"截图区域地址: {range_obj.Address}")

            range_obj.CopyPicture(Format=1)
            time.sleep(1)

            left = range_obj.Left
            top = range_obj.Top
            width = range_obj.Width
            height = range_obj.Height

            chart_obj = sheet.ChartObjects().Add(left, top, width, height)
            chart = chart_obj.Chart
            chart_obj.Activate()
            try:
                chart.Paste()
            except Exception as e:
                self.logger.error(f"Paste异常：{str(e)}", exc_info=True)
                chart_obj.Delete()
                return False
            chart.Export(output_path)
            chart_obj.Delete()
            return os.path.exists(output_path)
        except Exception as e:
            self.logger.error(f"截图异常：{str(e)}", exc_info=True)
            return False

    def _generate_path(self, prefix: str) -> str:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        task_tag = os.path.splitext(os.path.basename(self.file_path))[0]
        screenshots_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshots_temp")
        os.makedirs(screenshots_dir, exist_ok=True)
        return os.path.join(
            screenshots_dir,
            f"{task_tag}_{prefix}_{timestamp}.png"
        )

# ---------------------------- 任务处理器 ----------------------------
class ReportTask:
    """报表任务实例"""

    def __init__(self, config: dict, test_webhook: str = None, error_webhook: str = None, upload_url_template: str = None, task_id: int = 0, retry_config: dict = None):
        self.config = self._validate_config(config)
        self.retry_limit = 3
        self.error_webhook = error_webhook or "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=833b098e-d8b8-43ea-bfdf-cade0d040fb6"
        self.test_webhook = test_webhook
        self.upload_url_template = upload_url_template
        self.task_id = task_id
        self.logger = get_task_logger(task_id)
        
        # 重试配置
        self.retry_config = retry_config or {}
        self.retry_enabled = self.retry_config.get("enabled", False)
        self.retry_delay_minutes = self.retry_config.get("delay_minutes", 10)
        self.retry_max_attempts = self.retry_config.get("max_attempts", 3)
        self.retry_count = 0
        # Webhook级别重试记录 {wh_idx: retry_count}
        self.webhook_retry_counts = {}
        
        # 任务超时配置（默认15分钟）
        self.task_timeout_minutes = self.retry_config.get("task_timeout_minutes", 15)
        # 当前执行的excel实例（用于超时终止）
        self.current_excel = None

    def _validate_config(self, config: dict) -> dict:
        """配置完整性检查，支持新的多webhook配置和旧的单webhook配置"""
        required_fields = ["excel_path"]
        missing = [f for f in required_fields if f not in config]
        if missing:
            raise ValueError(f"缺失必要配置：{missing}")

        if not os.path.exists(config["excel_path"]):
            raise FileNotFoundError(config["excel_path"])

        # 支持新的webhooks配置格式
        if "webhooks" in config:
            if not isinstance(config["webhooks"], list) or len(config["webhooks"]) == 0:
                raise ValueError("webhooks必须是非空列表")
            
            for idx, webhook_config in enumerate(config["webhooks"]):
                if "webhook" not in webhook_config:
                    raise ValueError(f"webhook配置[{idx}]缺失webhook字段")
                if "times" not in webhook_config:
                    raise ValueError(f"webhook配置[{idx}]缺失times字段")
                if "capture_configs" not in webhook_config:
                    raise ValueError(f"webhook配置[{idx}]缺失capture_configs字段")
                
                # 设置默认值
                if "send_file_enable" not in webhook_config:
                    webhook_config["send_file_enable"] = 0
        elif "schedule" in config:
            # 兼容旧格式，转换为新格式
            schedule = config["schedule"]
            webhook = schedule.get("webhook", "")
            times = schedule.get("times", [])
            capture_configs = config.get("capture_configs", [])
            send_file_enable = config.get("send_file_enable", 0)
            
            config["webhooks"] = [{
                "webhook": webhook,
                "times": times,
                "capture_configs": capture_configs,
                "send_file_enable": send_file_enable
            }]
            self.logger.warning("检测到旧版配置格式，已自动转换为多webhook格式")
        else:
            raise ValueError("任务配置必须包含webhooks或schedule字段")

        return config

    def _get_upload_url(self, webhook: str) -> str:
        """根据webhook获取文件上传URL"""
        if not self.upload_url_template:
            return ""
        
        try:
            key = webhook.split("key=")[-1]
            return self.upload_url_template.format(key=key)
        except Exception as e:
            self.logger.warning(f"构建上传URL失败：{str(e)}")
            return ""

    def _get_task_id_str(self, wh_idx: int = None) -> str:
        """获取任务序号字符串（用于消息内容）"""
        if wh_idx is not None and wh_idx >= 0:
            return f"[任务{self.task_id}-{wh_idx}]"
        return f"[任务{self.task_id}]"

    def execute(self, debug_mode=False, webhook_configs=None, is_manual=False):
        """
        执行任务流程
        :param debug_mode: 是否调试模式
        :param webhook_configs: 特定的webhook配置（单个dict或列表，None表示执行所有webhook）
        :param is_manual: 是否手动执行（用于决定是否立即发送失败通知）
        :return: True表示成功，False表示失败
        """
        task_name = os.path.basename(self.config['excel_path'])
        print_task_header(self.task_id, task_name, target_logger=self.logger)
        
        start_time = time.time()
        results_to_deliver = []
        failed_webhooks = []
        success = True
        
        try:
            with ExcelProcessor(
                self.config["excel_path"], 
                visible=debug_mode,
                task_logger=self.logger
            ) as excel:
                # 保存当前excel实例供超时终止使用
                self.current_excel = excel
                # 刷新数据（所有webhook共享一次刷新）
                if not excel.refresh_data():
                    task_id_str = self._get_task_id_str()
                    self.logger.warning(f"{task_id_str}数据刷新失败")
                    
                    # 判断是否需要立即发送通知
                    # 条件：(重试未启用) 或 (手动执行)
                    should_notify = not self.retry_enabled or is_manual
                    
                    if should_notify:
                        self._send_wechat(
                            type="text",
                            data={
                                "content": f"{task_id_str}数据刷新失败（超时或重试3次后仍有表格未刷新成功），请检查文件：{os.path.basename(self.config['excel_path'])}",
                                "mentioned_list": ["zhufuzhe"]
                            },
                            description="数据刷新失败通知",
                            webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=833b098e-d8b8-43ea-bfdf-cade0d040fb6"
                        )
                    
                    success = False
                    return success, []

                # 数据校验（任务级别）
                if self.config.get("data_check_enable", False):
                    check_sheet = self.config["data_check"]["check_sheet"]
                    check_range = self.config["data_check"]["check_range"]
                    check_frequency = self.config["data_check"]["check_frequency"]
                    if not excel.validate_date(check_sheet, check_range, check_frequency):
                        task_id_str = self._get_task_id_str()
                        self.logger.warning(f"{task_id_str}数据校验未通过，发送通知并终止任务")
                        self._send_wechat(
                            type="text",
                            data={"content": f"{task_id_str}{self.config['data_check']['notify_message']}", 
                                "mentioned_list": self.config["data_check"]["notify_users"]
                            },
                            description="数据校验失败通知",
                            webhook = self.config["data_check"]["warning_webhook"]
                        )
                        success = False
                        return success, []
                
                # 确定要执行的webhook配置
                target_webhooks = []
                target_webhooks_with_indices = []
                
                if webhook_configs:
                    if isinstance(webhook_configs, list):
                        target_webhooks = webhook_configs
                        # 找到每个配置在原始webhooks中的索引
                        for wh_config in webhook_configs:
                            wh_idx = -1
                            # 尝试通过比较内容来查找索引
                            for i, config in enumerate(self.config["webhooks"]):
                                if config == wh_config:
                                    wh_idx = i
                                    break
                            target_webhooks_with_indices.append((wh_config, wh_idx))
                    else:
                        target_webhooks = [webhook_configs]
                        wh_idx = -1
                        # 尝试通过比较内容来查找索引
                        for i, config in enumerate(self.config["webhooks"]):
                            if config == webhook_configs:
                                wh_idx = i
                                break
                        target_webhooks_with_indices.append((webhook_configs, wh_idx))
                else:
                    target_webhooks = self.config["webhooks"]
                    for wh_idx, wh_config in enumerate(self.config["webhooks"]):
                        target_webhooks_with_indices.append((wh_config, wh_idx))
                
                # 为每个webhook执行截图
                failed_webhooks = []  # 记录失败的webhook
                for wh_config, wh_idx in target_webhooks_with_indices:
                    # 获取带主任务-子任务序号的logger
                    if wh_idx == -1:
                        wh_logger = self.logger
                    else:
                        wh_logger = get_task_logger(f"{self.task_id}-{wh_idx}")
                    
                    subtask_name = wh_config.get("name", "")
                    print_subtask_header(self.task_id, wh_idx if wh_idx >= 0 else 0, subtask_name, target_logger=wh_logger)
                    
                    wh_logger.info(f"处理Webhook：{wh_config['webhook'].split('key=')[-1][:8]}...")
                    
                    # 更新excel的logger为带子任务序号的logger
                    excel.logger = wh_logger
                    
                    screenshots, failed_capture_configs = excel.capture_screenshots(
                        wh_config["capture_configs"],
                        retry_times=3
                    )

                    if failed_capture_configs:
                        actual_failed = []
                        for item in failed_capture_configs:
                            start_day = item.get("start_day", 1)
                            end_day = item.get("end_day", 31)
                            if is_in_monthly_range(start_day, end_day):
                                actual_failed.append(item)
                        
                        if not actual_failed:
                            self.logger.info(f"{self._get_task_id_str(wh_idx)}所有失败的截图区域均为日期范围外，跳过失败通知")
                            print_subtask_footer(self.task_id, wh_idx if wh_idx >= 0 else 0, True, target_logger=wh_logger)
                            continue
                        
                        failed_regions_text = "；".join(
                            [
                                f"{item.get('name', '未命名')}({item.get('sheet_name', '未知工作表')}:{item.get('range', '未知区域')})"
                                for item in actual_failed
                            ]
                        )
                        task_id_str = self._get_task_id_str(wh_idx)
                        wh_logger.error(
                            f"{task_id_str}截图在重试 3 次后仍失败，共 {len(actual_failed)} 个区域：{failed_regions_text}"
                        )

                        if screenshots:
                            self._cleanup(screenshots, wh_logger)
                        self._send_wechat(
                            type="text",
                            data={
                                "content": (
                                    f"{task_id_str}截图失败：重试3次后仍有 {len(actual_failed)} 个区域未成功截图，"
                                    f"任务已终止。文件：{os.path.basename(self.config['excel_path'])}。"
                                    f"失败区域：{failed_regions_text}"
                                )
                            },
                            description="截图失败通知",
                            webhook=wh_config["webhook"],
                            task_logger=wh_logger
                        )
                        # 记录失败的webhook用于重试
                        failed_webhooks.append((wh_config, wh_idx))
                        print_subtask_footer(self.task_id, wh_idx if wh_idx >= 0 else 0, False, target_logger=wh_logger)
                        continue

                    results_to_deliver.append({
                        "screenshots": screenshots,
                        "webhook_config": wh_config,
                        "wh_idx": wh_idx
                    })
                    print_subtask_footer(self.task_id, wh_idx if wh_idx >= 0 else 0, True, target_logger=wh_logger)
                
                # 退出Excel上下文，释放文件
                excel = None
            
            # Excel已关闭，现在发送文件
            for result in results_to_deliver:
                # 获取对应子任务的logger
                wh_idx = result["wh_idx"]
                if wh_idx == -1:
                    wh_logger = self.logger
                else:
                    wh_logger = get_task_logger(f"{self.task_id}-{wh_idx}")
                deliver_result = self.deliver_results(result["screenshots"], result["webhook_config"], wh_logger)
                if deliver_result and not deliver_result.get("success"):
                    failed_webhooks.append((result["webhook_config"], wh_idx))

        except Exception as e:
            error_text = str(e)
            task_id_str = self._get_task_id_str()
            self.logger.error(f"{task_id_str}任务异常：{error_text}", exc_info=debug_mode)
            success = False
            
            # 清理已生成的截图文件
            if results_to_deliver:
                for result in results_to_deliver:
                    if result.get("screenshots"):
                        self._cleanup(result["screenshots"], self.logger)
            
            if "Excel 启动失败" in error_text:
                self._send_wechat(
                    type="text",
                    data={
                        "content": (
                            f"{task_id_str}任务启动失败：{os.path.basename(self.config['excel_path'])}\n"
                            f"错误信息：{error_text}"
                        ),
                        "mentioned_list": ["zhufuzhe"]
                    },
                    description="任务启动失败通知",
                    webhook=self.error_webhook
                )
        finally:
            # 清理当前excel实例引用
            self.current_excel = None
            
            elapsed_time = time.time() - start_time
            
            if success:
                self._backup_file()
            
            task_name = os.path.basename(self.config['excel_path'])
            print_task_footer(self.task_id, task_name, success, elapsed_time, target_logger=self.logger)
        
        # 返回任务结果和失败的webhook列表
        return success, failed_webhooks

    def deliver_results(self, screenshots: list, webhook_config: dict, task_logger=None):
        """根据webhook配置交付结果，所有配置的消息都发送成功才算任务成功"""
        if task_logger is None:
            task_logger = self.logger
        webhook = webhook_config["webhook"]
        send_file_enable = webhook_config.get("send_file_enable", 0)
        
        webhook_key = webhook.split("key=")[-1][:8]
        task_logger.info(f"_deliver_results 被调用")
        task_logger.info(f"  webhook: {webhook_key}...")
        task_logger.info(f"  screenshots数量: {len(screenshots)}")
        task_logger.info(f"  send_file_enable: {send_file_enable}")

        success = True

        # 先上传文件获取media_id（如果启用）
        media_id = None
        file_path = None
        if send_file_enable:
            task_logger.info("send_file_enable 为真，先上传文件")
            media_id = self._upload_attachment_for_send(webhook, task_logger)
            if not media_id:
                file_path = self.config.get("file_path") or self.config.get("excel_path")
                if file_path and os.path.exists(file_path):
                    MAX_FILE_SIZE = 20 * 1024 * 1024
                    file_size = os.path.getsize(file_path)
                    if file_size <= MAX_FILE_SIZE:
                        task_logger.error("文件上传失败")
                        success = False

        # 发送截图
        for img_path in screenshots:
            try:
                self._send_wechat(
                    type="image",
                    data=self._prepare_image(img_path),
                    description=f"截图 {os.path.basename(img_path)}",
                    webhook=webhook,
                    task_logger=task_logger
                )
            except Exception as e:
                task_logger.error(f"截图发送失败：{str(e)}")
                success = False

        # 发送文件（如果上传成功）
        if send_file_enable and media_id:
            task_logger.info("发送文件")
            try:
                file_path = self.config.get("file_path") or self.config.get("excel_path")
                self._send_wechat(
                    type="file",
                    data={"media_id": media_id},
                    description=f"文件 {os.path.basename(file_path)}",
                    webhook=webhook,
                    task_logger=task_logger
                )
                task_logger.info("文件发送成功")
            except Exception as e:
                task_logger.error(f"文件发送失败：{str(e)}")
                success = False
        elif send_file_enable and not media_id:
            task_logger.warning("文件上传失败，跳过发送文件")
        
        # 清理临时文件（仅当任务成功时）
        if success:
            self._cleanup(screenshots, task_logger)
        
        if not success:
            task_logger.error("文件发送失败，任务失败")
            return {"success": False, "reason": "file_send_failed"}
        
        return {"success": True, "reason": "success"}

    def _send_attachment(self, webhook: str, task_logger=None):
        """发送关联文件（带重试）"""
        if task_logger is None:
            task_logger = self.logger
        
        target_webhook = self.test_webhook if self.test_webhook else webhook
        
        file_path = self.config.get("file_path") or self.config.get("excel_path")
        task_logger.info(f"准备发送文件，file_path: {file_path}")
        task_logger.info(f"send_file_enable 检查，webhook: {target_webhook.split('key=')[-1][:8]}...")
        
        if not file_path:
            task_logger.warning("文件路径为空，跳过发送")
            return
        if not os.path.exists(file_path):
            task_logger.warning(f"文件不存在：{file_path}，跳过发送")
            return
        
        MAX_FILE_SIZE = 20 * 1024 * 1024
        file_size = os.path.getsize(file_path)
        if file_size > MAX_FILE_SIZE:
            task_logger.warning(f"文件大小 {file_size / 1024 / 1024:.2f} MB 超过企微webhook限制（20MB），取消发送文件")
            self._send_wechat(
                type="text",
                data={"content": f"⚠️ 文件发送提醒：{os.path.basename(file_path)} 大小为 {file_size / 1024 / 1024:.2f} MB，超过企微webhook限制（20MB），已取消发送"},
                description="文件大小超限提醒",
                webhook=target_webhook,
                task_logger=task_logger
            )
            return
        task_logger.info(f"文件大小: {file_size / 1024 / 1024:.2f} MB")
        
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            task_logger.info(f"开始发送文件（第 {attempt}/{max_retries} 次尝试）：{os.path.basename(file_path)}")
            try:
                with open(file_path, "rb") as f:
                    media_id = self._upload_file(f, target_webhook, task_logger)
                    if media_id:
                        task_logger.info(f"文件上传成功，media_id: {media_id}")
                        self._send_wechat(
                            type="file",
                            data={"media_id": media_id},
                            description=f"文件 {os.path.basename(file_path)}",
                            webhook=target_webhook,
                            task_logger=task_logger
                        )
                        task_logger.info(f"文件发送成功（第 {attempt}/{max_retries} 次尝试）")
                        return
                    else:
                        task_logger.warning(f"文件上传失败，未获取到media_id（第 {attempt}/{max_retries} 次尝试）")
            except Exception as e:
                task_logger.error(f"文件发送异常（第 {attempt}/{max_retries} 次尝试）：{str(e)}")
            
            if attempt < max_retries:
                task_logger.info(f"文件发送失败，{2 ** attempt}秒后重试...")
                time.sleep(2 ** attempt)
        
        task_logger.error(f"文件发送最终失败，已重试 {max_retries} 次")

    def _upload_file(self, file_obj, webhook: str, task_logger=None) -> str:
        """上传文件到临时素材（带重试）"""
        if task_logger is None:
            task_logger = self.logger
        
        upload_url = self._get_upload_url(webhook)
        task_logger.info(f"上传URL: {upload_url}")
        
        if not upload_url:
            task_logger.warning("无效的上传URL，跳过文件上传")
            return None
            
        filename = os.path.basename(file_obj.name)
        name, ext = os.path.splitext(filename)
        filename_with_time = f"{name}_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}{ext}"
        task_logger.info(f"上传文件名：{filename_with_time}")
        
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                file_obj.seek(0, os.SEEK_END)
                file_size = file_obj.tell()
                file_obj.seek(0)
                task_logger.info(f"文件大小: {file_size / 1024 / 1024:.2f} MB")
                task_logger.info(f"正在上传文件（第 {attempt}/{max_retries} 次尝试）：{file_obj.name}")
                
                response = requests.post(
                    upload_url,
                    files={"media": (filename_with_time, file_obj)},
                    timeout=60
                )
                
                task_logger.info(f"上传响应状态码: {response.status_code}")
                task_logger.info(f"上传响应内容: {response.text}")
                
                response.raise_for_status()
                result = response.json()
                task_logger.info(f"上传结果: {result}")
                return result.get("media_id")
            except Exception as e:
                task_logger.warning(f"文件上传失败（第 {attempt}/{max_retries} 次尝试）：{str(e)}")
                if attempt < max_retries:
                    time.sleep(3 * attempt)
                else:
                    task_logger.error(f"文件上传最终失败：{str(e)}")
        
        return None

    def _upload_attachment_for_send(self, webhook: str, task_logger=None) -> str:
        """上传文件获取media_id（用于deliver_results）"""
        if task_logger is None:
            task_logger = self.logger
        
        target_webhook = self.test_webhook if self.test_webhook else webhook
        file_path = self.config.get("file_path") or self.config.get("excel_path")
        
        if not file_path or not os.path.exists(file_path):
            task_logger.warning("文件路径无效，跳过上传")
            return None
        
        MAX_FILE_SIZE = 20 * 1024 * 1024
        file_size = os.path.getsize(file_path)
        if file_size > MAX_FILE_SIZE:
            task_logger.warning(f"文件大小 {file_size / 1024 / 1024:.2f} MB 超过企微webhook限制（20MB），取消上传")
            self._send_wechat(
                type="text",
                data={"content": f"⚠️ 文件发送提醒：{os.path.basename(file_path)} 大小为 {file_size / 1024 / 1024:.2f} MB，超过企微webhook限制（20MB），已取消发送"},
                description="文件大小超限提醒",
                webhook=target_webhook,
                task_logger=task_logger
            )
            return None
        
        task_logger.info(f"文件大小: {file_size / 1024 / 1024:.2f} MB")
        
        with open(file_path, "rb") as f:
            return self._upload_file(f, target_webhook, task_logger)

    def _prepare_image(self, img_path: str) -> dict:
        """准备图片数据"""
        max_size = 2 * 1024 * 1024
        min_width = 800
        min_height = 600

        with open(img_path, "rb") as f:
            img_data = f.read()
            if len(img_data) > max_size:
                img = Image.open(io.BytesIO(img_data))
                img = img.convert("RGB")
                buf = io.BytesIO()
                quality = 85

                while True:
                    buf.seek(0)
                    img.save(buf, format="JPEG", quality=quality)
                    if buf.tell() <= max_size or quality <= 60:
                        break
                    quality -= 5

                if buf.tell() > max_size:
                    width, height = img.size
                    while buf.tell() > max_size and width > min_width and height > min_height:
                        width = int(width * 0.9)
                        height = int(height * 0.9)
                        img = img.resize((width, height), Image.LANCZOS)
                        buf.seek(0)
                        img.save(buf, format="JPEG", quality=quality)
                img_data = buf.getvalue()

        return {
            "base64": base64.b64encode(img_data).decode(),
            "md5": hashlib.md5(img_data).hexdigest()
        }

    def _send_wechat(self, type: str, data: dict, description: str, webhook: str, task_logger=None):
        """发送到企业微信（带重试）"""
        if task_logger is None:
            task_logger = self.logger
        # 如果是测试模式，使用测试webhook
        target_webhook = self.test_webhook if self.test_webhook else webhook
        
        payload = {"msgtype": type, type: data}
        
        for attempt in range(1, self.retry_limit+1):
            try:
                response = requests.post(
                    target_webhook,
                    json=payload,
                    timeout=10
                )
                response.raise_for_status()
                result = response.json()
                if result.get("errcode", 0) != 0:
                    raise Exception(f"企业微信API返回错误：errcode={result.get('errcode')}, errmsg={result.get('errmsg')}")
                task_logger.info(f"发送成功：{description}")
                return
            except Exception as e:
                task_logger.warning(f"发送失败（{attempt}/{self.retry_limit}）：{description} - {str(e)}")
                if attempt == self.retry_limit:
                    task_logger.error(f"最终发送失败：{str(e)}")
                time.sleep(2 ** attempt)

    def send_final_failure_notification(self):
        """发送最终失败通知（仅在第三次重试失败后发送）"""
        task_id_str = self._get_task_id_str()
        task_name = self.config.get("name", os.path.basename(self.config["excel_path"]))
        
        self._send_wechat(
            type="text",
            data={
                "content": (
                    f"{task_id_str}任务最终失败通知\n"
                    f"任务名称：{task_name}\n"
                    f"重试次数：{self.retry_count}次\n"
                    f"文件路径：{self.config['excel_path']}\n"
                    f"任务已达到最大重试次数（{self.retry_max_attempts}次），请检查文件或相关配置。"
                ),
                "mentioned_list": ["zhufuzhe"]
            },
            description="任务最终失败通知",
            webhook=self.error_webhook
        )

    def _cleanup(self, files: list, task_logger=None):
        """清理临时文件"""
        if task_logger is None:
            task_logger = self.logger
        for f in files:
            try:
                os.remove(f)
                task_logger.debug(f"清理临时文件：{os.path.basename(f)}")
            except Exception as e:
                task_logger.warning(f"文件清理失败：{str(e)}")

    def _backup_file(self):
        """
        在任务结束后备份文件
        目录结构：backup_dir/YYYY-MM/文件名_扩展名/文件名_YYYYMMDD_HHMMSS.扩展名
        
        设计说明：
        - 年月目录：按 YYYY-MM 格式组织，便于按时间查找
        - 文件目录：采用 文件名_扩展名 格式，避免同文件名不同扩展名冲突
        - 备份文件名：文件名_日期时间.扩展名，确保唯一性
        """
        backup_config = self.config.get("backup", {})
        if not backup_config.get("enable", False):
            self.logger.info("备份功能未启用")
            return
        
        source_path = self.config.get("file_path") or self.config.get("excel_path")
        if not source_path or not os.path.exists(source_path):
            self.logger.warning(f"源文件不存在，跳过备份：{source_path}")
            return
        
        backup_dir = backup_config.get("backup_dir", "./backups")
        if not os.path.isabs(backup_dir):
            backup_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), backup_dir))
        
        now = datetime.now()
        year_month = now.strftime("%Y-%m")
        file_name = os.path.basename(source_path)
        file_base, file_ext = os.path.splitext(file_name)
        ext_without_dot = file_ext[1:].lower() if file_ext else ""
        
        month_dir = os.path.join(backup_dir, year_month)
        file_dir = os.path.join(month_dir, f"{file_base}_{ext_without_dot}")
        
        try:
            os.makedirs(file_dir)
            
            timestamp = now.strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{file_base}_{timestamp}{file_ext}"
            backup_path = os.path.join(file_dir, backup_filename)
            
            shutil.copy2(source_path, backup_path)
            
            self.logger.info(f"文件备份成功：{backup_path}")
            return backup_path
        except FileExistsError:
            timestamp = now.strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{file_base}_{timestamp}{file_ext}"
            backup_path = os.path.join(file_dir, backup_filename)
            
            shutil.copy2(source_path, backup_path)
            
            self.logger.info(f"文件备份成功：{backup_path}")
            return backup_path
        except Exception as e:
            error_msg = f"文件备份失败：{str(e)}"
            self.logger.error(error_msg)
            
            self._send_wechat(
                type="text",
                data={
                    "content": f"[{self.task_id}]文件备份失败\n文件：{file_name}\n错误：{str(e)}",
                    "mentioned_list": ["zhufuzhe"]
                },
                description="文件备份失败通知",
                webhook=self.error_webhook
            )
            
            return None

# ---------------------------- 任务调度器 ----------------------------
class TaskScheduler:
    """多任务调度引擎"""

    def __init__(self, config_path: str, debug=False, test_webhook=None, error_webhook=None):
        self.config_path = config_path
        self.tasks = []
        self.retry_config = {}
        self._load_config(test_webhook, error_webhook)
        self.debug_mode = debug
        self._scheduler_lock = None
        self._retry_jobs = {}  # 存储重试任务 {task_id: job}
        logger.setLevel(logging.DEBUG if debug else logging.INFO)

    def _load_config(self, test_webhook=None, error_webhook=None):
        """加载配置文件"""
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)

            if not isinstance(config.get("tasks"), list):
                raise ValueError("配置文件格式错误")

            if error_webhook is None and config.get("error_webhook"):
                error_webhook = config["error_webhook"]

            upload_url_template = config.get("upload_url_template", "")
            
            backup_config = config.get("backup", {})
            
            if backup_config.get("enable", False):
                backup_dir = backup_config.get("backup_dir", "./backups")
                if not os.path.isabs(backup_dir):
                    backup_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), backup_dir))
                
                if not os.path.exists(backup_dir):
                    error_msg = f"备份目录不存在：{backup_dir}"
                    logger.error(error_msg)
                    print(f"错误：{error_msg}")
                    print("请先创建备份目录或修改配置文件中的 backup_dir 路径")
                    exit(1)
            
            # 加载重试配置
            self.retry_config = config.get("retry", {})
            if self.retry_config.get("enabled", 0) == 1:
                timeout_minutes = self.retry_config.get("task_timeout_minutes", 15)
                logger.info(f"重试机制已启用，延迟时间：{self.retry_config.get('delay_minutes', 10)}分钟，最大重试次数：{self.retry_config.get('max_attempts', 3)}次，任务超时时间：{timeout_minutes}分钟")
            else:
                logger.info("重试机制已禁用")
            
            # 加载任务
            for idx, task in enumerate(config["tasks"]):
                task["backup"] = backup_config
                self.tasks.append(ReportTask(task, test_webhook, error_webhook, upload_url_template, idx, self.retry_config))
            
            logger.info(f"成功加载 {len(self.tasks)} 个任务")
        except Exception as e:
            logger.error(f"配置加载失败：{str(e)}")
            raise

    def start(self):
        """启动调度服务"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        locks_dir = os.path.join(base_dir, "locks")
        os.makedirs(locks_dir, exist_ok=True)
        scheduler_lock_path = os.path.join(locks_dir, "scheduler.lock")
        self._scheduler_lock = FileLock(scheduler_lock_path)
        
        if not self._scheduler_lock.acquire(timeout=5):
            logger.error("检测到已有调度器实例在运行，退出...")
            return
        
        logger.info("启动任务调度器...")
        self._schedule_tasks()
        
        try:
            while True:
                schedule.run_pending()
                time.sleep(1)
        except KeyboardInterrupt:
            logger.info("正在关闭调度器...")
        finally:
            if self._scheduler_lock:
                self._scheduler_lock.release()
                self._scheduler_lock = None

    def _get_task_execution_lock(self):
        """获取全局任务执行锁，确保同一时间只有一个任务在执行"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        locks_dir = os.path.join(base_dir, "locks")
        os.makedirs(locks_dir, exist_ok=True)
        lock_file_path = os.path.join(locks_dir, "task_execution.lock")
        return FileLock(lock_file_path)

    def _schedule_tasks(self):
        """配置定时任务，相同时间点的多个webhook合并为一个任务，只刷新一次"""
        time_tasks = {}
        
        for task_idx, task in enumerate(self.tasks):
            task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
            
            for webhook_idx, webhook_config in enumerate(task.config["webhooks"]):
                webhook_key = webhook_config["webhook"].split("key=")[-1][:8]
                for trigger_time in webhook_config["times"]:
                    key = (task_idx, trigger_time)
                    if key not in time_tasks:
                        time_tasks[key] = {
                            "task": task,
                            "task_name": task_name,
                            "trigger_time": trigger_time,
                            "webhook_configs": []
                        }
                    time_tasks[key]["webhook_configs"].append({
                        "idx": webhook_idx,
                        "config": webhook_config,
                        "key": webhook_key
                    })
        
        for key, info in time_tasks.items():
            task_idx, trigger_time = key
            task = info["task"]
            task_name = info["task_name"]
            webhook_configs = [wh["config"] for wh in info["webhook_configs"]]
            webhook_keys = [f"[{wh['idx']}][{wh['key']}]" for wh in info["webhook_configs"]]
            
            schedule.every().day.at(trigger_time).do(
                self._run_task, task, webhook_configs
            )
            logger.info(f"已安排任务：{trigger_time} → [{task_idx}] {task_name} → webhooks:{','.join(webhook_keys)}")

    def _run_task(self, task: ReportTask, webhook_configs: list, is_retry=False, retry_webhook_idx=None):
        """串行执行任务（支持多个webhook配置共享一次刷新，带超时控制）"""
        task_lock = self._get_task_execution_lock()
        if not task_lock.acquire(timeout=300):
            task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
            error_msg = f"任务 [{task.task_id}] {task_name} 获取全局任务执行锁超时，跳过本次执行"
            logger.error(error_msg)
            
            try:
                task._send_wechat(
                    type="text",
                    data={
                        "content": (
                            f"⚠️ 任务锁获取失败通知\n"
                            f"任务ID：{task.task_id}\n"
                            f"任务名称：{task_name}\n"
                            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                            f"原因：获取全局任务执行锁超时（300秒）\n"
                            f"可能原因：另一个任务正在执行或锁文件被异常占用"
                        ),
                        "mentioned_list": ["zhufuzhe"]
                    },
                    description="任务锁获取失败通知",
                    webhook=task.error_webhook
                )
            except Exception as e:
                logger.error(f"发送锁获取失败通知异常：{str(e)}")
            
            return
        
        # 任务执行结果
        execution_result = {
            "success": False,
            "failed_webhooks": [],
            "exception": None
        }
        
        def task_executor():
            """任务执行器（在子线程中运行）"""
            try:
                pythoncom.CoInitialize()
                separator = "=" * 100
                logger.info("")
                if is_retry:
                    if retry_webhook_idx is not None:
                        logger.info(f"Webhook重试任务触发 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    else:
                        logger.info(f"重试任务触发 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                else:
                    logger.info(f"定时任务触发 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                logger.info(separator)
                
                webhook_keys = [wh["webhook"].split("key=")[-1][:8] for wh in webhook_configs]
                logger.info(f"本次任务将发送到 {len(webhook_configs)} 个 webhook: {','.join(webhook_keys)}")
                
                success, failed_webhooks = task.execute(self.debug_mode, webhook_configs)
                execution_result["success"] = success
                execution_result["failed_webhooks"] = failed_webhooks
            except Exception as e:
                execution_result["exception"] = e
                execution_result["success"] = False
                logger.error(f"任务执行异常：{str(e)}")
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        
        try:
            # 获取超时时间（分钟转秒）
            timeout_seconds = task.task_timeout_minutes * 60
            task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
            logger.info(f"任务 [{task.task_id}] {task_name} 执行开始，超时时间：{task.task_timeout_minutes}分钟")
            
            # 在子线程中执行任务
            exec_thread = threading.Thread(target=task_executor, daemon=True)
            exec_thread.start()
            
            # 等待任务完成或超时
            exec_thread.join(timeout=timeout_seconds)
            
            if exec_thread.is_alive():
                # 任务超时，强制终止
                logger.error(f"任务 [{task.task_id}] {task_name} 执行超时（{task.task_timeout_minutes}分钟），开始强制终止")
                
                # 发送超时通知
                try:
                    task._send_wechat(
                        type="text",
                        data={
                            "content": (
                                f"⚠️ 任务执行超时通知\n"
                                f"任务ID：{task.task_id}\n"
                                f"任务名称：{task_name}\n"
                                f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                                f"原因：任务执行时间超过 {task.task_timeout_minutes} 分钟，已强制终止\n"
                                f"文件：{os.path.basename(task.config['excel_path'])}\n"
                                f"任务将在 {task.retry_delay_minutes} 分钟后重试"
                            ),
                            "mentioned_list": ["zhufuzhe"]
                        },
                        description="任务执行超时通知",
                        webhook=task.error_webhook
                    )
                except Exception as e:
                    logger.error(f"发送超时通知异常：{str(e)}")
                
                # 强制终止Excel进程
                if task.current_excel:
                    try:
                        task.current_excel.force_terminate()
                    except Exception as e:
                        logger.error(f"强制终止Excel进程异常：{str(e)}")
                
                # 标记任务失败，触发重试
                execution_result["success"] = False
                execution_result["failed_webhooks"] = []
            else:
                # 任务正常完成
                logger.info(f"任务 [{task.task_id}] {task_name} 执行完成")
            
            trigger_time = datetime.now().strftime("%H:%M")
            record_execution(task.task_id, task_name, trigger_time, execution_result["success"], manual=False)
            
            # 如果是重试任务，成功后清除重试记录
            if is_retry and execution_result["success"]:
                task.retry_count = 0
                if retry_webhook_idx is not None and retry_webhook_idx in task.webhook_retry_counts:
                    task.webhook_retry_counts[retry_webhook_idx] = 0
                if task.task_id in self._retry_jobs:
                    schedule.cancel_job(self._retry_jobs[task.task_id])
                    del self._retry_jobs[task.task_id]
            
            # 处理Webhook级重试（文件成功但部分Webhook失败）
            if task.retry_enabled and execution_result["failed_webhooks"]:
                for wh_config, wh_idx in execution_result["failed_webhooks"]:
                    self._schedule_webhook_retry(task, wh_config, wh_idx)
                # 如果有Webhook级重试，任务不算完全失败
                if execution_result["success"] is False:
                    execution_result["success"] = True  # 标记为部分成功，避免触发文件级重试
            
            # 如果是重试任务且有Webhook失败，重新调度这些Webhook的重试
            if is_retry and retry_webhook_idx is not None and execution_result["success"]:
                # 清理该Webhook的重试计数
                if retry_webhook_idx in task.webhook_retry_counts:
                    task.webhook_retry_counts[retry_webhook_idx] = 0
            
            # 如果任务失败（文件级失败）且启用了重试机制，触发文件级重试
            if not execution_result["success"] and task.retry_enabled:
                self._schedule_retry(task, webhook_configs)
        finally:
            task_lock.release()

    def _schedule_retry(self, task: ReportTask, webhook_configs: list):
        """调度重试任务"""
        if not task.retry_enabled:
            return
        
        # 检查是否已过午夜0点，如果是则取消重试
        now = datetime.now()
        if now.hour < 6:
            logger.info(f"任务 [{task.task_id}] 当前时间为午夜0点后，取消重试")
            task.send_final_failure_notification()
            return
        
        task.retry_count += 1
        max_attempts = task.retry_max_attempts
        
        if task.retry_count > max_attempts:
            logger.info(f"任务 [{task.task_id}] 已达到最大重试次数 ({max_attempts}次)，不再重试")
            task.send_final_failure_notification()
            return
        
        # 计算重试时间
        retry_delay_seconds = task.retry_delay_minutes * 60
        retry_time = datetime.now() + timedelta(seconds=retry_delay_seconds)
        retry_time_str = retry_time.strftime("%H:%M")
        
        # 检查重试时间是否与其他配置任务时间冲突
        conflict_time = self._check_time_conflict(task, retry_time_str)
        if conflict_time:
            # 如果冲突，将重试时间延迟到配置任务执行时间之后（+1分钟）
            logger.info(f"重试时间 {retry_time_str} 与配置任务时间 {conflict_time} 冲突，延迟到 {conflict_time} 之后")
            # 解析冲突时间并延迟1分钟
            conflict_hour, conflict_minute = map(int, conflict_time.split(":"))
            retry_datetime = datetime.now().replace(hour=conflict_hour, minute=conflict_minute, second=0)
            # 如果冲突时间已过，设置为下一天
            if retry_datetime <= datetime.now():
                retry_datetime += timedelta(days=1)
            # 延迟1分钟
            retry_datetime += timedelta(minutes=1)
            retry_time_str = retry_datetime.strftime("%H:%M")
            # 重新计算延迟时间
            retry_delay_seconds = (retry_datetime - datetime.now()).total_seconds()
            if retry_delay_seconds < 0:
                retry_delay_seconds = 60  # 至少延迟1分钟
        
        task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
        logger.info(f"任务 [{task.task_id}] {task_name} 失败，第 {task.retry_count}/{max_attempts} 次重试将在 {retry_time_str} 执行（延迟 {retry_delay_seconds/60:.1f} 分钟）")
        
        # 如果存在未执行的重试任务，先取消
        if task.task_id in self._retry_jobs:
            schedule.cancel_job(self._retry_jobs[task.task_id])
            del self._retry_jobs[task.task_id]
        
        # 调度重试任务（使用延迟调度，只执行一次）
        job = schedule.every(retry_delay_seconds).seconds.do(
            self._run_task, task, webhook_configs, True
        )
        self._retry_jobs[task.task_id] = job

    def _schedule_webhook_retry(self, task: ReportTask, wh_config: dict, wh_idx: int):
        """调度Webhook级重试任务"""
        if not task.retry_enabled:
            return
        
        # 检查是否已过午夜0点，如果是则取消重试
        now = datetime.now()
        if now.hour < 6:
            wh_key = wh_config["webhook"].split("key=")[-1][:8]
            logger.info(f"任务 [{task.task_id}] Webhook[{wh_idx}]({wh_key}...) 当前时间为午夜0点后，取消重试")
            task.send_final_failure_notification()
            return
        
        # 获取或初始化该Webhook的重试计数
        if wh_idx not in task.webhook_retry_counts:
            task.webhook_retry_counts[wh_idx] = 0
        
        task.webhook_retry_counts[wh_idx] += 1
        max_attempts = task.retry_max_attempts
        current_retry = task.webhook_retry_counts[wh_idx]
        
        if current_retry > max_attempts:
            wh_key = wh_config["webhook"].split("key=")[-1][:8]
            logger.info(f"任务 [{task.task_id}] Webhook[{wh_idx}]({wh_key}...) 已达到最大重试次数 ({max_attempts}次)，不再重试")
            task.send_final_failure_notification()
            return
        
        # 计算重试时间
        retry_delay_seconds = task.retry_delay_minutes * 60
        retry_time = datetime.now() + timedelta(seconds=retry_delay_seconds)
        retry_time_str = retry_time.strftime("%H:%M")
        
        # 检查重试时间是否与其他配置任务时间冲突
        conflict_time = self._check_time_conflict(task, retry_time_str)
        if conflict_time:
            logger.info(f"Webhook重试时间 {retry_time_str} 与配置任务时间 {conflict_time} 冲突，延迟到 {conflict_time} 之后")
            conflict_hour, conflict_minute = map(int, conflict_time.split(":"))
            retry_datetime = datetime.now().replace(hour=conflict_hour, minute=conflict_minute, second=0)
            if retry_datetime <= datetime.now():
                retry_datetime += timedelta(days=1)
            retry_datetime += timedelta(minutes=1)
            retry_time_str = retry_datetime.strftime("%H:%M")
        
        task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
        wh_key = wh_config["webhook"].split("key=")[-1][:8]
        logger.info(f"任务 [{task.task_id}] Webhook[{wh_idx}]({wh_key}...) 失败，第 {current_retry}/{max_attempts} 次重试将在 {retry_time_str} 执行（延迟 {retry_delay_seconds/60:.1f} 分钟）")
        
        # 使用唯一的key存储Webhook重试任务
        webhook_retry_key = f"{task.task_id}_wh_{wh_idx}"
        if webhook_retry_key in self._retry_jobs:
            schedule.cancel_job(self._retry_jobs[webhook_retry_key])
        
        # 调度Webhook重试任务（只重试该Webhook，使用延迟调度，只执行一次）
        job = schedule.every(retry_delay_seconds).seconds.do(
            self._run_webhook_retry, task, wh_config, wh_idx, True
        )
        self._retry_jobs[webhook_retry_key] = job

    def _run_webhook_retry(self, task: ReportTask, wh_config: dict, wh_idx: int, is_retry=True):
        """执行单个Webhook的重试"""
        pythoncom.CoInitialize()
        try:
            wh_logger = get_task_logger(f"{task.task_id}-{wh_idx}")
            separator = "=" * 100
            logger.info("")
            logger.info(f"Webhook重试任务执行 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info(separator)
            
            wh_key = wh_config["webhook"].split("key=")[-1][:8]
            wh_logger.info(f"重试Webhook[{wh_idx}]({wh_key}...)")
            
            # 重新打开Excel并只刷新该Webhook的数据
            with ExcelProcessor(task.config["excel_path"], visible=self.debug_mode, task_logger=wh_logger) as excel:
                # 只对失败的Webhook执行截图
                screenshots, failed_capture_configs = excel.capture_screenshots(
                    wh_config["capture_configs"],
                    retry_times=3
                )
                
                if failed_capture_configs:
                    # 截图仍然失败，触发下次重试
                    wh_logger.error(f"Webhook[{wh_idx}] 重试截图仍失败，{len(failed_capture_configs)} 个区域未成功")
                    self._schedule_webhook_retry(task, wh_config, wh_idx)
                else:
                    # 截图成功，发送结果
                    wh_logger.info(f"Webhook[{wh_idx}] 重试截图成功")
                    result = task.deliver_results(screenshots, wh_config, wh_logger)
                    if result and not result.get("success"):
                        wh_logger.error(f"Webhook[{wh_idx}] 文件发送失败，触发文件上传级重试")
                        file_path = task.config.get("file_path") or task.config.get("excel_path")
                        self._schedule_file_retry(task, wh_config, screenshots, file_path)
                    else:
                        if wh_idx in task.webhook_retry_counts:
                            task.webhook_retry_counts[wh_idx] = 0
                    
        except Exception as e:
            wh_logger.error(f"Webhook[{wh_idx}] 重试异常：{str(e)}")
            self._schedule_webhook_retry(task, wh_config, wh_idx)
        finally:
            pythoncom.CoUninitialize()

    def _schedule_file_retry(self, task: ReportTask, wh_config: dict, screenshots: list, file_path: str):
        """调度文件上传级重试（只重试文件上传和发送，不重新刷新Excel）"""
        if not task.retry_enabled:
            return
        
        now = datetime.now()
        if now.hour < 6:
            wh_key = wh_config["webhook"].split("key=")[-1][:8]
            logger.info(f"任务 [{task.task_id}] 文件上传级重试 当前时间为午夜0点后，取消重试")
            task.send_final_failure_notification()
            return
        
        wh_idx = task.config["webhooks"].index(wh_config)
        
        if not hasattr(task, 'file_retry_counts'):
            task.file_retry_counts = {}
        if wh_idx not in task.file_retry_counts:
            task.file_retry_counts[wh_idx] = 0
        
        task.file_retry_counts[wh_idx] += 1
        max_attempts = task.retry_max_attempts
        current_retry = task.file_retry_counts[wh_idx]
        
        if current_retry > max_attempts:
            wh_key = wh_config["webhook"].split("key=")[-1][:8]
            logger.info(f"任务 [{task.task_id}] 文件上传级重试已达到最大次数 ({max_attempts}次)，不再重试")
            task.send_final_failure_notification()
            return
        
        retry_delay_seconds = task.retry_delay_minutes * 60
        retry_time = datetime.now() + timedelta(seconds=retry_delay_seconds)
        retry_time_str = retry_time.strftime("%H:%M")
        
        conflict_time = self._check_time_conflict(task, retry_time_str)
        if conflict_time:
            logger.info(f"文件上传级重试时间 {retry_time_str} 与配置任务时间 {conflict_time} 冲突，延迟到 {conflict_time} 之后")
            conflict_hour, conflict_minute = map(int, conflict_time.split(":"))
            retry_datetime = datetime.now().replace(hour=conflict_hour, minute=conflict_minute, second=0)
            if retry_datetime <= datetime.now():
                retry_datetime += timedelta(days=1)
            retry_datetime += timedelta(minutes=1)
            retry_time_str = retry_datetime.strftime("%H:%M")
            retry_delay_seconds = (retry_datetime - datetime.now()).total_seconds()
            if retry_delay_seconds < 0:
                retry_delay_seconds = 60
        
        task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
        wh_key = wh_config["webhook"].split("key=")[-1][:8]
        logger.info(f"[任务{task.task_id}] {task_name} - 调度文件上传级重试，时间：{retry_time_str}，webhook: {wh_key}...")
        
        file_retry_key = f"{task.task_id}_file_{wh_idx}"
        if file_retry_key in self._retry_jobs:
            schedule.cancel_job(self._retry_jobs[file_retry_key])
        
        job = schedule.every(retry_delay_seconds).seconds.do(
            self._run_file_retry, task, wh_config, screenshots, file_path, wh_idx
        )
        self._retry_jobs[file_retry_key] = job

    def _run_file_retry(self, task: ReportTask, wh_config: dict, screenshots: list, file_path: str, wh_idx: int):
        """执行文件上传级重试（跳过Excel刷新和截图）"""
        task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
        webhook_key = wh_config["webhook"].split("key=")[-1][:8]
        wh_logger = get_task_logger(task.task_id)
        
        file_retry_key = f"{task.task_id}_file_{wh_idx}"
        
        wh_logger.info(f"[任务{task.task_id}] {task_name} - 执行文件上传级重试，webhook: {webhook_key}...")
        
        try:
            result = task.deliver_results(screenshots, wh_config, wh_logger)
            
            if result and result.get("success"):
                wh_logger.info(f"[任务{task.task_id}] {task_name} - 文件上传级重试成功")
                if hasattr(task, 'file_retry_counts') and wh_idx in task.file_retry_counts:
                    task.file_retry_counts[wh_idx] = 0
            else:
                wh_logger.error(f"[任务{task.task_id}] {task_name} - 文件上传级重试失败")
                
                if not hasattr(task, 'file_retry_counts'):
                    task.file_retry_counts = {}
                max_attempts = task.retry_max_attempts
                current_retry = task.file_retry_counts.get(wh_idx, 0)
                
                if current_retry < max_attempts:
                    self._schedule_file_retry(task, wh_config, screenshots, file_path)
                else:
                    wh_logger.error(f"[任务{task.task_id}] {task_name} - 文件上传级重试已达最大次数({max_attempts})，放弃重试")
                    if screenshots:
                        task._cleanup(screenshots, wh_logger)
        
        except Exception as e:
            wh_logger.error(f"[任务{task.task_id}] {task_name} - 文件上传级重试异常：{str(e)}")
            
            if not hasattr(task, 'file_retry_counts'):
                task.file_retry_counts = {}
            max_attempts = task.retry_max_attempts
            current_retry = task.file_retry_counts.get(wh_idx, 0)
            
            if current_retry < max_attempts:
                self._schedule_file_retry(task, wh_config, screenshots, file_path)
            else:
                wh_logger.error(f"[任务{task.task_id}] {task_name} - 文件上传级重试已达最大次数({max_attempts})，放弃重试")
                if screenshots:
                    task._cleanup(screenshots, wh_logger)
        finally:
            if file_retry_key in self._retry_jobs:
                schedule.cancel_job(self._retry_jobs[file_retry_key])
                del self._retry_jobs[file_retry_key]

    def _check_time_conflict(self, task: ReportTask, retry_time: str) -> str:
        """检查重试时间是否与其他配置任务时间冲突"""
        # 收集所有已配置的任务时间点
        configured_times = set()
        
        for t in self.tasks:
            for webhook_config in t.config["webhooks"]:
                for time_str in webhook_config["times"]:
                    configured_times.add(time_str)
        
        # 检查重试时间是否与配置时间冲突
        if retry_time in configured_times:
            return retry_time
        
        return ""

    def run_now(self, task_specs: list = None):
        """立即执行任务（调试）
        :param task_specs: 任务规格列表，每个元素为 (task_id, webhook_id) 元组
        """
        logger.info("进入调试模式...")
        
        # 确定要执行的任务列表
        if not task_specs:
            # 没有指定任务，执行所有任务
            targets = [(task, None) for task in self.tasks]
        else:
            # 解析任务规格列表
            targets = []
            for task_id, webhook_id in task_specs:
                task = self.tasks[task_id]
                targets.append((task, webhook_id))
        
        # 执行任务
        for idx, (task, webhook_id) in enumerate(targets, 1):
            task_lock = self._get_task_execution_lock()
            if not task_lock.acquire(timeout=300):
                task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
                error_msg = f"任务 [{task.task_id}] {task_name} 获取全局任务执行锁超时，跳过执行"
                logger.error(error_msg)
                
                try:
                    task._send_wechat(
                        type="text",
                        data={
                            "content": (
                                f"⚠️ 任务锁获取失败通知（手动执行）\n"
                                f"任务ID：{task.task_id}\n"
                                f"任务名称：{task_name}\n"
                                f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                                f"原因：获取全局任务执行锁超时（300秒）\n"
                                f"可能原因：另一个任务正在执行或锁文件被异常占用"
                            ),
                            "mentioned_list": ["zhufuzhe"]
                        },
                        description="任务锁获取失败通知（手动执行）",
                        webhook=task.error_webhook
                    )
                except Exception as e:
                    logger.error(f"发送锁获取失败通知异常：{str(e)}")
                
                continue
            
            try:
                pythoncom.CoInitialize()
                try:
                    if len(targets) > 1:
                        separator = "=" * 100
                        logger.info("")
                        logger.info(f"任务 {idx}/{len(targets)}")
                        logger.info(separator)
                    
                    trigger_time = datetime.now().strftime("%H:%M")
                    task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
                    
                    # 确定要执行的webhook
                    if webhook_id is not None:
                        webhook_config = task.config["webhooks"][webhook_id]
                        logger.info(f"执行任务 {task_specs[idx-1][0]} 的 webhook {webhook_id}: {webhook_config['webhook'].split('key=')[-1][:10]}...")
                        success = task.execute(self.debug_mode, webhook_config, is_manual=True)
                    else:
                        # 执行所有webhook配置
                        if task_specs:
                            logger.info(f"执行任务 {task_specs[idx-1][0]}: {task_name}")
                        success = task.execute(self.debug_mode, is_manual=True)
                    
                    record_execution(task.task_id, task_name, trigger_time, success, manual=True)
                except Exception as e:
                    logger.error(f"执行异常：{str(e)}")
                    if task_specs:
                        task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
                        record_execution(task.task_id, task_name, datetime.now().strftime("%H:%M"), False, manual=True)
                finally:
                    pythoncom.CoUninitialize()
            finally:
                task_lock.release()

# ---------------------------- 主程序 ----------------------------
def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description="Excel 自动化")
    parser.add_argument("--run-all", action="store_true", help="立即执行所有任务")
    parser.add_argument("--task", type=str, nargs='*', help="执行指定序号的任务，格式：任务索引 或 任务索引-webhook索引，可指定多个")
    parser.add_argument("--list", action="store_true", help="列出所有已配置的任务")
    parser.add_argument("--debug", action="store_true", help="开启调试模式")
    parser.add_argument("--test", action="store_true", help="启用测试webhook，发送到测试群")
    args = parser.parse_args()

    test_webhook = None
    if args.test:
        logger.info("已启用测试webhook模式")
        with open("config.yml", "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            if config.get("test_webhook"):
                test_webhook = config["test_webhook"]

    try:
        if args.list:
            print_task_list()
            return
        
        task_specs = []
        if args.task:
            for task_str in args.task:
                if "-" in task_str:
                    task_id_str, webhook_id_str = task_str.split("-", 1)
                    task_id = int(task_id_str)
                    webhook_id = int(webhook_id_str)
                    task_specs.append((task_id, webhook_id))
                else:
                    task_id = int(task_str)
                    task_specs.append((task_id, None))
        
        scheduler = TaskScheduler("config.yml", debug=args.debug, test_webhook=test_webhook)
        
        with open("config.yml", "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            backup_config = config.get("backup", {})
            backup_enable = backup_config.get("enable", 0)
            backup_dir = backup_config.get("backup_dir", "./backups")
            if not os.path.isabs(backup_dir):
                backup_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), backup_dir))
            
            print(f"备份配置: enable={backup_enable}, backup_dir={backup_dir}")
            print()

        if args.run_all or args.task is not None:
            scheduler.run_now(task_specs)
        else:
            scheduler.start()
    except Exception as e:
        logger.error(f"系统异常：{str(e)}", exc_info=args.debug)
        exit(1)

def print_task_list():
    """打印所有已配置的任务列表及当日执行情况"""
    try:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yml")
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        
        backup_config = config.get("backup", {})
        backup_enable = backup_config.get("enable", False)
        backup_dir = backup_config.get("backup_dir", "./backups")
        
        execution_log = load_execution_log()
        today = datetime.now().strftime("%Y-%m-%d")
        current_time = datetime.now().strftime("%H:%M")
        
        today_executions = execution_log.get(today, {}).get("tasks", {})
        
        tasks = config.get("tasks", [])
        
        print(f"{COLOR_BLUE}═══════════════════════════════════════════════════════════════════════════════════{COLOR_RESET}")
        print(f"{COLOR_BLUE}                    任务列表 - {today} {current_time}{COLOR_RESET}")
        print(f"{COLOR_BLUE}═══════════════════════════════════════════════════════════════════════════════════{COLOR_RESET}")
        print()
        
        print(f"备份配置: enable={backup_enable}, backup_dir={backup_dir}")
        print()
        
        total_success = 0
        total_failed = 0
        total_pending = 0
        
        for idx, task in enumerate(tasks):
            task_name = task.get("name", os.path.basename(task["excel_path"]))
            webhooks = task.get("webhooks", [])
            
            all_times = []
            for wh_config in webhooks:
                all_times.extend(wh_config.get("times", []))
            
            all_times_set = set(all_times)
            all_times = sorted(all_times_set)
            
            task_executions = today_executions.get(str(idx), {}).get("executions", {})
            
            success_count = 0
            failed_count = 0
            pending_count = 0
            
            manual_executions = []
            
            for t, execution in task_executions.items():
                if t not in all_times_set:
                    if isinstance(execution, dict):
                        status = execution.get("status")
                        manual = execution.get("manual", False)
                    else:
                        status = execution
                        manual = True
                    
                    if manual:
                        status_icon = f"{COLOR_GREEN}✓{COLOR_RESET}" if status == "success" else f"{COLOR_RED}✗{COLOR_RESET}"
                        manual_executions.append(f"{status_icon} {t}")
            
            import re
            def get_display_length(s):
                return len(re.sub(r'\x1b\[[0-9;]*m', '', s))
            
            all_lines = []
            
            for wh_idx, wh_config in enumerate(webhooks):
                wh_name = wh_config.get("name", "")
                wh_times = wh_config.get("times", [])
                wh_times = sorted(wh_times)
                
                wh_line_parts = []
                for t in wh_times:
                    execution = task_executions.get(t)
                    if isinstance(execution, dict):
                        status = execution.get("status")
                    else:
                        status = execution
                    
                    if status == "success":
                        wh_line_parts.append(f"{COLOR_GREEN}✓{COLOR_RESET} {t}")
                        success_count += 1
                    elif status == "failed":
                        wh_line_parts.append(f"{COLOR_RED}✗{COLOR_RESET} {t}")
                        failed_count += 1
                    else:
                        wh_line_parts.append(f"{COLOR_GRAY}○{COLOR_RESET} {t}")
                        pending_count += 1
                
                prefix = f"Webhook[{wh_idx}]({wh_name}):" if wh_name else f"Webhook[{wh_idx}]:"
                all_lines.append(prefix)
                
                for i in range(0, len(wh_line_parts), 5):
                    chunk = wh_line_parts[i:i+5]
                    line = "         " + ", ".join(chunk)
                    all_lines.append(line)
            
            total_success += success_count
            total_failed += failed_count
            total_pending += pending_count
            
            total_count = len(all_times)
            progress = success_count + failed_count
            progress_percent = int(progress / total_count * 100) if total_count > 0 else 0
            
            progress_bar_length = 20
            progress_bar = "█" * int(progress_percent / (100 / progress_bar_length)) + "░" * (progress_bar_length - int(progress_percent / (100 / progress_bar_length)))
            
            progress_line = f"进度: [{progress_bar}] {progress}/{total_count} ({progress_percent}%)"
            
            manual_parts = []
            if manual_executions:
                manual_parts.append("手动执行:")
                
                for i in range(0, len(manual_executions), 5):
                    chunk = manual_executions[i:i+5]
                    line = "         " + ", ".join(chunk)
                    manual_parts.append(line)
            
            all_lines = [progress_line] + all_lines + manual_parts
            
            box_width = 80
            content_width = box_width - 4
            
            print(f"[{idx}] {task_name}")
            print("    ┌" + "─" * content_width + "┐")
            
            for line in all_lines:
                display_len = get_display_length(line)
                if display_len <= content_width:
                    padding = " " * (content_width - display_len)
                    print(f"    │{line}{padding}│")
                else:
                    current_line = line
                    current_display_len = display_len
                    while current_display_len > content_width:
                        split_idx = content_width
                        part_line = current_line[:split_idx]
                        part_display_len = get_display_length(part_line)
                        while part_display_len > content_width:
                            split_idx -= 1
                            part_line = current_line[:split_idx]
                            part_display_len = get_display_length(part_line)
                        padding = " " * (content_width - part_display_len)
                        print(f"    │{part_line}{padding}│")
                        current_line = current_line[split_idx:]
                        current_display_len = get_display_length(current_line)
                    if current_display_len > 0:
                        padding = " " * (content_width - current_display_len)
                        print(f"    │{current_line}{padding}│")
            
            print("    └" + "─" * content_width + "┘")
            print()
        
        print(f"{COLOR_BLUE}═══════════════════════════════════════════════════════════════════════════════════{COLOR_RESET}")
        
        total_executed = total_success + total_failed
        total_all = total_executed + total_pending
        overall_rate = (total_executed / total_all * 100) if total_all > 0 else 0
        
        print(f"总任务: {len(tasks)}  │  总时间点: {total_all}  │  {COLOR_GREEN}✓{COLOR_RESET} 成功: {total_success}  │  {COLOR_RED}✗{COLOR_RESET} 失败: {total_failed}  │  {COLOR_GRAY}○{COLOR_RESET} 未执行: {total_pending}  │  执行率: {overall_rate:.1f}%")
        
    except Exception as e:
        logger.error(f"读取任务列表失败：{str(e)}")

if __name__ == "__main__":
    main()
