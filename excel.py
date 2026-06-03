import win32com.client as win32
import yaml
import os
import time
import schedule
import requests
import base64
import hashlib
from datetime import datetime
import logging
import argparse
import pythoncom
import io
import threading

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
def setup_logging():
    """配置日志系统：按年/月分级目录，每日独立日志文件"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    now = datetime.now()
    log_dir = os.path.join(base_dir, "logs", str(now.year), f"{now.month:02d}")
    os.makedirs(log_dir, exist_ok=True)
    
    log_filename = now.strftime("%Y-%m-%d.log")
    log_path = os.path.join(log_dir, log_filename)
    
    log_format = "%(asctime)s [%(levelname)s] [任务%(task_id)s] %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    
    formatter = logging.Formatter(log_format, datefmt=date_format)
    
    handlers = [
        logging.StreamHandler(),
        logging.FileHandler(log_path, mode='a', encoding='utf-8')
    ]
    
    for handler in handlers:
        handler.setFormatter(formatter)
    
    logging.basicConfig(
        level=logging.INFO,
        handlers=handlers,
        force=True
    )
    
    logger = logging.getLogger("ExcelBot")
    logger = logging.LoggerAdapter(logger, {"task_id": "-"})
    logger.info(f"日志系统已初始化，日志文件：{log_path}")
    return logger

logger = setup_logging()

def get_task_logger(task_id: int):
    """获取带任务序号的日志记录器"""
    base_logger = logging.getLogger("ExcelBot")
    return logging.LoggerAdapter(base_logger, {"task_id": task_id})

# ---------------------------- 文件锁工具类 ----------------------------
class FileLock:
    """文件锁工具类，用于防止并发访问"""
    
    def __init__(self, lock_file_path):
        self.lock_file_path = lock_file_path
        self.lock_file = None
    
    def acquire(self, timeout=300, poll_interval=2):
        """获取文件锁
        
        :param timeout: 最大等待时间（秒），默认300秒
        :param poll_interval: 轮询间隔（秒），默认2秒
        :return: True表示获取锁成功，False表示超时
        """
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                self.lock_file = open(self.lock_file_path, 'w')
                portalocker.lock(self.lock_file, portalocker.LOCK_EX | portalocker.LOCK_NB)
                logger.info(f"成功获取文件锁：{self.lock_file_path}")
                return True
            except (IOError, BlockingIOError, portalocker.LockException):
                if self.lock_file:
                    try:
                        self.lock_file.close()
                    except Exception:
                        pass
                    self.lock_file = None
                remaining = int(timeout - (time.time() - start_time))
                logger.info(f"文件锁被占用，等待中...（剩余 {remaining} 秒）")
                time.sleep(poll_interval)
        
        logger.error(f"获取文件锁超时（{timeout}秒）：{self.lock_file_path}")
        return False
    
    def release(self):
        """释放文件锁"""
        if self.lock_file:
            try:
                portalocker.unlock(self.lock_file)
                self.lock_file.close()
                logger.info(f"成功释放文件锁：{self.lock_file_path}")
            except Exception as e:
                logger.warning(f"释放文件锁异常：{str(e)}")
            finally:
                self.lock_file = None
    
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
        lock_file_path = self.file_path + ".lock"
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
                
                time.sleep(1)
                
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
        screenshots = []
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
                    try:
                        sheet = self.workbook.Worksheets(cfg["sheet_name"])
                        output_path = self._generate_path(cfg["name"])

                        if self._capture_range(sheet, cfg["range"], output_path):
                            screenshots.append(output_path)
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

    def __init__(self, config: dict, test_webhook: str = None, error_webhook: str = None, upload_url_template: str = None, task_id: int = 0):
        self.config = self._validate_config(config)
        self.retry_limit = 3
        self.error_webhook = error_webhook or "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=833b098e-d8b8-43ea-bfdf-cade0d040fb6"
        self.test_webhook = test_webhook
        self.upload_url_template = upload_url_template
        self.task_id = task_id
        self.logger = get_task_logger(task_id)

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

    def execute(self, debug_mode=False, webhook_configs=None):
        """
        执行任务流程
        :param debug_mode: 是否调试模式
        :param webhook_configs: 特定的webhook配置（单个dict或列表，None表示执行所有webhook）
        """
        separator = "=" * 100
        self.logger.info(separator)
        self.logger.info(f"启动任务：{os.path.basename(self.config['excel_path'])}")
        
        if webhook_configs:
            if isinstance(webhook_configs, list):
                webhook_keys = [wh["webhook"].split("key=")[-1][:8] + "..." for wh in webhook_configs]
                self.logger.info(f"Webhooks：{', '.join(webhook_keys)}")
            else:
                webhook_key = webhook_configs["webhook"].split("key=")[-1][:8] + "..."
                self.logger.info(f"Webhook：{webhook_key}")
        self.logger.info(separator)
        
        start_time = time.time()
        results_to_deliver = []
        
        try:
            with ExcelProcessor(
                self.config["excel_path"], 
                visible=debug_mode,
                task_logger=self.logger
            ) as excel:
                # 刷新数据（所有webhook共享一次刷新）
                if not excel.refresh_data():
                    task_id_str = self._get_task_id_str()
                    self.logger.warning(f"{task_id_str}数据刷新失败，发送通知并终止任务")
                    self._send_wechat(
                        type="text",
                        data={
                            "content": f"{task_id_str}数据刷新失败（超时或重试3次后仍有表格未刷新成功），请检查文件：{os.path.basename(self.config['excel_path'])}",
                            "mentioned_list": ["zhufuzhe"]
                        },
                        description="数据刷新失败通知",
                        webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=833b098e-d8b8-43ea-bfdf-cade0d040fb6"
                    )
                    return

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
                        return
                
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
                for wh_config, wh_idx in target_webhooks_with_indices:
                    # 获取带主任务-子任务序号的logger
                    if wh_idx == -1:
                        wh_logger = self.logger
                    else:
                        wh_logger = get_task_logger(f"{self.task_id}-{wh_idx}")
                    
                    wh_logger.info(f"处理Webhook：{wh_config['webhook'].split('key=')[-1][:8]}...")
                    
                    # 更新excel的logger为带子任务序号的logger
                    excel.logger = wh_logger
                    
                    screenshots, failed_capture_configs = excel.capture_screenshots(
                        wh_config["capture_configs"],
                        retry_times=3
                    )

                    if failed_capture_configs:
                        failed_regions_text = "；".join(
                            [
                                f"{item.get('name', '未命名')}({item.get('sheet_name', '未知工作表')}:{item.get('range', '未知区域')})"
                                for item in failed_capture_configs
                            ]
                        )
                        task_id_str = self._get_task_id_str(wh_idx)
                        wh_logger.error(
                            f"{task_id_str}截图在重试 3 次后仍失败，共 {len(failed_capture_configs)} 个区域：{failed_regions_text}"
                        )

                        if screenshots:
                            self._cleanup(screenshots, wh_logger)
                        self._send_wechat(
                            type="text",
                            data={
                                "content": (
                                    f"{task_id_str}截图失败：重试3次后仍有 {len(failed_capture_configs)} 个区域未成功截图，"
                                    f"任务已终止。文件：{os.path.basename(self.config['excel_path'])}。"
                                    f"失败区域：{failed_regions_text}"
                                )
                            },
                            description="截图失败通知",
                            webhook=wh_config["webhook"],
                            task_logger=wh_logger
                        )
                        continue

                    results_to_deliver.append({
                        "screenshots": screenshots,
                        "webhook_config": wh_config,
                        "wh_idx": wh_idx
                    })
                
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
                self._deliver_results(result["screenshots"], result["webhook_config"], wh_logger)

        except Exception as e:
            error_text = str(e)
            task_id_str = self._get_task_id_str()
            self.logger.error(f"{task_id_str}任务异常：{error_text}", exc_info=debug_mode)
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
            elapsed_time = time.time() - start_time
            self.logger.info(f"任务耗时：{elapsed_time:.2f}s")
            separator = "=" * 100
            self.logger.info(separator)
            print(separator)

    def _deliver_results(self, screenshots: list, webhook_config: dict, task_logger=None):
        """根据webhook配置交付结果"""
        if task_logger is None:
            task_logger = self.logger
        webhook = webhook_config["webhook"]
        send_file_enable = webhook_config.get("send_file_enable", 0)
        
        webhook_key = webhook.split("key=")[-1][:8]
        task_logger.info(f"_deliver_results 被调用")
        task_logger.info(f"  webhook: {webhook_key}...")
        task_logger.info(f"  screenshots数量: {len(screenshots)}")
        task_logger.info(f"  send_file_enable: {send_file_enable}")

        # 发送截图
        for img_path in screenshots:
            self._send_wechat(
                type="image",
                data=self._prepare_image(img_path),
                description=f"截图 {os.path.basename(img_path)}",
                webhook=webhook,
                task_logger=task_logger
            )

        # 发送文件
        if send_file_enable:
            task_logger.info("send_file_enable 为真，准备发送文件")
            self._send_attachment(webhook, task_logger)
        else:
            task_logger.info("send_file_enable 为假，跳过发送文件")
        
        # 清理临时文件
        self._cleanup(screenshots, task_logger)

    def _send_attachment(self, webhook: str, task_logger=None):
        """发送关联文件"""
        if task_logger is None:
            task_logger = self.logger
        file_path = self.config.get("file_path") or self.config.get("excel_path")
        task_logger.info(f"准备发送文件，file_path: {file_path}")
        task_logger.info(f"send_file_enable 检查，webhook: {webhook.split('key=')[-1][:8]}...")
        
        if not file_path:
            task_logger.warning("文件路径为空，跳过发送")
            return
        if not os.path.exists(file_path):
            task_logger.warning(f"文件不存在：{file_path}，跳过发送")
            return
        
        task_logger.info(f"开始上传文件：{os.path.basename(file_path)}")
        try:
            with open(file_path, "rb") as f:
                media_id = self._upload_file(f, webhook, task_logger)
                if media_id:
                    task_logger.info(f"文件上传成功，media_id: {media_id}")
                    self._send_wechat(
                        type="file",
                        data={"media_id": media_id},
                        description=f"文件 {os.path.basename(file_path)}",
                        webhook=webhook,
                        task_logger=task_logger
                    )
                else:
                    task_logger.warning("文件上传失败，未获取到media_id")
        except Exception as e:
            task_logger.error(f"文件发送异常：{str(e)}")

    def _upload_file(self, file_obj, webhook: str, task_logger=None) -> str:
        """上传文件到临时素材"""
        if task_logger is None:
            task_logger = self.logger
        try:
            upload_url = self._get_upload_url(webhook)
            task_logger.info(f"上传URL: {upload_url}")
            
            if not upload_url:
                task_logger.warning("无效的上传上传，跳过文件上传")
                return None
                
            task_logger.info(f"正在上传文件：{file_obj.name}")
            filename = os.path.basename(file_obj.name)
            name, ext = os.path.splitext(filename)
            filename_with_time = f"{name}_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}{ext}"
            task_logger.info(f"上传文件名：{filename_with_time}")
            
            response = requests.post(
                upload_url,
                files={"media": (filename_with_time, file_obj)},
                timeout=15
            )
            
            task_logger.info(f"上传响应状态码: {response.status_code}")
            task_logger.info(f"上传响应内容: {response.text}")
            
            response.raise_for_status()
            result = response.json()
            task_logger.info(f"上传结果: {result}")
            return result.get("media_id")
        except Exception as e:
            task_logger.error(f"文件上传异常：{str(e)}")
            return None

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
                task_logger.info(f"发送成功：{description}")
                return
            except Exception as e:
                task_logger.warning(f"发送失败（{attempt}/{self.retry_limit}）：{description}")
                if attempt == self.retry_limit:
                    task_logger.error(f"最终发送失败：{str(e)}")
                time.sleep(2 ** attempt)

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

# ---------------------------- 任务调度器 ----------------------------
class TaskScheduler:
    """多任务调度引擎"""

    def __init__(self, config_path: str, debug=False, test_webhook=None, error_webhook=None):
        self.tasks = self._load_tasks(config_path, test_webhook, error_webhook)
        self.debug_mode = debug
        self._scheduler_lock = None
        logger.setLevel(logging.DEBUG if debug else logging.INFO)

    def _load_tasks(self, config_path: str, test_webhook=None, error_webhook=None) -> list:
        """加载配置文件"""
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)

            if not isinstance(config.get("tasks"), list):
                raise ValueError("配置文件格式错误")

            if error_webhook is None and config.get("error_webhook"):
                error_webhook = config["error_webhook"]

            upload_url_template = config.get("upload_url_template", "")

            logger.info(f"成功加载 {len(config['tasks'])} 个任务")
            return [ReportTask(task, test_webhook, error_webhook, upload_url_template, idx) for idx, task in enumerate(config["tasks"])]
        except Exception as e:
            logger.error(f"配置加载失败：{str(e)}")
            raise

    def start(self):
        """启动调度服务"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        scheduler_lock_path = os.path.join(base_dir, "scheduler.lock")
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

    def _run_task(self, task: ReportTask, webhook_configs: list):
        """串行执行任务（支持多个webhook配置共享一次刷新）"""
        pythoncom.CoInitialize()
        try:
            separator = "=" * 100
            logger.info("")
            logger.info(f"定时任务触发 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info(separator)
            
            webhook_keys = [wh["webhook"].split("key=")[-1][:8] for wh in webhook_configs]
            logger.info(f"本次任务将发送到 {len(webhook_configs)} 个 webhook: {','.join(webhook_keys)}")
            
            task.execute(self.debug_mode, webhook_configs)
        finally:
            pythoncom.CoUninitialize()

    def run_now(self, task_specs: list = None):
        """立即执行任务（调试）
        :param task_specs: 任务规格列表，每个元素为 (task_id, webhook_id) 元组
        """
        logger.info("进入调试模式...")
        
        # 显示任务配置信息
        logger.info("当前任务配置：")
        for idx, task in enumerate(self.tasks):
            task_name = os.path.basename(task.config["excel_path"])
            logger.info(f"  [{idx}] {task_name}")
            for wh_idx, wh_config in enumerate(task.config["webhooks"]):
                wh_key = wh_config["webhook"].split("key=")[-1][:10]
                logger.info(f"    - [{wh_idx}] webhook: {wh_key}..., times: {wh_config['times']}")
        
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
            pythoncom.CoInitialize()
            try:
                if len(targets) > 1:
                    separator = "=" * 100
                    logger.info("")
                    logger.info(f"任务 {idx}/{len(targets)}")
                    logger.info(separator)
                
                # 确定要执行的webhook
                if webhook_id is not None:
                    webhook_config = task.config["webhooks"][webhook_id]
                    logger.info(f"执行任务 {task_specs[idx-1][0]} 的 webhook {webhook_id}: {webhook_config['webhook'].split('key=')[-1][:10]}...")
                    task.execute(self.debug_mode, webhook_config)
                else:
                    # 执行所有webhook配置
                    task_name = os.path.basename(task.config["excel_path"])
                    if task_specs:
                        logger.info(f"执行任务 {task_specs[idx-1][0]}: {task_name}")
                    task.execute(self.debug_mode)
            except Exception as e:
                logger.error(f"执行异常：{str(e)}")
            finally:
                pythoncom.CoUninitialize()

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

        if args.run_all or args.task is not None:
            scheduler.run_now(task_specs)
        else:
            scheduler.start()
    except Exception as e:
        logger.error(f"系统异常：{str(e)}", exc_info=args.debug)
        exit(1)

def print_task_list():
    """打印所有已配置的任务列表"""
    try:
        with open("config.yml", "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        
        tasks = config.get("tasks", [])
        print("已配置的任务列表：")
        print("-" * 100)
        
        for idx, task in enumerate(tasks):
            task_name = task.get("name", os.path.basename(task["excel_path"]))
            webhooks = task.get("webhooks", [])
            
            if webhooks:
                for wh_idx, wh_config in enumerate(webhooks):
                    times = wh_config.get("times", [])
                    times_str = str(times).replace("'", '"')
                    if wh_idx == 0:
                        print(f"[{idx}] {task_name}")
                        print(f"      webhook[{wh_idx}]: times={times_str}")
                    else:
                        print(f"      webhook[{wh_idx}]: times={times_str}")
            else:
                print(f"[{idx}] {task_name}")
                print("      (无webhook配置)")
        
        print("-" * 100)
        print(f"共 {len(tasks)} 个任务")
        
    except Exception as e:
        logger.error(f"读取任务列表失败：{str(e)}")

if __name__ == "__main__":
    main()
