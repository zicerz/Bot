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
from PIL import Image
import io
import threading

try:
    # UI Automation (UIA) for blocking Excel dialogs
    from pywinauto import Desktop
except Exception:
    Desktop = None

# ---------------------------- 日志配置 ----------------------------
def setup_logging():
    """配置日志系统：按年/月分级目录，每日独立日志文件"""
    # 获取当前脚本所在目录
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 创建日志目录结构：logs/年/月/
    now = datetime.now()
    log_dir = os.path.join(base_dir, "logs", str(now.year), f"{now.month:02d}")
    os.makedirs(log_dir, exist_ok=True)
    
    # 生成日志文件名：年-月-日.log
    log_filename = now.strftime("%Y-%m-%d.log")
    log_path = os.path.join(log_dir, log_filename)
    
    # 配置日志格式
    log_format = "%(asctime)s [%(levelname)s] %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    
    # 创建处理器：控制台输出 + 文件输出（追加模式）
    handlers = [
        logging.StreamHandler(),  # 控制台输出
        logging.FileHandler(log_path, mode='a', encoding='utf-8')  # 文件输出（追加模式）
    ]
    
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        datefmt=date_format,
        handlers=handlers,
        force=True  # 强制重新配置，避免重复配置问题
    )
    
    logger = logging.getLogger("ExcelBot")
    logger.info(f"日志系统已初始化，日志文件：{log_path}")
    return logger

logger = setup_logging()

# ---------------------------- Excel 处理器 ----------------------------
class ExcelProcessor:
    """Excel 操作引擎"""
    
    def __init__(self, file_path: str, visible=True):
        """
        初始化处理器
        :param file_path: Excel 文件绝对路径
        :param visible: 是否显示 Excel 界面（调试用）
        """
    
        self.file_path = os.path.abspath(file_path)
        self.visible = True  # 可视化调试模式
        self.excel = None
        self.workbook = None
        self._refresh_timeout = 500  # 数据刷新超时时间（秒）
        self._dialog_watchdog_stop = threading.Event()

    def __enter__(self):
        """安全启动 Excel 实例"""
        try:
            self.excel = win32.Dispatch("Excel.Application")
            self.excel.Visible = self.visible
            self.excel.DisplayAlerts = False
            self.workbook = self.excel.Workbooks.Open(self.file_path)
            # 自动设置所有工作表的缩放比例为220%
            for sheet in self.workbook.Worksheets:
                try:
                    sheet.Activate()
                    sheet.Application.ActiveWindow.Zoom = 220
                except Exception as e:
                    logger.debug(f"设置缩放失败：{str(e)}")
            logger.debug(f"成功打开文件：{os.path.basename(self.file_path)}")
            return self
        except Exception as e:
            self._safe_shutdown()
            raise RuntimeError(f"Excel 启动失败：{str(e)}")

    def __exit__(self, exc_type, exc_val, exc_tb):
        """确保资源释放"""
        self._safe_shutdown()

    def _start_dialog_watchdog(self, timeout_s: float = 90.0):
        """
        启动 UIA 弹窗守护线程，避免 COM 调用被模态弹窗阻塞。
        目前覆盖：共享冲突弹窗“其他人也在更改” → 点击“查看所有人的内容(E)”
        """
        if Desktop is None:
            logger.debug("未安装 pywinauto，跳过 UIA 弹窗守护")
            return

        self._dialog_watchdog_stop.clear()

        def _run():
            end_at = time.time() + float(timeout_s)
            while not self._dialog_watchdog_stop.is_set() and time.time() < end_at:
                try:
                    self._dismiss_other_people_editing_dialog()
                except Exception:
                    pass
                time.sleep(0.25)

        threading.Thread(target=_run, daemon=True).start()

    def _stop_dialog_watchdog(self):
        try:
            self._dialog_watchdog_stop.set()
        except Exception:
            pass

    def _dismiss_other_people_editing_dialog(self) -> bool:
        """
        处理弹窗：
        标题：其他人也在更改
        按钮：查看所有人的内容(E) / 仅查看我的内容(M)
        """
        if Desktop is None:
            return False

        dlg_title = "其他人也在更改"
        try:
            windows = Desktop(backend="uia").windows(title=dlg_title, control_type="Window", visible_only=True)
        except Exception:
            return False

        for w in windows:
            try:
                dlg = w
                # 优先按“按钮文本前缀”查找，避免热键括号差异导致匹配失败
                btn = dlg.child_window(title_re=r"^查看所有人的内容.*", control_type="Button")
                if btn.exists(timeout=0.1):
                    try:
                        btn.invoke()
                    except Exception:
                        btn.click_input()
                    logger.info("检测到弹窗“其他人也在更改”，已自动选择【查看所有人的内容】")
                    return True
            except Exception:
                continue
        return False


    def _safe_shutdown(self):
        """安全关闭 Excel 进程"""
        try:
            self._stop_dialog_watchdog()
            if self.workbook:
                self.workbook.Close(SaveChanges=True)
            if self.excel:
                self.excel.Quit()
            logger.debug("Excel 进程已释放")
        except Exception as e:
            logger.warning(f"资源释放异常：{str(e)}")

    def refresh_data(self) -> bool:
        """带超时检测的数据刷新"""
        logger.info("开始刷新数据...")
        start_time = time.time()
        
        
        try:
            # 启动弹窗守护，防止共享冲突弹窗导致阻塞
            self._start_dialog_watchdog(timeout_s=self._refresh_timeout + 120)

            # 刷新之前，判断哪些表是链接了数据源的表格
            linked_tables = []
            for sheet in self.workbook.Worksheets:
                try:
                    # 检查工作表中的表格（ListObjects）
                    list_objects = sheet.ListObjects
                    for i in range(1, list_objects.Count + 1):
                        table = list_objects.Item(i)
                        try:
                            # 检查表格是否有查询表（QueryTable），这表示链接了外部数据源
                            if hasattr(table, 'QueryTable') and table.QueryTable is not None:
                                # 获取表格的范围
                                table_range = table.Range.Address if hasattr(table, 'Range') else "未知"
                                linked_tables.append({
                                    'sheet': sheet.Name,
                                    'table': table.Name,
                                    'range': table_range,
                                    'table_obj': table,
                                    'query_table': table.QueryTable
                                })
                                logger.debug(f"发现链接数据源的表格：工作表 [{sheet.Name}] - 查询 [{table.Name}] - 范围 [{table_range}]")
                        except Exception as e:
                            # logger.debug(f"检查工作表 [{sheet.Name}] 的表格 [{table.Name if hasattr(table, 'Name') else 'Unknown'}] 时出错：{e}")
                            logger.debug(f"工作表 [{sheet.Name}] 的表格 [{table.Name if hasattr(table, 'Name') else 'Unknown'}] 未连接数据源")
                except Exception as e:
                    logger.debug(f"检查工作表 [{sheet.Name}] 时出错：{e}")
            
            if linked_tables:
                logger.info(f"共发现 {len(linked_tables)} 个链接了数据源的表格")
                for item in linked_tables:
                    logger.info(f"  - 工作表 [{item['sheet']}] - 查询 [{item['table']}] - 范围 [{item['range']}]")
            else:
                logger.info("未发现链接了数据源的表格")

            # 在发现了数据源的表格中查询是否设置了全部刷新时刷新此连接
            refresh_tables = []
            if linked_tables:
                logger.info("检查数据连接属性：")
                for item in linked_tables:
                    try:
                        query_table = item['query_table']
                        
                        if query_table is not None:
                            # 检查各种刷新属性
                            will_refresh_on_refresh_all = False
                            # 通过QueryTable直接访问WorkbookConnection
                            try:
                                # 通过QueryTable的WorkbookConnection属性
                                workbook_conn = query_table.WorkbookConnection
                                if workbook_conn:
                                    # 检查RefreshWithRefreshAll属性（如果存在）
                                    try:
                                        if hasattr(workbook_conn, 'RefreshWithRefreshAll'):
                                            will_refresh_on_refresh_all = workbook_conn.RefreshWithRefreshAll
                                    except:
                                        pass
                            except Exception as conn_e:
                                logger.debug(f"检查连接对象时出错: {conn_e}")
                            
                            status = "✓ 已设置" if will_refresh_on_refresh_all else "✗ 未设置"
                            logger.info(f"  工作表 [{item['sheet']}] - 查询 [{item['table']}]:")
                            logger.info(f"    - 全部刷新时刷新此连接: {status}")
                            
                            # 如果设置了全部刷新时刷新此连接，打印表格范围
                            if will_refresh_on_refresh_all:
                                logger.info(f"    - 数据源表格范围: {item['range']}")
                                # 添加到需要设置单元格值的工作表列表
                                refresh_tables.append(item)
                
                    except Exception as e:
                        logger.warning(f"检查表格 [{item['sheet']}] - [{item['table']}] 的连接属性时出错: {e}")
            
            # 将每个链接了数据源并设置了全部刷新时刷新此连接的表格所在工作表的左上角单元格值设置为1
            if refresh_tables:
                logger.info(f"在 {len(refresh_tables)} 个工作表中设置左上角单元格值为1")
                for item in refresh_tables:
                    # print('--------------------------------------------------------------------------------------------')
                    # print(item['range'])
                    #将range转换为左上角单元格
                    range_start = item['range'].split(':')[0]
                    # print(range_start)
                    try:
                        sheet = self.workbook.Worksheets(item['sheet'])
                        # 设置表格范围的左上角单元格值为1
                        sheet.Range(range_start).Value = 1
                        logger.info(f"  工作表 [{item['sheet']}] - 已将 {range_start} 单元格值设置为 1")
                    except Exception as e:
                        logger.warning(f"设置工作表 [{item['sheet']}] 的 {range_start} 单元格值时出错: {e}")

            time.sleep(10)
            
            # 执行刷新并验证，最多重试3次
            max_retries = 3
            failed_tables = []  # 保存失败的表格列表
            
            for retry_count in range(max_retries + 1):
                if retry_count == 0:
                    # 第一次：执行全部刷新
                    logger.info("执行全部刷新...")
                    self.workbook.RefreshAll()
                    self.excel.CalculateUntilAsyncQueriesDone()
                else:
                    # 重试：只刷新失败的表格
                    if not failed_tables:
                        # 如果没有失败的表格，说明已经全部成功，退出循环
                        break
                    
                    logger.warning(f"第 {retry_count} 次重试刷新，发现 {len(failed_tables)} 个表格刷新失败")
                    for item in failed_tables:
                        logger.warning(f"  重试刷新：工作表 [{item['sheet']}] - 查询 [{item['table']}]")
                        try:
                            
                            # 刷新单个表格
                            if item['query_table']:
                                item['query_table'].Refresh()
                        except Exception as e:
                            logger.error(f"重试刷新表格 [{item['sheet']}] - [{item['table']}] 时出错: {e}")
                    
                    self.excel.CalculateUntilAsyncQueriesDone()
                
                # 轮询检查计算状态
                calculation_timeout = 300  # 5分钟超时
                calculation_start = time.time()
                while time.time() - calculation_start < calculation_timeout:
                    if self.excel.CalculationState == 0:  # 0 表示计算完成
                        break
                    time.sleep(5)
                else:
                    logger.warning("计算状态检查超时，继续验证单元格值")
                
                # 检查刷新结果：验证所有表格的左上角单元格值
                if refresh_tables:
                    failed_tables = []  # 重置失败列表
                    for item in refresh_tables:
                        range_start = item['range'].split(':')[0]
                        try:
                            sheet = self.workbook.Worksheets(item['sheet'])
                            cell_value = sheet.Range(range_start).Value
                            print(item['sheet'])
                            print(item['table'])
                            print(range_start)
                            print(cell_value)
                            # 修正判断逻辑：单元格值为1表示刷新失败
                            # 考虑不同数据类型的情况，将单元格值转为字符串后与'1'比较，以避免由于数值/文本/其他类型造成的判断失误
                            if str(cell_value).strip() != '1':
                                logger.info(f"工作表 [{item['sheet']}] - 查询 [{item['table']}] 的 {range_start} 单元格值已更新，刷新成功")
                            else:
                                failed_tables.append(item)
                                logger.warning(f"工作表 [{item['sheet']}] - 查询 [{item['table']}] 的 {range_start} 单元格值仍为1，刷新失败")
                        except Exception as e:
                            logger.warning(f"检查工作表 [{item['sheet']}] 的 {range_start} 单元格值时出错: {e}")
                            failed_tables.append(item)
                    
                    if not failed_tables:
                        logger.info("所有表格刷新成功！")
                        # 刷新后重新应用所有表格的筛选和排序，并刷新所有的数据透视表
                        # 重新应用筛选时最容易触发“其他人也在更改”弹窗，提前再起一轮短守护
                        self._start_dialog_watchdog(timeout_s=90)
                        for sheet in self.workbook.Worksheets:
                            try:
                                if sheet.AutoFilter is not None:
                                    # 重新应用筛选
                                    sheet.AutoFilter.ApplyFilter()
                                    logger.debug(f"重新应用筛选：{sheet.Name}")
                            except Exception as e:
                                logger.debug(f"应用筛选/排序失败：{sheet.Name} - {e}")
                        # 刷新所有的数据透视表
                        self._start_dialog_watchdog(timeout_s=180)
                        for sheet in self.workbook.Worksheets:
                            try:
                                # 遍历工作表中的所有数据透视表（PivotTables 为集合）
                                if hasattr(sheet, "PivotTables"):
                                    for i in range(1, sheet.PivotTables().Count + 1):
                                        pt = sheet.PivotTables(i)
                                        pt.RefreshTable()
                                        logger.debug(f"刷新数据透视表：{sheet.Name} - {pt.Name}")
                                        time.sleep(20)
                            except Exception as e:
                                logger.debug(f"刷新数据透视表失败：{sheet.Name} - {e}")
                               


                                
                        return True
                    else:
                        if retry_count < max_retries:
                            logger.warning(f"发现 {len(failed_tables)} 个表格刷新失败，准备重试（剩余 {max_retries - retry_count} 次）")
                            time.sleep(5)  # 等待一段时间后重试
                        else:
                            # 达到最大重试次数，仍有失败的表格
                            logger.error(f"达到最大重试次数（{max_retries}次），仍有 {len(failed_tables)} 个表格刷新失败")
                            for item in failed_tables:
                                logger.error(f"  失败表格：工作表 [{item['sheet']}] - 查询 [{item['table']}]")
                            return False
                else:
                    # 没有需要验证的表格，直接返回成功
                    logger.info("没有需要验证的表格，刷新完成")
                    # 刷新后重新应用所有表格的筛选和排序
                    self._start_dialog_watchdog(timeout_s=90)
                    for sheet in self.workbook.Worksheets:
                        try:
                            if sheet.AutoFilter is not None:
                                # 重新应用筛选
                                sheet.AutoFilter.ApplyFilter()
                                logger.debug(f"重新应用筛选：{sheet.Name}")
                        except Exception as e:
                            logger.debug(f"应用筛选/排序失败：{sheet.Name} - {e}")
                    return True
            
            # 如果循环正常结束（理论上不应该到这里，因为所有情况都应该有return）
            logger.warning("刷新循环异常结束")
            return False
        except Exception as e:
            logger.error(f"刷新异常：{str(e)}")
            return False
        finally:
            # 尽量停止守护线程（daemon 不会阻塞退出，这里只是减少无意义轮询）
            self._stop_dialog_watchdog()

    def validate_date(self, check_sheet, check_range, check_frequency) -> bool:
        """带重试的数据校验"""
        for attempt in range(1, check_frequency+1):
            try:
                logger.debug(f"校验数据：工作表 [{check_sheet}] 区域 [{check_range}]")
                sheet = self.workbook.Worksheets(check_sheet)
                valid = sheet.Range(check_range).Value != 0
                logger.info(f"数据校验 {'通过' if valid else '失败'}（第 {attempt} 次尝试）共{check_frequency}次")
        
                if valid:
                    return True
                if attempt < check_frequency:
                    time.sleep(10)  # 重试间隔
                    # 重新刷新数据
                    self.refresh_data()
            except Exception as e:
                logger.error(f"校验异常：{str(e)}")
        return False

    def capture_screenshots(self, configs: list) -> list:
        """批量截图（自动清理临时图表）"""
        screenshots = []
        for cfg in configs:
            try:
                sheet = self.workbook.Worksheets(cfg["sheet_name"])
                output_path = self._generate_path(cfg["name"])
                
                if self._capture_range(sheet, cfg["range"], output_path):
                    screenshots.append(output_path)
                    logger.debug(f"生成截图：{os.path.basename(output_path)}")
            except Exception as e:
                logger.error(f"截图失败 [{cfg['name']}]：{str(e)}")


        # 截图完成后，将所有工作表缩放比例恢复为100%
        try:
            for sheet in self.workbook.Worksheets:
                sheet.Activate()
                sheet.Application.ActiveWindow.Zoom = 100
            logger.debug("已将所有工作表缩放比例恢复为100%")
        except Exception as e:
            logger.warning(f"恢复缩放比例失败：{str(e)}")
       

        return screenshots
    

    def _capture_range(self, sheet, range_addr: str, output_path: str) -> bool:
        """执行区域截图"""
        try:
            if ":" in range_addr:
                range_obj = sheet.Range(range_addr)
            else:
                start_cell = sheet.Range(range_addr.split(":")[0])
                range_obj = start_cell.CurrentRegion

            logger.debug(f"截图区域地址: {range_obj.Address}")
            # try:
            #     val = range_obj.Value
            #     logger.debug(f"截图区域首行首列值: {val[0][0] if isinstance(val, tuple) else val}")
            # except Exception as e:
            #     logger.debug(f"无法获取区域值: {e}")

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
                logger.error(f"Paste异常：{str(e)}", exc_info=True)
                chart_obj.Delete()
                return False
            chart.Export(output_path)
            chart_obj.Delete()
            return os.path.exists(output_path)
        except Exception as e:
            logger.error(f"截图异常：{str(e)}", exc_info=True)
            return False

    def _generate_path(self, prefix: str) -> str:
        """生成唯一文件名"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")  # 加微秒
        # 加入excel文件名或任务名，防止同名
        task_tag = os.path.splitext(os.path.basename(self.file_path))[0]
        return os.path.join(
            os.path.dirname(self.file_path),
            f"{task_tag}_{prefix}_{timestamp}.png"
        )

# ---------------------------- 任务处理器 ----------------------------
class ReportTask:
    """报表任务实例"""
    
    def __init__(self, config: dict):
        self.config = self._validate_config(config)
        self.retry_limit = 3  # 微信发送重试次数

    def _validate_config(self, config: dict) -> dict:
        """配置完整性检查"""
        required_fields = ["excel_path", "schedule", "capture_configs"]
        missing = [f for f in required_fields if f not in config]
        if missing:
            raise ValueError(f"缺失必要配置：{missing}")

        # 检查 schedule 下的 times 和 webhook
        schedule = config["schedule"]
        if "times" not in schedule or "webhook" not in schedule:
            raise ValueError("缺失必要配置：['schedule.times', 'schedule.webhook']")

        if not os.path.exists(config["excel_path"]):
            raise FileNotFoundError(config["excel_path"])
            
        return config

    def execute(self, debug_mode=False):
        """执行任务流程"""
        # 任务开始分割线
        separator = "=" * 100
        logger.info(separator)
        logger.info(f"启动任务：{os.path.basename(self.config['excel_path'])}")
        logger.info(separator)
        start_time = time.time()
        
        try:
            with ExcelProcessor(
                self.config["excel_path"], 
                visible=debug_mode
            ) as excel:
                # 核心流程

                # 刷新数据
                if not excel.refresh_data():
                    logger.warning("数据刷新失败，发送通知并终止任务")
                    # 发送异常通知
                    self._send_wechat(
                        type="text",
                        data={
                            "content": f"数据刷新失败（超时或重试3次后仍有表格未刷新成功），请检查文件：{os.path.basename(self.config['excel_path'])}",
                            "mentioned_list": ["zhufuzhe"]
                        },
                        description="数据刷新失败通知",
                        webhook="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=833b098e-d8b8-43ea-bfdf-cade0d040fb6"
                    )
                    return

                # 数据校验
                if self.config.get("data_check_enable", False):
                    check_sheet = self.config["data_check"]["check_sheet"]
                    check_range = self.config["data_check"]["check_range"]
                    check_frequency = self.config["data_check"]["check_frequency"]
                    if not excel.validate_date(check_sheet, check_range, check_frequency):
                        logger.warning("数据校验未通过，发送通知并终止任务")
                        #发送异常通知
                        self._send_wechat(
                            type="text",
                            data={"content": self.config["data_check"]["notify_message"], 
                                "mentioned_list": self.config["data_check"]["notify_users"]
                            },
                            description="数据校验失败通知",
                            webhook = self.config["data_check"]["warning_webhook"]
                        )
                        return
                
                screenshots = excel.capture_screenshots(self.config["capture_configs"])
            self._deliver_results(screenshots)

        except Exception as e:
            logger.error(f"任务异常：{str(e)}", exc_info=debug_mode)
        finally:
            elapsed_time = time.time() - start_time
            logger.info(f"任务耗时：{elapsed_time:.2f}s")
            # 任务结束分割线
            separator = "=" * 100
            logger.info(separator)
            print(separator)  # 同时输出到控制台

    def _deliver_results(self, screenshots: list):
        """结果交付（图片+文件）"""

        # 发送截图
        for img_path in screenshots:
            self._send_wechat(
                type="image",
                data=self._prepare_image(img_path),
                description=f"截图 {os.path.basename(img_path)}",
                webhook = self.config["schedule"]["webhook"]
            )
        
        # 发送文件
        if self.config.get("send_file_enable", False):
            self._send_attachment()
        # 清理临时文件
        self._cleanup(screenshots)

    def _send_attachment(self):
        """发送关联文件"""
        file_path = self.config.get("file_path")
        if not file_path or not os.path.exists(file_path):
            logger.warning("无效的文件路径，跳过发送")
            return
        try:
            with open(file_path, "rb") as f:
                media_id = self._upload_file(f)
                if media_id:
                    self._send_wechat(
                        type="file",
                        data={"media_id": media_id},
                        description=f"文件 {os.path.basename(file_path)}",
                        webhook = self.config["schedule"]["webhook"]
                    )
        except Exception as e:
            logger.error(f"文件发送失败：{str(e)}")

    def _upload_file(self, file_obj) -> str:
        """上传文件到临时素材"""
        try:
            logger.debug(f"正在上传文件：{file_obj.name}")
            #文件路径改为文件名
            filename = os.path.basename(file_obj.name)
            name, ext = os.path.splitext(filename)
            filename_with_time = f"{name}_{datetime.now().strftime('%Y-%m-%d')}{ext}"
            
            # 上传文件
            response = requests.post(
                self.config["upload_url"],
                files={"media": (filename_with_time, file_obj)},
                timeout=15
            )
            response.raise_for_status()
            return response.json().get("media_id")
        except Exception as e:
            logger.error(f"文件上传异常：{str(e)}")
            return None

    def _prepare_image(self, img_path: str) -> dict:
        """准备图片数据"""
        max_size = 2 * 1024 * 1024  # 2MB
        min_width = 800  # 最小宽度，防止图片太小
        min_height = 600 # 最小高度

        with open(img_path, "rb") as f:
            img_data = f.read()
            if len(img_data) > max_size:
                img = Image.open(io.BytesIO(img_data))
                img = img.convert("RGB")  # 保证兼容性
                buf = io.BytesIO()
                quality = 85

                # 先尝试只压缩质量
                while True:
                    buf.seek(0)
                    img.save(buf, format="JPEG", quality=quality)
                    if buf.tell() <= max_size or quality <= 60:
                        break
                    quality -= 5

                # 如果还超出2M，再缩放尺寸
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

    def _send_wechat(self, type: str, data: dict, description: str, webhook):
        """发送到企业微信（带重试）"""
        payload = {"msgtype": type, type: data}
        
        for attempt in range(1, self.retry_limit+1):
            try:
                response = requests.post(
                    webhook,
                    json=payload,
                    timeout=10
                )
                response.raise_for_status()
                logger.info(f"发送成功：{description}")
                return
            except Exception as e:
                logger.warning(f"发送失败（{attempt}/{self.retry_limit}）：{description}")
                if attempt == self.retry_limit:
                    logger.error(f"最终发送失败：{str(e)}")
                time.sleep(2 ** attempt)

    def _cleanup(self, files: list):
        """清理临时文件"""
        for f in files:
            try:
                os.remove(f)
                logger.debug(f"清理临时文件：{os.path.basename(f)}")
            except Exception as e:
                logger.warning(f"文件清理失败：{str(e)}")

# ---------------------------- 任务调度器 ----------------------------
class TaskScheduler:
    """多任务调度引擎"""
    
    def __init__(self, config_path: str, debug=False):
        self.tasks = self._load_tasks(config_path)
        self.debug_mode = debug
        logger.setLevel(logging.DEBUG if debug else logging.INFO)

    def _load_tasks(self, config_path: str) -> list:
        """加载配置文件"""
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
                
            if not isinstance(config.get("tasks"), list):
                raise ValueError("配置文件格式错误")
                
            logger.info(f"成功加载 {len(config['tasks'])} 个任务")
            return [ReportTask(task) for task in config["tasks"]]
        except Exception as e:
            logger.error(f"配置加载失败：{str(e)}")
            raise

    def start(self):
        """启动调度服务"""
        logger.info("启动任务调度器...")
        self._schedule_tasks()
        
        try:
            while True:
                schedule.run_pending()
                time.sleep(1)
        except KeyboardInterrupt:
            logger.info("正在关闭调度器...")

    def _schedule_tasks(self):
        """配置定时任务"""
        for task in self.tasks:
            for trigger_time in task.config["schedule"]["times"]:
                schedule.every().day.at(trigger_time).do(
                    self._run_task, task
                )
                logger.info(f"已安排任务：{trigger_time} → {os.path.basename(task.config['excel_path'])}")

    def _run_task(self, task: ReportTask):
        """串行执行任务"""
        pythoncom.CoInitialize()
        try:
            # 定时任务执行时添加分割线
            separator = "=" * 100
            logger.info("")
            logger.info(f"定时任务触发 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info(separator)
            task.execute(self.debug_mode)
        finally:
            pythoncom.CoUninitialize()

    def run_now(self, task_id: int = None):
        """立即执行任务（调试）"""
        logger.info("进入调试模式...")
        targets = self.tasks if task_id is None else [self.tasks[task_id]]
        
        for idx, task in enumerate(targets, 1):
            try:
                # 多个任务之间添加分割线
                if len(targets) > 1:
                    separator = "=" * 100
                    logger.info("")
                    logger.info(f"任务 {idx}/{len(targets)}")
                    logger.info(separator)
                task.execute(self.debug_mode)
            except Exception as e:
                logger.error(f"执行异常：{str(e)}")

# ---------------------------- 主程序 ----------------------------
def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description="Excel 自动化")
    parser.add_argument("--run-all", action="store_true", help="立即执行所有任务")
    parser.add_argument("--task", type=int, help="执行指定序号的任务")
    parser.add_argument("--debug", action="store_true", help="开启调试模式")
    args = parser.parse_args()

    try:
        scheduler = TaskScheduler("config.yml", debug=args.debug)
        
        if args.run_all or args.task is not None:
            scheduler.run_now(args.task)
        else:
            scheduler.start()
    except Exception as e:
        logger.error(f"系统异常：{str(e)}", exc_info=args.debug)
        exit(1)
if __name__ == "__main__":
    main()