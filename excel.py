"""
Excel 自动化报表系统
功能：多文件定时刷新、智能校验、区域截图、微信推送
特性：
1. 多任务独立配置，支持不同触发时间和接收群组
2. 智能重试机制（数据刷新+消息发送）
3. 可视化调试模式
4. 完善的异常处理和资源管理
"""

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
import threading
import argparse
import pythoncom
from PIL import Image
import io

# ---------------------------- 日志配置 ----------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("ExcelBot")

# ---------------------------- Excel 处理器 ----------------------------
class ExcelProcessor:
    """Excel 操作引擎"""
    
    def __init__(self, file_path: str, visible=False):
        """
        初始化处理器
        :param file_path: Excel 文件绝对路径
        :param visible: 是否显示 Excel 界面（调试用）
        """
    
        self.file_path = os.path.abspath(file_path)
        self.visible = True
        self.excel = None
        self.workbook = None
        self._refresh_timeout = 120  # 数据刷新超时时间（秒）

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
        """保证资源释放"""
        self._safe_shutdown()
    def _safe_shutdown(self):
        """安全关闭 Excel 进程"""
        try:
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
            self.workbook.RefreshAll()
            self.excel.CalculateUntilAsyncQueriesDone()
            
            # 轮询检查计算状态
            while time.time() - start_time < self._refresh_timeout:
                if self.excel.CalculationState == 0:  # 0 表示计算完成
                    logger.info(f"数据刷新完成（耗时 {time.time()-start_time:.1f}s）")


                    # 刷新后重新应用所有表格的筛选和排序
                    for sheet in self.workbook.Worksheets:
                        try:
                            if sheet.AutoFilter is not None:
                                # 重新应用筛选
                                sheet.AutoFilter.ApplyFilter()
                                logger.debug(f"重新应用筛选：{sheet.Name}")
                            # 如有排序需求，可在此补充排序逻辑
                        except Exception as e:
                            logger.debug(f"应用筛选/排序失败：{sheet.Name} - {e}")




                    return True
                time.sleep(5)
            
            logger.error("数据刷新超时！")
            return False
        except Exception as e:
            logger.error(f"刷新异常：{str(e)}")
            return False

    def validate_date(self, check_range, check_frequency) -> bool:
        """带重试的日期校验"""
        for attempt in range(1, check_frequency+1):
            try:
                sheet = self.workbook.Worksheets("日期校验")
                valid = sheet.Range(check_range).Value == 1
                logger.info(f"日期校验 {'通过' if valid else '失败'}（第 {attempt} 次尝试）共{check_frequency}次")
        
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
        logger.info(f"启动任务：{os.path.basename(self.config['excel_path'])}")
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
                    #发送异常通知
                    self._send_wechat(
                        type="text",
                        data={"content": "数据刷新失败，请检查网络！",
                            "mentioned_list": ["zhufuzhe"]
                        },
                        description="数据刷新失败通知",
                        webhook = self.config["data_check"]["warning_webhook"]

                    )
                    return

                # 日期校验
                if self.config.get("data_check_enable", False):
                    check_range = self.config["data_check"]["check_range"]
                    check_frequency = self.config["data_check"]["check_frequency"]
                    if not excel.validate_date(check_range, check_frequency):
                        logger.warning("数据日期校验未通过，发送通知并终止任务")
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
            logger.info(f"任务耗时：{time.time() - start_time:.2f}s")

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
            print(f"正在上传文件：{file_obj.name}")
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
        def thread_func():
            pythoncom.CoInitialize()  # 关键：初始化COM
            try:
                task.execute(self.debug_mode)
            finally:
                pythoncom.CoUninitialize()
        """线程执行任务"""
        thread = threading.Thread(
            target=thread_func,
            name=f"Task-{os.path.basename(task.config['excel_path'])}",
            daemon=True
        )
        thread.start()

    def run_now(self, task_id: int = None):
        """立即执行任务（调试）"""
        logger.info("进入调试模式...")
        targets = self.tasks if task_id is None else [self.tasks[task_id]]
        
        for task in targets:
            try:
                logger.info(f"立即执行：{os.path.basename(task.config['excel_path'])}")
                task.execute(self.debug_mode)
            except Exception as e:
                logger.error(f"执行异常：{str(e)}")

# ---------------------------- 主程序 ----------------------------
def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description="Excel 自动化报表系统")
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