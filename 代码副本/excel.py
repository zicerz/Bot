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

class ExcelScreenshotBot:
    def __init__(self, config_path="config.yml"):
        self.config = self.load_config(config_path)
        self.excel = None
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

    def load_config(self, config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
        
    """日期校验"""  
    def check_excel_date(self,wb):
        try:
            ws = wb.Worksheets('日期校验')
            cell_value = ws.Range('D4').Value
            cell_value_data = ws.Range('B2').Value2
            print(f"日期校验单元格：{cell_value}")
            print(f"最新成单日期：{cell_value_data}")
            if cell_value == 1:
                print("日期校验通过")
                return 1
            else :
                print("日期校验失败，取消发送")
                data = {
                    "msgtype": "text",
                    "text": {
                        "content": "日期校验失败，请检查",
                        "mentioned_list":["hezhengsong","zhufuzhe"],
                    }
                }
                response = requests.post(
                    "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=ee57f939-0e4e-4707-beba-422dae91538e",
                    json=data,
                    headers={'Content-Type': 'application/json'}
                )
                if response.json().get('errcode') == 0:
                    print(f"已发送提醒")
                else:
                    print(f"[错误] 发送失败：{response.text}")
                return 0
            
        except Exception as e:
            print(f"[错误] 日期校验出错：{str(e)}")
            return 0

    """截取Excel指定区域为图片"""
    def capture_range(self, worksheet, range_address, output_path):
        try:
            range_obj = worksheet.Range(range_address)
            range_obj.CopyPicture(Format=1) # 1 = xlBitmap
            chart = worksheet.ChartObjects().Add(0, 0, range_obj.Width, range_obj.Height)
            chart.Activate()
            self.excel.ActiveChart.Paste()
            chart.Chart.Export(output_path)
            chart.Delete()
            return True
        except Exception as e:
            print(f"[错误] 截取区域 {range_address} 失败：{str(e)}")
            return False

    def capture_all_ranges(self):
        """截取所有配置的区域"""
        wb = None
        try:
            self.excel = win32.Dispatch('Excel.Application')
            self.excel.Visible = False              # 显示Excel窗口
            self.excel.DisplayAlerts = False        # 禁用提示框
            wb = self.excel.Workbooks.Open(self.config['excel_path'])
            
            # 刷新所有数据连接和数据透视表
            print("正在刷新数据...")
            wb.RefreshAll()
            # 等待所有异步查询完成
            self.excel.CalculateUntilAsyncQueriesDone()
            
            # 检查日期
            print("正在检查日期...")
            if self.check_excel_date(wb) == 0:
                print("日期校验失败，停止截图")
                return False

            screenshot_paths = []
            
            for item in self.config['capture_ranges']:
                ws = wb.Worksheets(item['sheet_name'])
                
                
                range_addr = item['range']
                range_name = item.get('name', range_addr)
                output_path = os.path.join(self.script_dir, f"{range_name}_{int(time.time())}.png")
                if self.capture_range(ws, range_addr, output_path) and os.path.exists(output_path):
                    print(f"已截取区域：{range_name} ({range_addr}) -> {output_path}")
                    screenshot_paths.append(output_path)

            return screenshot_paths
        except Exception as e:
            print(f"[错误] Excel操作异常：{str(e)}")
            return []
        finally:
            if hasattr(self, 'excel') and self.excel:
                try:
                    if wb is not None:
                        wb.Close(SaveChanges=True) # 保存更改
                except Exception as e:
                    print(f"关闭工作簿失败: {str(e)}")
                self.excel.Quit()
                # 强制释放COM资源
                del self.excel


    def send_to_wechat(self, image_path):
        """发送图片到企业微信"""
        try:
            if not os.path.exists(image_path):
                print(f"[错误] 文件不存在：{image_path}")
                return

            with open(image_path, 'rb') as f:
                image_data = f.read()
            
            
            
            data = {
                "msgtype": "image",
                "image": {
                    "base64": base64.b64encode(image_data).decode('utf-8'),
                    "md5": hashlib.md5(image_data).hexdigest()
                }
            }

            response = requests.post(
                self.config['webhook_url'],
                json=data,
                headers={'Content-Type': 'application/json'}
            )
            
            if response.json().get('errcode') == 0:
                print(f"已发送图片：{os.path.basename(image_path)}")
            else:
                print(f"[错误] 发送失败：{response.text}")
        except Exception as e:
            print(f"[错误] 发送异常：{str(e)}")

    # 获取media_id
    def upload_robot_file(self):
        
        try:
            
            
            with open(self.config['file_path'], 'rb') as f:
                files = {
                    # 'media': (file_name, f, 'application/octet-stream')  # 显式指定文件名和类型
                    'media': ("奖励池名单 " + datetime.now().strftime('%Y-%m-%d') +".xlsx", f, 'application/octet-stream')
                }
                
                # 发送POST请求
                response = requests.post(
                    url=self.config['upload_url'],
                    files=files
                )
                
                response.raise_for_status()  # 自动处理HTTP错误
                
                # 解析响应数据
                result = response.json()
                if result.get('errcode') != 0:
                    return None, f"API返回错误：{result.get('errmsg')}"
                return result.get('media_id'), None
                
        except IOError as e:
            return None, f"文件读取失败：{str(e)}"
        except Exception as e:
            return None, f"请求异常：{str(e)}"
        
    def send_file_to_wechat(self):
        """发送文件到企业微信"""
        try:
            media_id, error = self.upload_robot_file()
            if error or not media_id:
                print(f"❗ 文件上传失败：{error}")
                return
            
            file_data = {
                "msgtype": "file",
                "file": {
                    "media_id": media_id,
                }
            }
            # 发送请求
            response = requests.post(
                self.config['webhook_url'],
                json=file_data,
                headers={'Content-Type': 'application/json'}
            )
            
            if response.json().get('errcode') == 0:
                print(f"已发送文件：奖励池名单 " + datetime.now().strftime('%Y-%m-%d') +".xlsx")
            else:
                print(f"[错误] 发送失败：{response.text}")
        except Exception as e:
            print(f"[错误] 发送异常：{str(e)}")

    def run_job(self):
        """执行任务"""
        print(f"\n===== 任务触发 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====")
        
        screenshots = self.capture_all_ranges()
        if screenshots == False:
            print("日期校验失败，跳过后续操作")
        else:
            #发送截图
            for img_path in screenshots:
                print(f"正在发送图片：{img_path}")
                self.send_to_wechat(img_path)
                try:
                    if os.path.exists(img_path):
                        os.remove(img_path)
                except Exception as e:
                    print(f"[警告] 文件删除失败：{str(e)}")
            # 发送文件
            if self.config['send_file'] and self.config['file_path']:
                print(f"正在发送文件：{self.config['file_path']}")
                self.send_file_to_wechat()
            else:
                print("未配置文件路径。")
        
        
        print("===== 任务完成 =====")

    def setup_schedule(self):
        """配置定时任务"""
        for trigger_time in self.config['schedule']['times']:
            schedule.every().day.at(trigger_time).do(self.run_job)
            print(f"已设置定时任务：每天 {trigger_time}")

    def start(self):
        if self.config['schedule']['enabled']:
            # 定时模式
            self.setup_schedule()
            while True:
                schedule.run_pending()
                time.sleep(1)
                
        else:
            # 单次模式
            self.run_job()
            

if __name__ == "__main__":
    bot = ExcelScreenshotBot()
    try:
        bot.start()
    except KeyboardInterrupt:
        print("\n程序已退出")