# 任务序列号显示计划

## 目标
在运行 `python excel.py --debug` 时显示的已安排任务中，每个任务都显示任务序列号和子任务（webhook）序列号。

## 当前状态
`_schedule_tasks` 方法在第940行获取任务名称时只使用文件名，没有显示任务序列号：
```python
task_name = os.path.basename(task.config["excel_path"])
```

第948行的日志输出格式：
```
已安排任务：08:35 → 学员明细备份.xlsx → webhook[833b098e]
```

## 修改方案

### 文件：excel.py

**修改位置 1：第939-942行**
```python
# 修改前
for task in self.tasks:
    task_name = os.path.basename(task.config["excel_path"])
    
    for webhook_config in task.config["webhooks"]:

# 修改后
for task_idx, task in enumerate(self.tasks):
    task_name = task.config.get("name", os.path.basename(task.config["excel_path"]))
    
    for webhook_idx, webhook_config in enumerate(task.config["webhooks"]):
```

**修改位置 2：第948行**
```python
# 修改前
logger.info(f"已安排任务：{trigger_time} → {task_name} → webhook[{webhook_key}]")

# 修改后
logger.info(f"已安排任务：{trigger_time} → [{task_idx}] {task_name} → webhook[{task_idx}-{webhook_idx}][{webhook_key}]")
```

## 预期输出效果

修改前：
```
已安排任务：08:35 → 学员明细备份.xlsx → webhook[833b098e]
```

修改后：
```
已安排任务：08:35 → [0] 学员明细自动播报 → webhook[0-0][833b098e]
已安排任务：08:40 → [0] 学员明细自动播报 → webhook[0-1][ee57f939]
```

## 实施步骤

1. 修改 `_schedule_tasks` 方法中的外层循环，添加 `enumerate` 获取任务索引
2. 修改任务名称获取逻辑，优先使用 `name` 字段
3. 修改内层循环，添加 `enumerate` 获取webhook索引
4. 修改日志输出格式，添加任务序列号和webhook序列号（格式：`任务索引-webhook索引`）
