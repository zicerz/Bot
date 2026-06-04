# 任务级重试通知机制优化方案

## 需求分析

**用户需求**：当任务级重试启用时，任务第一次失败不发送通知，三次都失败后再发送通知。

**当前问题**：
- 当前代码在任务执行失败时（如数据刷新失败）会立即发送通知
- 即使启用了重试机制，第一次失败也会发送通知，导致通知过于频繁

## 代码分析

### 当前重试配置
```yaml
retry:
  enabled: 1                      # 是否启用任务级重试（1启用，0禁用）
  delay_minutes: 5               # 重试间隔时间（分钟）
  max_attempts: 3                 # 最大重试次数
```

### 当前通知逻辑

1. **数据刷新失败通知**（`excel.py` 第828-841行）：
   - 当 `excel.refresh_data()` 返回 False 时，立即发送失败通知
   - 无论是否启用重试，都会发送通知

2. **最终失败通知**（`excel.py` 第1182-1201行）：
   - `send_final_failure_notification()` 方法已存在
   - 在 `_schedule_retry` 中，当重试次数超过最大值时调用

3. **手动执行任务**（`run_now` 方法）：
   - 处理 `--task` 和 `--run-all` 参数
   - 调用 `task.execute()` 执行任务，但**不会触发重试机制**
   - 失败时会记录日志，但当前会立即发送通知

### 核心问题
- `execute` 方法中的刷新失败通知逻辑没有考虑重试机制的状态
- 手动执行和定时任务使用相同的 `execute` 方法，无法区分处理

## 修改方案

### 修改目标

**定时任务（由调度器触发）**：
- 当重试机制启用时：首次失败和重试失败都不发送通知，仅在达到最大重试次数后发送最终失败通知
- 当重试机制禁用时：失败时立即发送通知

**手动执行（`--task` 参数）**：
- 无论重试机制是否启用，失败时都立即发送通知（因为用户正在等待结果）

### 修改位置

#### 1. 修改 `ReportTask.execute()` 方法（第828-841行）

**新增参数**：添加 `is_manual=False` 参数，用于区分手动执行和定时任务

**修改后代码**：
```python
def execute(self, debug_mode=False, webhook_configs=None, is_manual=False):
    """
    执行任务流程
    :param debug_mode: 是否调试模式
    :param webhook_configs: 特定的webhook配置（单个dict或列表，None表示执行所有webhook）
    :param is_manual: 是否手动执行（用于决定是否立即发送失败通知）
    :return: True表示成功，False表示失败
    """
    # ... 省略其他代码 ...
    
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
```

#### 2. 修改 `TaskScheduler.run_now()` 方法（第1619-1667行）

在调用 `task.execute()` 时传入 `is_manual=True` 参数：

**修改后代码**：
```python
if webhook_id is not None:
    webhook_config = task.config["webhooks"][webhook_id]
    logger.info(f"执行任务 {task_specs[idx-1][0]} 的 webhook {webhook_id}: {webhook_config['webhook'].split('key=')[-1][:10]}...")
    success = task.execute(self.debug_mode, webhook_config, is_manual=True)
else:
    # 执行所有webhook配置
    if task_specs:
        logger.info(f"执行任务 {task_specs[idx-1][0]}: {task_name}")
    success = task.execute(self.debug_mode, is_manual=True)
```

#### 3. 修改 `TaskScheduler._run_task()` 方法（第1404行）

保持不变，`is_retry` 参数已存在，定时任务调用时 `is_manual` 默认为 False

### 修改文件列表

| 文件路径 | 修改内容 |
| :--- | :--- |
| `excel.py` | 修改 `ReportTask.execute()` 方法，添加 `is_manual` 参数 |
| `excel.py` | 修改 `TaskScheduler.run_now()` 方法，传入 `is_manual=True` |

### 预期行为

| 场景 | 重试启用 | 执行方式 | 通知行为 |
| :--- | :--- | :--- | :--- |
| 首次执行失败 | 是 | 定时任务 | 不发送通知，进入重试 |
| 重试失败 | 是 | 定时任务 | 不发送通知，继续重试 |
| 达到最大重试次数 | 是 | 定时任务 | 发送最终失败通知 |
| 任意失败 | 否 | 定时任务 | 立即发送通知 |
| 任意失败 | 是 | 手动执行（--task） | 立即发送通知 |
| 任意失败 | 否 | 手动执行（--task） | 立即发送通知 |

### 风险评估

1. **低风险**：修改仅涉及通知逻辑，不影响核心业务流程
2. **兼容性**：保留了重试禁用时的原有行为，不影响现有配置
3. **可测试**：可通过配置不同的重试参数验证行为
4. **手动执行支持**：手动执行时立即通知，符合用户预期

### 验证方案

1. **定时任务测试**（重试启用）：
   - 模拟任务失败，验证首次失败不发送通知
   - 验证达到最大重试次数后发送通知

2. **定时任务测试**（重试禁用）：
   - 模拟任务失败，验证立即发送通知

3. **手动执行测试**（重试启用）：
   - 使用 `--task` 参数执行，验证失败时立即发送通知

4. **手动执行测试**（重试禁用）：
   - 使用 `--task` 参数执行，验证失败时立即发送通知