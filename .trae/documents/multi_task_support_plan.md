
# 实现 --task 参数支持多任务配置计划

## 概述
修改命令行接口，使 `--task` 参数支持一次性配置多个任务，提高使用灵活性。

## 当前实现分析
- `excel.py` 的 `main()` 函数（第1126-1167行）处理命令行参数
- 目前 `--task` 使用 `type=str`，只接受单个任务标识
- `TaskScheduler.run_now()` 方法（第1085-1123行）目前只处理单个 `task_id` 和 `webhook_id`

## 修改方案

### 1. 修改命令行参数解析（main函数）
- 将 `--task` 参数改为支持多个值的列表形式（使用 `nargs='*'`）
- 保持向后兼容：单个参数仍然有效
- 支持混合格式：可以同时包含任务索引和任务索引-webhook索引格式

### 2. 修改 TaskScheduler.run_now() 方法
- 将方法签名从接受单个 `task_id` 和 `webhook_id` 改为接受任务列表
- 每个任务项可以包含任务索引和可选的 webhook 索引
- 支持执行多个指定的任务组合

### 3. 具体实现步骤
1. 修改 `argparse` 中 `--task` 的定义，使用 `nargs='*'`
2. 添加解析多个任务参数的逻辑
3. 重构 `run_now()` 方法，支持处理任务列表
4. 更新相关的日志输出和用户提示信息

## 预期使用方式
```bash
# 单个任务（保持兼容）
python excel.py --task 0
python excel.py --task 0-1

# 多个任务
python excel.py --task 0 1 2
python excel.py --task 0-1 1-0 2
```

## 修改文件
- `excel.py`
