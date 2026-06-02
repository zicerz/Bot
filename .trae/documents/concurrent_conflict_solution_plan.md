# 程序并发冲突解决方案

## 问题分析

当程序运行时再开一个终端运行单次任务时，会产生以下冲突：

1. **Excel文件锁冲突**：两个进程同时尝试打开同一个Excel文件，导致文件被占用错误
2. **COM对象竞争**：多个进程同时初始化和使用Excel COM对象
3. **资源竞争**：对截图临时文件、日志文件的并发读写

## 解决方案

### 方案概述

实现文件锁机制，确保同一时间只有一个进程可以操作Excel文件。使用Python的`fcntl`或`portalocker`库实现跨平台的文件锁。

### 具体实现步骤

#### 1. 添加文件锁工具类

创建一个文件锁工具类，用于管理Excel文件的访问权限：

- 使用文件锁机制防止并发访问
- 实现带超时的锁获取机制
- 添加锁状态检查

#### 2. 修改ExcelProcessor类

在ExcelProcessor中集成文件锁机制：

- 在打开文件前获取锁
- 在关闭文件后释放锁
- 添加锁获取失败时的重试逻辑

#### 3. 修改TaskScheduler类

在任务调度器中添加全局锁检查：

- 启动时检查是否已有实例运行
- 添加进程互斥机制

## 文件修改清单

| 文件 | 修改内容 |
|------|----------|
| `excel.py` | 添加FileLock工具类、修改ExcelProcessor、修改TaskScheduler |

## 超时时间配置

根据用户要求，文件锁超时时间设定为300秒：

| 超时类型 | 默认值 | 说明 |
|----------|--------|------|
| 文件锁等待超时 | **300秒** | 等待获取Excel文件锁的最长时间 |
| 锁轮询间隔 | 2秒 | 每次尝试获取锁的间隔时间 |
| Excel刷新超时 | 300秒 | 数据刷新操作的最大等待时间 |
| 截图操作超时 | 60秒 | 单个截图操作的超时时间 |

**配置原则**：文件锁超时时间设置为300秒，大于单个任务的最大执行时间，确保有足够时间等待锁释放。

## 实现代码

### 文件锁工具类

```python
import os
import time
import fcntl

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
                fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                return True
            except (IOError, BlockingIOError):
                if self.lock_file:
                    self.lock_file.close()
                    self.lock_file = None
                time.sleep(poll_interval)
        
        return False
    
    def release(self):
        """释放文件锁"""
        if self.lock_file:
            try:
                fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_UN)
                self.lock_file.close()
                self.lock_file = None
            except Exception:
                pass
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.release()
```

### 修改ExcelProcessor类

在`__enter__`方法中添加文件锁获取逻辑。

### 修改TaskScheduler类

在`start()`方法中添加全局锁检查。

## 取消文件锁超时机制的影响

如果取消文件锁超时机制（即无限等待锁释放），会带来以下风险：

### 严重风险
1. **死锁风险**：如果持有锁的进程异常退出（如强制终止、崩溃、断电），锁文件不会自动释放，其他进程将永久等待，无法继续执行任务。

2. **资源耗尽**：无限等待会导致进程一直占用系统资源（内存、CPU），但不做任何有用的工作。

3. **用户体验差**：用户无法判断程序是在正常等待还是已经陷入死锁，没有任何反馈。

4. **任务积压**：定时任务会不断累积，导致系统负载增加。

### 对比：有无超时机制的行为差异

| 场景 | 有超时机制 | 无超时机制 |
|------|------------|------------|
| 正常获取锁 | 获取成功，执行任务 | 获取成功，执行任务 |
| 锁被占用 | 等待一段时间后超时退出，记录日志 | 无限等待，直到锁释放 |
| 持有锁进程崩溃 | 超时后退出，下次任务可正常执行 | 永久阻塞，需要人工干预 |
| 用户感知 | 超时后收到通知 | 无任何反馈，程序"卡住" |

### 结论

**不建议取消超时机制**。超时机制是防止死锁的重要保障。如果担心超时时间设置不合理，可以根据实际任务执行时间动态调整，但不应完全取消。

## 风险评估

| 风险项 | 风险等级 | 应对措施 |
|--------|----------|----------|
| 锁文件残留 | 低 | 程序退出时确保释放锁，启动时检查并清理无效锁 |
| 死锁 | 中 | 设置合理的超时时间，避免无限等待 |
| 性能影响 | 低 | 使用非阻塞锁，配合重试机制 |
| 取消超时的风险 | 高 | **强烈建议保留超时机制** |

## 测试方案

1. **并发测试**：启动两个终端同时执行任务，验证锁机制生效
2. **超时测试**：模拟长时间占用，验证超时机制
3. **异常退出测试**：强制终止进程，验证锁自动释放

## 预期效果

- 当一个任务正在执行时，另一个任务会等待直到锁释放
- 如果等待超时，任务会优雅地退出并记录日志
- 避免Excel文件被多个进程同时打开导致的冲突

---

## 实施步骤

### 步骤1：安装依赖

```bash
pip install portalocker
```

### 步骤2：修改excel.py

1. 添加FileLock类
2. 修改ExcelProcessor.__enter__方法
3. 修改TaskScheduler.start方法

### 步骤3：测试验证

运行并发任务，验证锁机制是否正常工作。