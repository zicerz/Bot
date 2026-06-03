
# 文件备份功能实现方案

## 一、需求分析

用户要求在每个任务结束后将对应Excel文件备份到指定文件夹中，文件夹结构按照以下层级分级：

```
备份根目录/
├── 2026-06/                    # 年月目录（格式：YYYY-MM）
│   ├── 学员明细备份_xlsx/       # 文件名_扩展名目录（避免同文件名不同扩展名冲突）
│   │   ├── 学员明细备份_20260603_083500.xlsx   # 版本文件（格式：文件名_日期_时间.扩展名）
│   │   ├── 学员明细备份_20260603_103500.xlsx
│   │   └── ...
│   └── 往期顺次_xlsx/
│       ├── 往期顺次_20260603_083500.xlsx
│       └── ...
├── 2026-05/
│   └── ...
```

**关键改进**：目录名称采用 `文件名_扩展名` 格式（如 `学员明细备份_xlsx`），解决同文件名不同扩展名的冲突问题。

## 二、实现方案

### 1. 配置文件修改

在 `config.yml` 中添加备份相关配置：

```yaml
# 备份配置
backup:
  enable: true                    # 是否启用备份功能
  backup_dir: "./backups"        # 备份根目录路径（支持相对路径和绝对路径）
```

### 2. 代码修改

在 `excel.py` 中实现备份功能：

#### 2.1 添加备份工具函数

在 `ReportTask` 类中添加 `_backup_file` 方法：

```python
def _backup_file(self):
    """
    在任务结束后备份文件
    目录结构：backup_dir/YYYY-MM/文件名_扩展名/文件名_YYYYMMDD_HHMMSS.扩展名
    
    设计说明：
    - 年月目录：按 YYYY-MM 格式组织，便于按时间查找
    - 文件目录：采用 文件名_扩展名 格式，避免同文件名不同扩展名冲突
    - 备份文件名：文件名_日期时间.扩展名，确保唯一性
    """
    # 检查是否启用备份功能
    backup_config = self.config.get("backup", {})
    if not backup_config.get("enable", False):
        self.logger.debug("备份功能未启用")
        return
    
    source_path = self.config.get("file_path") or self.config.get("excel_path")
    if not source_path or not os.path.exists(source_path):
        self.logger.warning(f"源文件不存在，跳过备份：{source_path}")
        return
    
    # 获取备份根目录
    backup_dir = backup_config.get("backup_dir", "./backups")
    if not os.path.isabs(backup_dir):
        backup_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), backup_dir))
    
    # 构建备份目录结构
    now = datetime.now()
    year_month = now.strftime("%Y-%m")  # 年月目录：2026-06
    file_name = os.path.basename(source_path)
    file_base, file_ext = os.path.splitext(file_name)  # 文件名和扩展名
    ext_without_dot = file_ext[1:].lower() if file_ext else ""  # 去除点号的扩展名
    
    # 创建目录结构：backup_dir/YYYY-MM/文件名_扩展名
    # 使用 文件名_扩展名 作为目录名，解决同文件名不同扩展名冲突
    month_dir = os.path.join(backup_dir, year_month)
    file_dir = os.path.join(month_dir, f"{file_base}_{ext_without_dot}")
    
    try:
        os.makedirs(file_dir, exist_ok=True)
        
        # 生成备份文件名：文件名_YYYYMMDD_HHMMSS.扩展名
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{file_base}_{timestamp}{file_ext}"
        backup_path = os.path.join(file_dir, backup_filename)
        
        # 复制文件（保留元数据）
        shutil.copy2(source_path, backup_path)
        
        self.logger.info(f"文件备份成功：{backup_path}")
        return backup_path
    except Exception as e:
        self.logger.error(f"文件备份失败：{str(e)}")
        return None
```

#### 2.2 在任务执行完成后调用备份

在 `ReportTask.execute()` 方法的 `finally` 块中调用备份方法：

```python
def execute(self, debug_mode=False, webhook_configs=None):
    # ... 原有代码 ...
    finally:
        elapsed_time = time.time() - start_time
        self.logger.info(f"任务耗时：{elapsed_time:.2f}s")
        
        # 执行文件备份
        self._backup_file()
        
        separator = "=" * 100
        self.logger.info(separator)
        print(separator)
```

## 三、预期效果

### 目录结构示例

```
backups/
├── 2026-06/
│   ├── 学员明细备份_xlsx/
│   │   ├── 学员明细备份_20260603_083500.xlsx
│   │   ├── 学员明细备份_20260603_103500.xlsx
│   │   └── ...
│   ├── 往期顺次_xlsx/
│   │   ├── 往期顺次_20260603_083500.xlsx
│   │   └── ...
│   └── 数据报表_csv/          # 同文件名不同扩展名的情况
│       ├── 数据报表_20260603_090000.csv
│       └── ...
└── 2026-05/
    ├── 学员明细备份_xlsx/
    │   └── ...
    └── 往期顺次_xlsx/
        └── ...
```

### 日志输出示例

```
2026-06-03 08:35:00 [INFO] [任务0] 文件备份成功：C:/Users/EDY/Desktop/BOT_TEST/backups/2026-06/学员明细备份_xlsx/学员明细备份_20260603_083500.xlsx
```

### 配置示例

```yaml
backup:
  enable: true
  backup_dir: "./backups"

tasks:
  - name: "学员明细自动播报"
    excel_path: "C:/Users/EDY/OneDrive - PPLINGO PTE LTD/自动播报文件/学员明细备份.xlsx"
    # ... 其他配置 ...
```

## 四、修改清单

| 文件 | 修改内容 |
|------|----------|
| `config.yml` | 添加 backup 配置段 |
| `excel.py` | 添加 `_backup_file` 方法，在 `execute` 方法中调用 |

## 五、依赖说明

需要导入 `shutil` 模块用于文件复制操作（代码中已有 `os` 导入，需新增 `shutil`）。

## 六、特殊情况处理

| 场景 | 处理方式 | 示例 |
|------|----------|------|
| 同文件名不同扩展名 | 目录名包含扩展名 | `data.xlsx` → `data_xlsx/`, `data.csv` → `data_csv/` |
| 无扩展名文件 | 目录名仅为文件名 | `README` → `README_/` |
| 扩展名大小写不同 | 统一转为小写 | `DATA.XLSX` → `DATA_xlsx/` |
