# 多Webhook配置功能实现计划

## 需求分析

用户希望在同一个任务中配置多个webhook，每个webhook可独立配置：
- `times`：发送时间列表
- `capture_configs`：截图配置列表
- `send_file_enable`：是否发送文件（可选，未配置默认为0）

## 配置结构设计

```yaml
tasks:
  - excel_path: "xxx.xlsx"
    # 多个webhook配置，每个可独立设置
    webhooks:
      - webhook: "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxx1"
        times: ["08:35", "10:35"]
        capture_configs:
          - sheet_name: "sheet2"
            range: B2
            name: "webhook1_specific"
        send_file_enable: 1  # 可选，未配置默认为0
      
      - webhook: "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxx2"
        times: ["08:40", "10:40"]
        capture_configs:
          - sheet_name: "sheet1"
            range: A2
            name: "webhook2"
        send_file_enable: 1 # 可选，未配置默认为0
```

## 修改内容

### 1. 配置文件结构更新 (config.yml)
- 移除原有的 `schedule` + `capture_configs` + `send_file_enable` 结构
- 新增 `webhooks` 列表，每个webhook包含独立的times、capture_configs、send_file_enable

### 2. 代码修改 (excel.py)
- **ReportTask类**：重构为支持多个webhook配置，每个webhook独立执行
- **TaskScheduler类**：修改任务调度逻辑，为每个webhook独立调度
- **任务执行流程**：根据触发的webhook配置执行对应的截图和发送逻辑

## 实现步骤

| 步骤 | 任务 | 说明 |
|------|------|------|
| 1 | 修改配置文件schema验证 | 更新`_validate_config`方法支持新的webhooks列表结构 |
| 2 | 修改ReportTask初始化 | 支持解析多个webhook配置 |
| 3 | 修改任务调度逻辑 | 为每个webhook的times独立创建定时任务 |
| 4 | 修改任务执行逻辑 | 根据webhook配置执行对应截图和发送 |
| 5 | 更新示例配置 | 在config.yml中添加多webhook示例 |

## 兼容性考虑

- 支持旧版配置结构（单webhook使用schedule），自动转换为新版格式

## 文件修改清单

| 文件 | 修改类型 | 说明 |
|------|----------|------|
| `config.yml` | 修改 | 添加多webhook配置示例 |
| `excel.py` | 修改 | 支持多webhook配置的解析和执行 |