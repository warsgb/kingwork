---
name: kingteam
description: 将个人 KingWork 数据同步至团队。从个人多维表读取周期内数据，字段映射后写入团队售前日报表；生成日报/周报并上传至团队文档文件夹。
---

# kingteam - 团队数据同步

## 用法

```bash
# 同步客户跟进记录（默认同步今天）
python skills/kingteam/run.py sync_customers

# 指定日期范围同步
python skills/kingteam/run.py sync_customers --start 2026-03-20 --end 2026-03-23

# 生成并上传日报
python skills/kingteam/run.py sync_daily

# 生成并上传周报
python skills/kingteam/run.py sync_weekly

# 同时同步客户跟进 + 日报（一次性完成）
python skills/kingteam/run.py sync_all
```

## 配置

team 相关配置统一存放在 `kingwork.yaml` 中：

```yaml
team:
  drive_id: "3064807774"           # kingworkteam drive ID
  daily_folder: "日报"             # 日报文件夹名称
  weekly_folder: "周报"            # 周报文件夹名称
  dbsheet_id: "489561486981"       # 团队多维表（售前日报表）
  sheet_id: "11"                   # 售前日报表 sheet_id

personal:
  dbsheet_id: "cbMwPNjcGRwD"       # 个人多维表 ID
```

## 同步说明

- **记录同步**（`sync_customers`）：读取个人日记中配置的可同步类型记录，LLM 智能映射字段后写入团队日报表；冲突策略为 skip（跳过）
- **可同步类型**（`kingwork.yaml` → `sync.sync_types`）：客户跟进、横向支持、团队事务、学习成长；默认全部开启
- **日报/周报**：调用 kingreflect 生成报告，Markdown 格式上传至团队文档文件夹
- 所有操作结果以结构化 JSON 输出，方便 agent 展示给用户
