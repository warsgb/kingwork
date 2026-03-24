---
name: kingauto
description: AI 自动分析技能。自动抓取当天聊天记录和会议记录，用大模型分析提取工作相关信息，去重后写入多维表，并提取惊喜内容。
---

# kingauto - AI 自动分析

## 用法

```bash
# 分析当天工作（默认）
python skills/kingauto/run.py

# 指定日期
python skills/kingauto/run.py --date 2026-03-19

# 指定日期范围
python skills/kingauto/run.py --start 2026-03-01 --end 2026-03-15

# 仅分析聊天记录
python skills/kingauto/run.py --source chat

# 仅分析会议记录
python skills/kingauto/run.py --source meeting

# 跳过相似度去重
python skills/kingauto/run.py --no-dedup

# 详细输出
python skills/kingauto/run.py --verbose
```

## 处理流程

1. 获取聊天记录（WPS365 IM Skill）
2. 获取会议记录（WPS365 Meeting Skill）
3. 大模型逐条分析（判断是否工作相关、提取信息）
4. 相似度去重（与已有记录比较）
5. 写入多维表各业务表
6. 提取惊喜内容
