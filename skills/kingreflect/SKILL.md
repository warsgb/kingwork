---
name: kingreflect
description: 日报周报月报生成技能。汇总指定周期内的工作数据，用大模型生成 Markdown 格式报告，包含客户跟进、学习成长、惊喜内容等。
---

# kingreflect - 日报周报月报

## 用法

```bash
# 生成今日日报
python skills/kingreflect/run.py --period daily

# 生成本周周报
python skills/kingreflect/run.py --period weekly

# 生成本月月报
python skills/kingreflect/run.py --period monthly

# 指定日期范围
python skills/kingreflect/run.py --start 2026-03-01 --end 2026-03-15

# 输出到文件
python skills/kingreflect/run.py --period daily --output daily_report.md

# 自然语言（由 OpenClaw 框架解析周期）
python skills/kingreflect/run.py "帮我写今天的日报"
```
