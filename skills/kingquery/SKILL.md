---
name: kingquery
description: 数据查询和统计技能。查询客户跟进统计、项目进展、工作数据分析、客户沟通时间线等。
---

# kingquery - 数据查询

## 用法

```bash
# 客户跟进统计（最近30天）
python skills/kingquery/run.py stats --type customer_followup --period 30d

# 客户沟通时间线
python skills/kingquery/run.py timeline --customer "某某公司"

# 查询所有客户
python skills/kingquery/run.py customers

# 查询项目列表
python skills/kingquery/run.py projects

# 查询最近工作记录
python skills/kingquery/run.py recent --days 7

# 搜索（关键词）
python skills/kingquery/run.py search "某某公司"
```
