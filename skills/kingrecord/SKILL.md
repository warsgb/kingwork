---
name: kingrecord
description: 工作日记记录技能。支持自然语言输入，自动分类（客户跟进/待办/学习/横向支持/团队事务/灵感），并分发到对应数据表。
---

# kingrecord - 工作日记记录

## 用法

```bash
# 自然语言输入（最常用）
python skills/kingrecord/run.py "今天和某某公司王总电话沟通了新项目需求"

# 快捷指令（跳过AI分类，直接指定类型）
python skills/kingrecord/run.py k1 "客户拜访内容"   # 客户跟进
python skills/kingrecord/run.py k2 "待办事项"        # 待办事项
python skills/kingrecord/run.py k3 "学习内容"        # 学习成长
python skills/kingrecord/run.py k4 "横向支持"        # 横向支持
python skills/kingrecord/run.py k5 "团队事务"        # 团队事务
python skills/kingrecord/run.py k6 "灵感记录"        # 灵感记录

# 结构化输入
python skills/kingrecord/run.py --type "客户跟进" --customer "某某公司" "内容"

# 关键词指定
python skills/kingrecord/run.py --keyword "客户" "内容"

# 详细模式
python skills/kingrecord/run.py --verbose "工作内容"
```

## 快捷指令映射
| 指令 | 类型 |
|------|------|
| k1 | 客户跟进 |
| k2 | 待办事项 |
| k3 | 学习成长 |
| k4 | 横向支持 |
| k5 | 团队事务 |
| k6 | 灵感记录 |
