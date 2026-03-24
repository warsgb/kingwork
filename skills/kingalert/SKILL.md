---
name: kingalert
description: AI 智能提醒技能。查询未完成待办事项和超期未跟进客户，生成提醒清单，支持交互式标记完成。
---

# kingalert - AI 智能提醒

## 用法

```bash
# 查看所有提醒（默认：待办 + 客户跟进 + 今日日程）
python skills/kingalert/run.py

# 仅查看待办事项
python skills/kingalert/run.py --todos-only

# 仅查看客户跟进提醒
python skills/kingalert/run.py --customers-only

# 仅查看今日日程
python skills/kingalert/run.py --calendar-only

# 标记待办完成
python skills/kingalert/run.py --complete <todo_id>

# 标记客户已跟进
python skills/kingalert/run.py --followup <客户名称>

# 自定义未跟进天数（默认15天）
python skills/kingalert/run.py --inactive-days 30
```

## 输出示例

```
## 工作提醒 - 2026-03-20

### 待办事项（3项）

1. [高优先级] 准备某某公司演示材料
   - 到期时间：2026-03-21
   - 关联客户：某某公司
   - 状态：明天到期

### 客户跟进提醒（2家）

1. 某某公司 - 未跟进18天
   - 最近跟进：2026-03-02
   - 建议：电话回访
```
