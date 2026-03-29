---
name: kingconfig
description: KingWork 配置管理。管理工作类型（增删改查、同步到多维表和提示词），以及管理任意数据表的枚举字段选项。
---

# kingconfig - 配置管理

## Agent 调用方式

### 工作类型管理

```bash
# 列出所有工作类型
python skills/kingconfig/run.py list-work-types

# 新增工作类型
python skills/kingconfig/run.py add-work-type "技术预研" --shortcut k10 --keywords "预研,调研,技术探索" --team-mapping "其他"

# 重命名工作类型
python skills/kingconfig/run.py rename-work-type "问题服务" "售后服务"

# 删除工作类型
python skills/kingconfig/run.py remove-work-type "灵感记录"
```

### 通用枚举字段管理

```bash
# 列出所有可配置的枚举字段
python skills/kingconfig/run.py list-enums

# 查看某个枚举的当前值
python skills/kingconfig/run.py list-enum "客户状态"

# 给枚举字段添加选项
python skills/kingconfig/run.py add-enum "客户状态" "休眠"

# 删除枚举字段选项
python skills/kingconfig/run.py remove-enum "学习类型" "其他"

# 重命名枚举选项
python skills/kingconfig/run.py rename-enum "客户状态" "流失" "已流失"
```

## 修改范围

工作类型变更时，自动同步以下文件：
1. `config/kingwork.yaml` — types、shortcuts、keywords、team.work_type_mapping
2. `config/fields_enum.yaml` — diary_records.工作类型
3. `config/tables.yaml` — diary_records.工作类型.options
4. `config/prompts.yaml` — work_type_classification 的类型列表

通用枚举变更时，自动同步：
1. `config/fields_enum.yaml` — 对应表.字段
2. `config/tables.yaml` — 对应表.字段.options
