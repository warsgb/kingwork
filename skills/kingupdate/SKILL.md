---
name: kingupdate
description: 更新 KingWork 多维表中的记录。支持描述定位（近3天）、交互确认、模式B结构化输出。Agent 调用方式。
---

# kingupdate - 记录更新

## Agent 调用方式（模式 B）

由 agent 调用，skill 返回结构化结果，agent 自然语言展示给用户。

### 搜索候选

```bash
python skills/kingupdate/run.py search "<描述>" [--days 3]
```

返回 JSON：
```json
{
  "candidates": [
    {
      "idx": 1,
      "sheet_name": "01日记记录",
      "sheet_id": "2",
      "record_id": "V",
      "type": "客户跟进",
      "customer": "泰康集团",
      "project": "WPS365部署项目",
      "content": "拜访泰康集团张总，沟通WPS365私有云部署需求...",
      "date": "2026/03/21",
      "source": "手动输入"
    }
  ],
  "needs_selection": true,
  "query": "泰康集团",
  "total": 1
}
```

### 执行更新

```bash
python skills/kingupdate/run.py update <record_id> <sheet_id> \
  --type "客户跟进" \
  --customer "新客户名" \
  --project "新项目名" \
  --content "新内容"
```

## 覆盖范围

- 日记记录表（sheet_id 动态从 tables.sheet_ids 读取）
- 客户跟进记录表
- 项目档案表
- 搜索范围：近 3 天（可配置）

## 模式 B 说明

- skill 不维持状态，不做交互式输入
- 所有需要用户确认的内容通过结构化 JSON 返回给 agent
- agent 负责自然语言展示和转译用户回复
