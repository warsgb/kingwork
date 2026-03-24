---
name: kingbrowse
description: 资料检索技能。在多个 WPS 多维表中按关键词模糊搜索，返回资料名称、链接、来源和匹配说明。配置见 config.yaml。
---

# kingbrowse - 资料检索

## 入口

**CLI 命令：**
```bash
python skills/kingbrowse/run.py search <关键词> [--top N]
python skills/kingbrowse/run.py search 端云一体 --top 5
```

**Agent 触发：**
当用户询问「帮我找一下xxx相关的资料/方案/案例」时，agent 调用 `kingbrowse.browse.search(keyword)` 获取结果。

## 搜索范围

在 `config.yaml` 中配置，支持多文件、多表：

```yaml
browse_sources:
  - file_id: "多维表file_id"
    file_name: "多维表名称（用于展示）"
    sheets:
      - sheet_name: "表名"
        search_fields:   # 搜索哪些字段
          - "字段A"
          - "字段B"
        name_field: "字段A"   # 作为结果标题的字段
        link_fields:        # 提取链接的字段（支持URL和Attachment类型）
          - "链接字段"
```

- 不写 `sheets` 配置 → 自动搜索该文件下所有表
- `search_fields` 不写 → 自动用 name_field
- Attachment 类型字段自动解析为 WPS 分享链接

## 搜索规则

- **模糊搜索**：关键词出现在字段中即匹配（不区分大小写）
- **匹配度评分**：精确开头 > 包含 > 部分包含
- **去重**：同一记录多条匹配只保留最高分
- **排序**：按匹配度降序返回

## 返回格式

```json
[
  {
    "file_id": "xxxxx",
    "file_name": "2026环京三角协同共享",
    "sheet_name": "主打方案库",
    "record_id": "M",
    "name": "端升云内网端云一体",
    "link": "https://www.kdocs.cn/l/ctoNYOZ8QqV3",
    "snippet": "数据不连通外网，将 WPS 专业版端升级至云端...",
    "match_field": "方案名称",
    "match_score": 0.95,
    "relevance": "高"
  }
]
```

## 配置文件

`config.yaml` 中 `browse_sources` 列表管理所有数据源。
如需新增多维表，在此处添加即可，无需改动代码。
