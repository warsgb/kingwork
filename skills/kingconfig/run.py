#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kingconfig - KingWork 配置管理
管理工作类型（增删改查）和通用枚举字段，自动同步到所有配置文件。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import os
import copy
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# 将 kingwork 根目录加入 path
_root = Path(__file__).resolve().parent.parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import yaml
from kingwork_client.base import KINGWORK_ROOT

# ── 配置文件路径 ──
CFG_PATH = KINGWORK_ROOT / "config" / "kingwork.yaml"
ENUM_PATH = KINGWORK_ROOT / "config" / "fields_enum.yaml"
TABLES_PATH = KINGWORK_ROOT / "config" / "tables.yaml"
PROMPTS_PATH = KINGWORK_ROOT / "config" / "prompts.yaml"


# ============================================================
# YAML 读写工具
# ============================================================

def _load_yaml(path: Path) -> dict:
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _save_yaml(path: Path, data: dict):
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


def _load_yaml_raw(path: Path) -> str:
    """读取原始文本，用于 prompts.yaml 的正则替换（保留格式）。"""
    if not path.exists():
        return ""
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def _save_raw(path: Path, content: str):
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


# ============================================================
# 工作类型管理
# ============================================================

def _get_work_types_dict(cfg: dict) -> dict:
    """返回 {类型名: 描述} 字典。兼容旧版 list 格式。"""
    types = cfg.get("work_types", {}).get("types", {})
    if isinstance(types, list):
        return {t: "" for t in types}
    return types if isinstance(types, dict) else {}


def _get_work_type_names(cfg: dict) -> list:
    """返回类型名称列表（保留顺序）。"""
    return list(_get_work_types_dict(cfg).keys())


def _find_enum_field(enum_data: dict, table_key: str, field_name: str) -> list | None:
    """在 fields_enum.yaml 中找到指定表的指定字段的选项列表。"""
    table = enum_data.get(table_key, {})
    return table.get(field_name)


def _find_table_field_options(tables_data: dict, table_key: str, field_name: str) -> tuple:
    """在 tables.yaml 中找到指定表的指定字段，返回 (table_dict, field_dict) 或 (None, None)。"""
    for table in tables_data.get("tables", []):
        if table.get("key") == table_key:
            for field in table.get("fields", []):
                if field.get("name") == field_name:
                    return table, field
    return None, None


def _update_prompt_work_types(types_dict: dict):
    """更新 prompts.yaml 中 work_type_classification 的类型定义列表。
    types_dict: {类型名: 描述}，使用正则精确替换，保留其他内容不变。
    """
    raw = _load_yaml_raw(PROMPTS_PATH)
    if not raw:
        return

    # 构建新的类型定义块（带描述）
    non_other = {k: v for k, v in types_dict.items() if k != "其他"}
    lines = []
    for i, (name, desc) in enumerate(non_other.items(), 1):
        if desc:
            lines.append(f"    {i}. {name}：{desc}")
        else:
            lines.append(f"    {i}. {name}")
    other_desc = types_dict.get("其他", "无法明确归入以上任何一类")
    lines.append(f"    {len(non_other) + 1}. 其他：{other_desc}")
    new_block = "\n".join(lines)

    count = len(non_other) + 1

    # 匹配 "## 工作类型定义（N类）" 到下一个 "##" 之间的编号列表
    pattern = r'(## 工作类型定义（)\d+(类）\n)((?:\s+\d+\..+\n)+)'

    def replacer(m):
        return f"{m.group(1)}{count}{m.group(2)}{new_block}\n"

    new_raw = re.sub(pattern, replacer, raw)
    if new_raw != raw:
        _save_raw(PROMPTS_PATH, new_raw)


def cmd_list_work_types():
    """列出所有工作类型及快捷键、关键词映射。"""
    cfg = _load_yaml(CFG_PATH)
    wt = cfg.get("work_types", {})
    types_dict = _get_work_types_dict(cfg)
    shortcuts = wt.get("shortcuts", {})
    keywords = wt.get("keywords", {})
    team_mapping = cfg.get("team", {}).get("work_type_mapping", {})

    # 反转 shortcuts: type -> [k1, k2, ...]
    type_shortcuts = {}
    for k, v in shortcuts.items():
        type_shortcuts.setdefault(v, []).append(k)

    # 反转 keywords: type -> [kw1, kw2, ...]
    type_keywords = {}
    for k, v in keywords.items():
        type_keywords.setdefault(v, []).append(k)

    print(json.dumps({
        "work_types": [
            {
                "name": name,
                "description": desc,
                "shortcuts": type_shortcuts.get(name, []),
                "keywords": type_keywords.get(name, []),
                "team_mapping": team_mapping.get(name, ""),
            }
            for name, desc in types_dict.items()
        ],
        "total": len(types_dict),
    }, ensure_ascii=False, indent=2))


def cmd_add_work_type(name: str, shortcut: str = None, keywords: str = None,
                      team_mapping: str = None, description: str = None):
    """新增工作类型，同步到所有配置文件。"""
    changes = []

    # 1. kingwork.yaml
    cfg = _load_yaml(CFG_PATH)
    wt = cfg.setdefault("work_types", {})
    types_raw = wt.setdefault("types", {})
    # 兼容旧版 list
    if isinstance(types_raw, list):
        types_raw = {t: "" for t in types_raw}
        wt["types"] = types_raw

    if name in types_raw:
        print(json.dumps({"success": False, "error": f"工作类型「{name}」已存在"}, ensure_ascii=False))
        return

    desc = description or ""
    # 插入到"其他"之前（保留 dict 顺序）
    new_types = {}
    inserted = False
    for k, v in types_raw.items():
        if k == "其他" and not inserted:
            new_types[name] = desc
            inserted = True
        new_types[k] = v
    if not inserted:
        new_types[name] = desc
    wt["types"] = new_types
    types_names = list(new_types.keys())

    if shortcut:
        wt.setdefault("shortcuts", {})[shortcut] = name
    if keywords:
        for kw in keywords.split(","):
            kw = kw.strip()
            if kw:
                wt.setdefault("keywords", {})[kw] = name

    team = cfg.setdefault("team", {})
    mapping = team.setdefault("work_type_mapping", {})
    mapping[name] = team_mapping or "其他"

    _save_yaml(CFG_PATH, cfg)
    changes.append("kingwork.yaml: types, shortcuts, keywords, team_mapping")

    # 2. fields_enum.yaml
    enum_data = _load_yaml(ENUM_PATH)
    diary_types = enum_data.setdefault("diary_records", {}).setdefault("工作类型", [])
    if name not in diary_types:
        if "其他" in diary_types:
            idx = diary_types.index("其他")
            diary_types.insert(idx, name)
        else:
            diary_types.append(name)
    _save_yaml(ENUM_PATH, enum_data)
    changes.append("fields_enum.yaml: diary_records.工作类型")

    # 3. tables.yaml
    tables_data = _load_yaml(TABLES_PATH)
    _, field = _find_table_field_options(tables_data, "diary_records", "工作类型")
    if field is not None:
        options = field.setdefault("options", [])
        if name not in options:
            if "其他" in options:
                idx = options.index("其他")
                options.insert(idx, name)
            else:
                options.append(name)
        _save_yaml(TABLES_PATH, tables_data)
        changes.append("tables.yaml: diary_records.工作类型.options")

    # 4. prompts.yaml
    _update_prompt_work_types(new_types)
    changes.append("prompts.yaml: work_type_classification 类型列表")

    print(json.dumps({
        "success": True,
        "action": "add_work_type",
        "name": name,
        "shortcut": shortcut,
        "keywords": keywords.split(",") if keywords else [],
        "team_mapping": team_mapping or "其他",
        "changes": changes,
    }, ensure_ascii=False, indent=2))


def cmd_rename_work_type(old_name: str, new_name: str):
    """重命名工作类型，同步所有配置文件。"""
    changes = []

    # 1. kingwork.yaml
    cfg = _load_yaml(CFG_PATH)
    wt = cfg.get("work_types", {})
    types_dict = _get_work_types_dict(cfg)
    if old_name not in types_dict:
        print(json.dumps({"success": False, "error": f"工作类型「{old_name}」不存在"}, ensure_ascii=False))
        return
    if new_name in types_dict:
        print(json.dumps({"success": False, "error": f"工作类型「{new_name}」已存在"}, ensure_ascii=False))
        return

    # 替换 key，保留顺序和描述
    new_types = {}
    for k, v in types_dict.items():
        if k == old_name:
            new_types[new_name] = v
        else:
            new_types[k] = v
    wt["types"] = new_types

    # shortcuts
    for k, v in wt.get("shortcuts", {}).items():
        if v == old_name:
            wt["shortcuts"][k] = new_name

    # keywords
    for k, v in list(wt.get("keywords", {}).items()):
        if v == old_name:
            wt["keywords"][k] = new_name

    # team mapping
    team_mapping = cfg.get("team", {}).get("work_type_mapping", {})
    if old_name in team_mapping:
        team_mapping[new_name] = team_mapping.pop(old_name)

    # sync_types
    sync_types = cfg.get("sync", {}).get("sync_types", [])
    if old_name in sync_types:
        sync_types[sync_types.index(old_name)] = new_name

    _save_yaml(CFG_PATH, cfg)
    changes.append("kingwork.yaml: types, shortcuts, keywords, team_mapping, sync_types")

    # 2. fields_enum.yaml
    enum_data = _load_yaml(ENUM_PATH)
    diary_types = enum_data.get("diary_records", {}).get("工作类型", [])
    if old_name in diary_types:
        diary_types[diary_types.index(old_name)] = new_name
    _save_yaml(ENUM_PATH, enum_data)
    changes.append("fields_enum.yaml: diary_records.工作类型")

    # 3. tables.yaml
    tables_data = _load_yaml(TABLES_PATH)
    _, field = _find_table_field_options(tables_data, "diary_records", "工作类型")
    if field is not None:
        options = field.get("options", [])
        if old_name in options:
            options[options.index(old_name)] = new_name
        _save_yaml(TABLES_PATH, tables_data)
        changes.append("tables.yaml: diary_records.工作类型.options")

    # 4. prompts.yaml — 全文替换
    raw = _load_yaml_raw(PROMPTS_PATH)
    if old_name in raw:
        new_raw = raw.replace(old_name, new_name)
        _save_raw(PROMPTS_PATH, new_raw)
        changes.append("prompts.yaml: 全文替换")

    print(json.dumps({
        "success": True,
        "action": "rename_work_type",
        "old_name": old_name,
        "new_name": new_name,
        "changes": changes,
    }, ensure_ascii=False, indent=2))


def cmd_remove_work_type(name: str):
    """删除工作类型，同步所有配置文件。"""
    if name == "其他":
        print(json.dumps({"success": False, "error": "不能删除「其他」类型"}, ensure_ascii=False))
        return

    changes = []

    # 1. kingwork.yaml
    cfg = _load_yaml(CFG_PATH)
    wt = cfg.get("work_types", {})
    types_dict = _get_work_types_dict(cfg)
    if name not in types_dict:
        print(json.dumps({"success": False, "error": f"工作类型「{name}」不存在"}, ensure_ascii=False))
        return

    del types_dict[name]
    wt["types"] = types_dict

    # shortcuts
    removed_shortcut_keys = [k for k, v in wt.get("shortcuts", {}).items() if v == name]
    for k in removed_shortcut_keys:
        del wt["shortcuts"][k]

    # keywords
    removed_keyword_keys = [k for k, v in wt.get("keywords", {}).items() if v == name]
    for k in removed_keyword_keys:
        del wt["keywords"][k]

    # team mapping
    team_mapping = cfg.get("team", {}).get("work_type_mapping", {})
    team_mapping.pop(name, None)

    # sync_types
    sync_types = cfg.get("sync", {}).get("sync_types", [])
    if name in sync_types:
        sync_types.remove(name)

    _save_yaml(CFG_PATH, cfg)
    changes.append("kingwork.yaml: types, shortcuts, keywords, team_mapping, sync_types")

    # 2. fields_enum.yaml
    enum_data = _load_yaml(ENUM_PATH)
    diary_types = enum_data.get("diary_records", {}).get("工作类型", [])
    if name in diary_types:
        diary_types.remove(name)
    _save_yaml(ENUM_PATH, enum_data)
    changes.append("fields_enum.yaml: diary_records.工作类型")

    # 3. tables.yaml
    tables_data = _load_yaml(TABLES_PATH)
    _, field = _find_table_field_options(tables_data, "diary_records", "工作类型")
    if field is not None:
        options = field.get("options", [])
        if name in options:
            options.remove(name)
        _save_yaml(TABLES_PATH, tables_data)
        changes.append("tables.yaml: diary_records.工作类型.options")

    # 4. prompts.yaml
    _update_prompt_work_types(types_dict)
    changes.append("prompts.yaml: work_type_classification 类型列表")

    print(json.dumps({
        "success": True,
        "action": "remove_work_type",
        "name": name,
        "removed_shortcuts": removed_shortcut_keys,
        "removed_keywords": removed_keyword_keys,
        "changes": changes,
    }, ensure_ascii=False, indent=2))


# ============================================================
# 通用枚举字段管理（Step 3）
# ============================================================

def _build_enum_index() -> list:
    """构建所有枚举字段的索引：[{table_key, table_name, field_name, options}, ...]"""
    enum_data = _load_yaml(ENUM_PATH)
    tables_data = _load_yaml(TABLES_PATH)

    # 建立 key -> name 映射
    key_to_name = {}
    for table in tables_data.get("tables", []):
        key_to_name[table.get("key", "")] = table.get("name", "")

    index = []
    for table_key, fields in enum_data.items():
        if not isinstance(fields, dict):
            continue
        table_name = key_to_name.get(table_key, table_key)
        for field_name, options in fields.items():
            if isinstance(options, list):
                index.append({
                    "table_key": table_key,
                    "table_name": table_name,
                    "field_name": field_name,
                    "options": options,
                })
    return index


def _find_enum_by_field_name(field_name: str) -> list:
    """根据字段名模糊查找枚举定义（可能多个表有同名字段）。"""
    index = _build_enum_index()
    return [e for e in index if e["field_name"] == field_name]


def cmd_list_enums():
    """列出所有可配置的枚举字段。"""
    index = _build_enum_index()
    print(json.dumps({
        "enums": [
            {
                "table": f"{e['table_name']}({e['table_key']})",
                "field": e["field_name"],
                "options_count": len(e["options"]),
            }
            for e in index
        ],
        "total": len(index),
    }, ensure_ascii=False, indent=2))


def cmd_list_enum(field_name: str, table_key: str = None):
    """查看某个枚举字段的当前选项值。"""
    matches = _find_enum_by_field_name(field_name)
    if table_key:
        matches = [m for m in matches if m["table_key"] == table_key]
    if not matches:
        print(json.dumps({"success": False, "error": f"找不到枚举字段「{field_name}」"}, ensure_ascii=False))
        return

    print(json.dumps({
        "field_name": field_name,
        "results": [
            {
                "table": f"{m['table_name']}({m['table_key']})",
                "options": m["options"],
            }
            for m in matches
        ],
    }, ensure_ascii=False, indent=2))


def _sync_enum_to_tables_yaml(table_key: str, field_name: str, new_options: list):
    """将枚举变更同步到 tables.yaml。"""
    tables_data = _load_yaml(TABLES_PATH)
    _, field = _find_table_field_options(tables_data, table_key, field_name)
    if field is not None:
        field["options"] = new_options
        _save_yaml(TABLES_PATH, tables_data)
        return True
    return False


def cmd_add_enum(field_name: str, value: str, table_key: str = None):
    """给枚举字段添加一个选项。"""
    enum_data = _load_yaml(ENUM_PATH)
    matches = _find_enum_by_field_name(field_name)
    if table_key:
        matches = [m for m in matches if m["table_key"] == table_key]

    if not matches:
        print(json.dumps({"success": False, "error": f"找不到枚举字段「{field_name}」"}, ensure_ascii=False))
        return

    if len(matches) > 1 and not table_key:
        print(json.dumps({
            "success": False,
            "error": f"字段「{field_name}」在多个表中存在，请用 --table 指定",
            "tables": [m["table_key"] for m in matches],
        }, ensure_ascii=False))
        return

    target = matches[0]
    tk = target["table_key"]
    options = enum_data.get(tk, {}).get(field_name, [])

    if value in options:
        print(json.dumps({"success": False, "error": f"选项「{value}」已存在于 {tk}.{field_name}"}, ensure_ascii=False))
        return

    options.append(value)
    _save_yaml(ENUM_PATH, enum_data)

    changes = [f"fields_enum.yaml: {tk}.{field_name}"]

    if _sync_enum_to_tables_yaml(tk, field_name, options):
        changes.append(f"tables.yaml: {tk}.{field_name}.options")

    print(json.dumps({
        "success": True,
        "action": "add_enum",
        "table": tk,
        "field": field_name,
        "added": value,
        "options_after": options,
        "changes": changes,
    }, ensure_ascii=False, indent=2))


def cmd_remove_enum(field_name: str, value: str, table_key: str = None):
    """删除枚举字段的一个选项。"""
    enum_data = _load_yaml(ENUM_PATH)
    matches = _find_enum_by_field_name(field_name)
    if table_key:
        matches = [m for m in matches if m["table_key"] == table_key]

    if not matches:
        print(json.dumps({"success": False, "error": f"找不到枚举字段「{field_name}」"}, ensure_ascii=False))
        return

    if len(matches) > 1 and not table_key:
        print(json.dumps({
            "success": False,
            "error": f"字段「{field_name}」在多个表中存在，请用 --table 指定",
            "tables": [m["table_key"] for m in matches],
        }, ensure_ascii=False))
        return

    target = matches[0]
    tk = target["table_key"]
    options = enum_data.get(tk, {}).get(field_name, [])

    if value not in options:
        print(json.dumps({"success": False, "error": f"选项「{value}」不存在于 {tk}.{field_name}"}, ensure_ascii=False))
        return

    options.remove(value)
    _save_yaml(ENUM_PATH, enum_data)

    changes = [f"fields_enum.yaml: {tk}.{field_name}"]

    if _sync_enum_to_tables_yaml(tk, field_name, options):
        changes.append(f"tables.yaml: {tk}.{field_name}.options")

    print(json.dumps({
        "success": True,
        "action": "remove_enum",
        "table": tk,
        "field": field_name,
        "removed": value,
        "options_after": options,
        "changes": changes,
    }, ensure_ascii=False, indent=2))


def cmd_rename_enum(field_name: str, old_value: str, new_value: str, table_key: str = None):
    """重命名枚举字段的一个选项。"""
    enum_data = _load_yaml(ENUM_PATH)
    matches = _find_enum_by_field_name(field_name)
    if table_key:
        matches = [m for m in matches if m["table_key"] == table_key]

    if not matches:
        print(json.dumps({"success": False, "error": f"找不到枚举字段「{field_name}」"}, ensure_ascii=False))
        return

    if len(matches) > 1 and not table_key:
        print(json.dumps({
            "success": False,
            "error": f"字段「{field_name}」在多个表中存在，请用 --table 指定",
            "tables": [m["table_key"] for m in matches],
        }, ensure_ascii=False))
        return

    target = matches[0]
    tk = target["table_key"]
    options = enum_data.get(tk, {}).get(field_name, [])

    if old_value not in options:
        print(json.dumps({"success": False, "error": f"选项「{old_value}」不存在于 {tk}.{field_name}"}, ensure_ascii=False))
        return
    if new_value in options:
        print(json.dumps({"success": False, "error": f"选项「{new_value}」已存在于 {tk}.{field_name}"}, ensure_ascii=False))
        return

    options[options.index(old_value)] = new_value
    _save_yaml(ENUM_PATH, enum_data)

    changes = [f"fields_enum.yaml: {tk}.{field_name}"]

    if _sync_enum_to_tables_yaml(tk, field_name, options):
        changes.append(f"tables.yaml: {tk}.{field_name}.options")

    print(json.dumps({
        "success": True,
        "action": "rename_enum",
        "table": tk,
        "field": field_name,
        "old_value": old_value,
        "new_value": new_value,
        "options_after": options,
        "changes": changes,
    }, ensure_ascii=False, indent=2))


# ============================================================
# CLI 入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="KingWork 配置管理")
    sub = parser.add_subparsers(dest="command")

    # 工作类型
    sub.add_parser("list-work-types", help="列出所有工作类型")

    p_add = sub.add_parser("add-work-type", help="新增工作类型")
    p_add.add_argument("name", help="类型名称")
    p_add.add_argument("--desc", dest="description", help="类型描述（用于 LLM 分类 prompt）")
    p_add.add_argument("--shortcut", help="快捷键，如 k10")
    p_add.add_argument("--keywords", help="关联关键词，逗号分隔")
    p_add.add_argument("--team-mapping", help="团队表映射类型")

    p_rename = sub.add_parser("rename-work-type", help="重命名工作类型")
    p_rename.add_argument("old_name", help="旧名称")
    p_rename.add_argument("new_name", help="新名称")

    p_remove = sub.add_parser("remove-work-type", help="删除工作类型")
    p_remove.add_argument("name", help="类型名称")

    # 通用枚举
    sub.add_parser("list-enums", help="列出所有枚举字段")

    p_le = sub.add_parser("list-enum", help="查看枚举字段选项")
    p_le.add_argument("field_name", help="字段名称")
    p_le.add_argument("--table", dest="table_key", help="指定数据表 key")

    p_ae = sub.add_parser("add-enum", help="添加枚举选项")
    p_ae.add_argument("field_name", help="字段名称")
    p_ae.add_argument("value", help="新选项值")
    p_ae.add_argument("--table", dest="table_key", help="指定数据表 key")

    p_re = sub.add_parser("remove-enum", help="删除枚举选项")
    p_re.add_argument("field_name", help="字段名称")
    p_re.add_argument("value", help="要删除的选项值")
    p_re.add_argument("--table", dest="table_key", help="指定数据表 key")

    p_rn = sub.add_parser("rename-enum", help="重命名枚举选项")
    p_rn.add_argument("field_name", help="字段名称")
    p_rn.add_argument("old_value", help="旧选项值")
    p_rn.add_argument("new_value", help="新选项值")
    p_rn.add_argument("--table", dest="table_key", help="指定数据表 key")

    args = parser.parse_args()

    if args.command == "list-work-types":
        cmd_list_work_types()
    elif args.command == "add-work-type":
        cmd_add_work_type(args.name, args.shortcut, args.keywords, args.team_mapping, args.description)
    elif args.command == "rename-work-type":
        cmd_rename_work_type(args.old_name, args.new_name)
    elif args.command == "remove-work-type":
        cmd_remove_work_type(args.name)
    elif args.command == "list-enums":
        cmd_list_enums()
    elif args.command == "list-enum":
        cmd_list_enum(args.field_name, args.table_key)
    elif args.command == "add-enum":
        cmd_add_enum(args.field_name, args.value, args.table_key)
    elif args.command == "remove-enum":
        cmd_remove_enum(args.field_name, args.value, args.table_key)
    elif args.command == "rename-enum":
        cmd_rename_enum(args.field_name, args.old_value, args.new_value, args.table_key)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()