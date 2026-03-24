#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KingWork 多维表初始化脚本。
创建 KingWork.dbt 多维表文件，并初始化11个数据表。

用法:
  python scripts/init_tables.py

环境变量:
  WPS_SID           - WPS 用户凭证（必须）
  WPS365_SKILL_PATH - WPS365 Skill 路径（可选，默认 ../wps365-skill）
  KINGWORK_FILE_ID  - 如已有多维表文件，可跳过创建步骤
"""
import json
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import yaml
from kingwork_client.base import (
    import_wpsv7client, get_wps365_root, load_tables_schema,
    save_config_file_id, save_config_sheet_ids, get_file_id, get_sheet_ids
)


def create_dbt_file(drive_client) -> tuple:
    """在云盘创建 KingWork.dbt 文件，返回 (file_id, link_id)。"""
    import subprocess
    import json
    print("  创建多维表文件 KingWork.dbt ...")
    # 调用wps365-skill的drive模块创建文件，保证参数正确
    wps365_root = get_wps365_root()
    # 第一步：创建文件
    cmd = [
        sys.executable,
        str(wps365_root / "skills" / "drive" / "run.py"),
        "create", "KingWork.dbt",
        "--drive", "private"
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if result.returncode != 0:
        raise RuntimeError(f"创建文件失败：{result.stderr or result.stdout}")
    
    # 第二步：搜索刚创建的文件获取file_id
    cmd_search = [
        sys.executable,
        str(wps365_root / "skills" / "drive" / "run.py"),
        "search", "KingWork.dbt",
        "--type", "file_name",
        "--json"
    ]
    result_search = subprocess.run(cmd_search, capture_output=True, text=True, encoding="utf-8")
    try:
        resp = json.loads(result_search.stdout)
    except json.JSONDecodeError:
        raise RuntimeError(f"获取文件ID失败：{result_search.stderr or result_search.stdout}")
    
    if resp.get("code") != 0:
        raise RuntimeError(f"获取文件ID失败：{resp.get('msg', '未知错误')}")
    items = (resp.get("data") or {}).get("items") or []
    if not items:
        raise RuntimeError(f"未找到刚创建的KingWork.dbt文件")
    
    # 取最新创建的文件
    items.sort(key=lambda x: x.get("ctime", ""), reverse=True)
    data = items[0]
    file_id = data.get("id") or data.get("file_id")
    link_id = data.get("link_id") or ""
    if not file_id:
        raise RuntimeError(f"创建文件返回数据异常：{json.dumps(data, ensure_ascii=False)}")
    print(f"  ✅ 文件创建成功，file_id = {file_id}")
    return file_id, link_id


def _build_fields_payload(fields: list) -> list:
    """将 tables.yaml 中的字段定义转为 WPS API 格式。"""
    color_pool = [
        4283466178, 4281378020, 4294947584,
        4291821568, 4294243690, 4287245286,
        4280327315, 4289573866
    ]
    result = []
    for field in fields:
        field_obj = {"name": field["name"], "field_type": field["field_type"]}
        if field["field_type"] in ("SingleSelect", "MultiSelect") and "options" in field:
            items = []
            for i, opt in enumerate(field["options"]):
                items.append({"value": opt, "color": color_pool[i % len(color_pool)]})
            field_obj["data"] = {"allow_add_item_while_inputting": True, "items": items}
        result.append(field_obj)
    return result


def create_table(file_id: str, table_schema: dict) -> str:
    """创建单个数据表，返回 sheet_id。"""
    from wpsv7client import dbsheet_create_sheet

    name = table_schema["name"]
    fields_def = _build_fields_payload(table_schema.get("fields", []))
    views_def = [{"name": "表格视图", "view_type": "Grid"}]
    for view in table_schema.get("views", []):
        if view["name"] != "表格视图":  # 避免重复
            views_def.append({"name": view["name"], "view_type": view.get("view_type", "Grid")})

    sheet_payload = {
        "name": name,
        "fields": fields_def,
        "views": views_def,
    }

    print(f"  创建数据表「{name}」...")
    resp = dbsheet_create_sheet(
        file_id=file_id,
        name=sheet_payload["name"],
        fields=sheet_payload["fields"],
        views=sheet_payload["views"]
    )

    if resp.get("code") != 0:
        msg = resp.get("msg", "") or ""
        # 重名错误：表已存在，不算失败
        if "已存在" in msg or "exists" in msg.lower() or resp.get("code") == 400010:
            print(f"  ⚠️ 「{name}」已存在，跳过")
            return None
        raise RuntimeError(f"创建表失败：{msg}")

    data = resp.get("data", {})
    # 从 sheet 信息中取 id
    sheet = data.get("sheet") or {}
    sheet_id = sheet.get("id") or data.get("id")
    if not sheet_id:
        # 有时返回格式不同，尝试其他路径
        sheet_id = data.get("sheet_id") or data.get("sheetId")
    if not sheet_id:
        print(f"    ⚠️  未获取到 sheet_id，响应：{json.dumps(data, ensure_ascii=False)[:200]}")
        return None

    print(f"  ✅ 「{name}」创建成功，sheet_id = {sheet_id}")
    return str(sheet_id)


def delete_empty_records(file_id: str, sheet_id: str):
    """删除空记录（新建表会有一条空记录）。"""
    from wpsv7client import dbsheet_list_records, dbsheet_batch_delete_records
    try:
        resp = dbsheet_list_records(file_id, sheet_id, page_size=10)
        if resp.get("code") != 0:
            return
        records = (resp.get("data") or {}).get("records") or []
        empty_ids = []
        for rec in records:
            raw = rec.get("fields")
            is_empty = True
            if raw:
                if isinstance(raw, str):
                    try:
                        fields = json.loads(raw)
                    except Exception:
                        fields = {}
                else:
                    fields = raw
                is_empty = all(not v for v in fields.values())
            if is_empty:
                empty_ids.append(rec["id"])
        if empty_ids:
            dbsheet_batch_delete_records(file_id, sheet_id, empty_ids)
    except Exception:
        pass


# key 名称映射（tables.yaml key → config.yaml key）
TABLE_KEY_MAP = {
    "diary_records": "diary_records",
    "todo_records": "todo_records",
    "customer_profiles": "customer_profiles",
    "project_profiles": "project_profiles",
    "customer_followups": "customer_followups",
    "learning_records": "learning_records",
    "support_records": "support_records",
    "team_records": "team_records",
    "idea_records": "idea_records",
    "surprise_docs": "surprise_docs",
    "surprise_communications": "surprise_communications",
    "surprise_meetings": "surprise_meetings",
}


def persist_file_id_to_bashrc(file_id: str) -> bool:
    """将 KINGWORK_FILE_ID 写入 ~/.bashrc，使其持久化。返回是否成功。"""
    import os
    bashrc_path = Path.home() / ".bashrc"
    marker = 'export KINGWORK_FILE_ID='
    new_line = f'{marker}"{file_id}"'

    try:
        content = bashrc_path.read_text(encoding="utf-8") if bashrc_path.exists() else ""

        # 检查是否已有该变量，逐行处理
        lines = content.splitlines()
        new_lines = []
        found = False
        for line in lines:
            if line.strip().startswith(marker):
                new_lines.append(new_line)
                found = True
            else:
                new_lines.append(line)

        if not found:
            new_lines.append(new_line)

        bashrc_path.write_text("\n".join(new_lines) + "\n", encoding="utf-8")

        # 即时生效：同步到当前进程环境变量
        os.environ["KINGWORK_FILE_ID"] = file_id
        return True
    except Exception as e:
        print(f"  ⚠️  自动写入 ~/.bashrc 失败：{e}")
        return False


def main():
    import argparse
    parser = argparse.ArgumentParser(description='KingWork 多维表初始化脚本')
    parser.add_argument('--file-id', type=str, help='显式指定多维表文件ID，如提供则直接使用该ID初始化数据表')
    parser.add_argument('--sync', action='store_true', help='同步模式：对比现有表结构，只创建缺失的表/字段，不删数据')
    args = parser.parse_args()

    print("\n## KingWork 多维表初始化\n")

    # 加载 WPS365 客户端
    import_wpsv7client()

    file_id = None
    # 优先使用用户显式传入的file_id
    if args.file_id:
        file_id = args.file_id
        print(f"使用用户显式指定的多维表ID：{file_id}")
    else:
        # 用户未显式提供，清空现有配置，自动新建多维表
        print("未显式指定多维表ID，清空现有配置并自动创建全新多维表...")
        import os
        if 'KINGWORK_FILE_ID' in os.environ:
            del os.environ['KINGWORK_FILE_ID']
        # 清空配置文件中的file_id
        save_config_file_id('')

    # 创建多维表文件
    if not file_id:
        print("\n### 第一步：创建多维表文件")
        try:
            file_id, link_id = create_dbt_file(None)
            save_config_file_id(file_id)
            print(f"\n  配置已保存：file_id = {file_id}")
            if link_id:
                print(f"  link_id = {link_id}")
        except Exception as e:
            print(f"\n❌ 创建文件失败：{e}")
            print("\n手动操作：")
            print("  1. 打开 WPS 365 创建一个多维表文件")
            print("  2. 设置环境变量：export KINGWORK_FILE_ID=<file_id>")
            print("  3. 重新运行此脚本")
            sys.exit(1)

    # 加载 Schema 定义
    tables_schema = load_tables_schema()
    if not tables_schema:
        print("❌ 未找到 config/tables.yaml")
        sys.exit(1)

    # 加载已有配置（用于 merge）
    existing_sheet_ids = get_sheet_ids() or {}

    # 同步模式：拉取现有多维表的 schema，对比后只处理差异
    if args.sync and file_id:
        from wpsv7client import dbsheet_get_schema as get_schema, dbsheet_batch_create_fields as batch_create_fields
        print(f"\n### 第二步：同步模式 - 对比现有表结构\n")
        try:
            schema_resp = get_schema(file_id)
            if schema_resp.get("code") != 0:
                print(f"⚠️  获取现有结构失败：{schema_resp.get('msg', '未知错误')}，退化为创建模式")
                args.sync = False
            else:
                existing_sheets = {(s.get("name") or ""): s for s in (schema_resp.get("data") or {}).get("sheets") or []}
                print(f"  现有 {len(existing_sheets)} 张表，将逐个对比...")
        except Exception as e:
            print(f"⚠️  获取现有结构失败：{e}，退化为创建模式")
            args.sync = False
    else:
        existing_sheets = {}

    print(f"\n### 第二步：{'同步' if args.sync else "创建"} {len(tables_schema)} 个数据表\n")

    sheet_id_map = {}
    failed = []

    for table_def in tables_schema:
        key = table_def["key"]
        table_name = table_def["name"]
        try:
            existing_sheet = existing_sheets.get(table_name)
            if existing_sheet:
                # 表已存在，同步字段
                sheet_id = str(existing_sheet.get("id") or existing_sheet.get("sheet_id") or "")
                if not sheet_id:
                    print(f"  ⚠️ 「{table_name}」已存在但无法获取ID，尝试创建新表")
                    sheet_id = create_table(file_id, table_def)
                    if sheet_id:
                        sheet_id_map[key] = sheet_id
                        delete_empty_records(file_id, sheet_id)
                    else:
                        failed.append(table_def["name"])
                    continue

                # 对比字段：找出缺失的
                existing_field_names = {f.get("name") for f in existing_sheet.get("fields") or [] if f.get("name")}
                desired_fields = table_def.get("fields", [])
                missing_fields = [f for f in desired_fields if f.get("name") not in existing_field_names]

                if missing_fields:
                    print(f"  「{table_name}」已存在，补充 {len(missing_fields)} 个缺失字段...")
                    fields_payload = _build_fields_payload(missing_fields)
                    try:
                        resp = batch_create_fields(file_id, int(sheet_id), fields_payload)
                        if resp.get("code") == 0:
                            print(f"    ✅ 补充字段成功：{[f['name'] for f in missing_fields]}")
                        else:
                            print(f"    ⚠️  补充字段失败：{resp.get('msg', '未知错误')}，请手动在WPS中补充")
                    except Exception as e:
                        print(f"    ⚠️  补充字段异常：{e}，请手动在WPS中补充")
                else:
                    print(f"  「{table_name}」已存在，结构一致，跳过")

                sheet_id_map[key] = sheet_id
                # 即使表存在，也清理空记录
                delete_empty_records(file_id, sheet_id)
            else:
                # 表不存在，创建新表
                sheet_id = create_table(file_id, table_def)
                if sheet_id:
                    sheet_id_map[key] = sheet_id
                    delete_empty_records(file_id, sheet_id)
                else:
                    # create_table返回None：表已存在（重名），查schema获取其ID以加入配置
                    from wpsv7client import dbsheet_get_schema as get_schema
                    try:
                        schema_resp = get_schema(file_id)
                        sheets = (schema_resp.get("data") or {}).get("sheets") or []
                        found = False
                        for s in sheets:
                            if (s.get("name") or "") == table_name:
                                sid = str(s.get("id") or s.get("sheet_id") or "")
                                if sid:
                                    sheet_id_map[key] = sid
                                    print(f"  ℹ️  「{table_name}」已存在，已获取ID：{sid}，加入配置")
                                    delete_empty_records(file_id, sid)
                                    found = True
                        if not found:
                            failed.append(table_def["name"])
                    except Exception as e:
                        print(f"  ⚠️  「{table_def['name']}」表已存在但获取ID失败：{e}")
                        failed.append(table_def["name"])
        except Exception as e:
            print(f"  ❌ 「{table_def['name']}」处理失败：{e}")
            failed.append(table_def["name"])

    # 保存 sheet_id 到配置文件（合并已有配置 + 本次新增/更新的）
    merged_ids = {**existing_sheet_ids, **sheet_id_map}
    if sheet_id_map:
        save_config_sheet_ids(merged_ids)
        print(f"\n### 第三步：保存配置\n")
        print(f"  已将 {len(sheet_id_map)} 个数据表 ID 合并保存到 config/kingwork.yaml（共 {len(merged_ids)} 张表）")

    # 提取枚举字段配置并保存
    print(f"\n  提取枚举字段配置...")
    enum_config = {}
    tables_schema = load_tables_schema()
    for table in tables_schema:
        table_key = table["key"]
        enum_fields = {}
        for field in table["fields"]:
            if "options" in field and field["options"]:
                enum_fields[field["name"]] = field["options"]
        if enum_fields:
            enum_config[table_key] = enum_fields
    
    # 保存到config/fields_enum.yaml
    enum_config_path = _root / "config" / "fields_enum.yaml"
    with open(enum_config_path, "w", encoding="utf-8") as f:
        yaml.dump(enum_config, f, allow_unicode=True, default_flow_style=False)
    print(f"  ✅ 已将 {sum(len(v) for v in enum_config.values())} 个枚举字段配置保存到 config/fields_enum.yaml")

    # 汇总
    print(f"\n## 初始化完成\n")
    print(f"  ✅ 成功创建：{len(sheet_id_map)} 个数据表")
    if failed:
        print(f"  ❌ 失败：{len(failed)} 个（{', '.join(failed)}）")
    print(f"\n  多维表 file_id：{file_id}")

    # 自动将 file_id 持久化到 ~/.bashrc
    if persist_file_id_to_bashrc(file_id):
        print(f"  ✅ 已自动写入 ~/.bashrc（下次新终端自动生效）")
        print(f"     当前终端执行以下命令立即生效：")
        print(f"     export KINGWORK_FILE_ID={file_id}")
    else:
        print(f"  ⚠️  请手动执行：export KINGWORK_FILE_ID={file_id}")

    print(f"\n  下一步：")
    print(f"  1. 记录工作日记：python skills/kingrecord/run.py \"今天和客户沟通了需求\"")
    print(f"  2. 查看提醒：python skills/kingalert/run.py")


if __name__ == "__main__":
    main()
