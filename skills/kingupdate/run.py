#!/usr/bin/env python3
"""
kingupdate - 更新 KingWork 多维表记录
模式B：结构化输出，由 agent 自然语言协调用户交互
"""
from __future__ import annotations
import sys
import os
import json
import argparse
from datetime import datetime, timedelta
from pathlib import Path

BROWSE_DIR = os.path.dirname(os.path.abspath(__file__))
SKILLS_DIR = BROWSE_DIR
KINGWORK_ROOT = os.path.dirname(os.path.dirname(SKILLS_DIR))
sys.path.insert(0, SKILLS_DIR)
sys.path.insert(0, KINGWORK_ROOT)

from kingwork_client.base import get_wps365_root as _get_wps365_root
sys.path.insert(0, str(_get_wps365_root()))
os.chdir(KINGWORK_ROOT)

from wpsv7client import (
    dbsheet_list_records,
    dbsheet_batch_delete_records as batch_delete_records,
    dbsheet_batch_update_records as batch_update_records,
)
import kingwork_client.tables as kt

FILE_ID = "cbMwPNjcGRwD"

# ------------------------------------------------------------------
# 表配置：sheet_id + 字段映射
# ------------------------------------------------------------------
# sheet_id 动态从 tables.sheet_ids 获取，不写死

TABLE_CONFIGS = {
    "todo": {
        "sheet_key": "todo_records",
        "search_fields": ["任务名称", "任务描述", "关联客户", "关联项目"],
        "date_field": "到期时间",
        "display_fields": ["任务名称", "任务描述", "优先级", "关联客户", "关联项目", "到期时间", "状态"],
    },
    "diary": {
        "sheet_key": "diary_records",
        "search_fields": ["内容", "工作类型", "关联客户", "关联项目"],
        "date_field": "记录时间",
        "display_fields": ["工作类型", "关联客户", "关联项目", "内容", "记录时间", "来源"],
    },
    "customer_followup": {
        "sheet_key": "customer_followups",
        "search_fields": ["跟进内容", "关联客户", "跟进时间"],
        "date_field": "跟进时间",
        "display_fields": ["关联客户", "跟进时间", "跟进内容", "来源"],
    },
    "project": {
        "sheet_key": "project_profiles",
        "search_fields": ["项目名称", "关联客户", "项目状态"],
        "date_field": "开始时间",
        "display_fields": ["项目名称", "关联客户", "项目状态", "开始时间"],
    },
}


def _get_sheet_id(sheet_key: str) -> str | None:
    """从 tables.sheet_ids 动态获取 sheet_id。"""
    cfg = TABLE_CONFIGS.get(sheet_key, {})
    real_key = cfg.get("sheet_key", sheet_key)  # TABLE_CONFIGS key → sheet_ids key
    try:
        tables = kt.KingWorkTables()
        sid = tables.sheet_ids.get(real_key)
        return str(sid or "")
    except Exception:
        return None


def _date_str(days_ago: int) -> str:
    """返回 N 天前的日期字符串 YYYY/MM/DD。"""
    d = datetime.now() - timedelta(days=days_ago)
    return d.strftime("%Y/%m/%d")


def _normalize_date(s: str) -> str:
    """统一日期格式为 YYYY/MM/DD。"""
    s = s.strip()
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d"):
        for try_fmt in (fmt,):
            try:
                return datetime.strptime(s[:10], try_fmt).strftime("%Y/%m/%d")
            except Exception:
                pass
    return s[:10] if s else ""


def _in_recent(date_str: str, days: int) -> bool:
    """判断 date_str 是否在近 days 天内。"""
    if not date_str:
        return True  # 无日期不过滤
    try:
        d = datetime.strptime(date_str[:10], "%Y/%m/%d")
        cutoff = datetime.now() - timedelta(days=days)
        return d >= cutoff
    except Exception:
        return True  # 无法解析日期，不过滤


def _get_display_text(record: dict, cfg: dict) -> dict:
    """从记录中提取用于展示的字段。"""
    fields = record.get("fields") or {}
    if isinstance(fields, str):
        try:
            fields = json.loads(fields)
        except Exception:
            fields = {}
    result = {}
    for f in cfg.get("display_fields", []):
        result[f] = fields.get(f, "")
    return result


def _match_score(record: dict, cfg: dict, query: str) -> float:
    """计算单条记录与 query 的匹配度。"""
    fields = record.get("fields") or {}
    if isinstance(fields, str):
        try:
            fields = json.loads(fields)
        except Exception:
            fields = {}
    q_lower = query.lower()
    scores = []
    for f in cfg.get("search_fields", []):
        v = fields.get(f, "")
        v_str = str(v).lower()
        if q_lower in v_str:
            scores.append(len(q_lower) / max(len(v_str), 1))
    return max(scores) if scores else 0.0


def _search_table(sheet_key: str, query: str, days: int) -> list[dict]:
    """
    在单张表中搜索近 N 天记录。
    优先使用服务端日期筛选，失败降级客户端过滤。
    返回匹配结果列表。
    """
    cfg = TABLE_CONFIGS.get(sheet_key, {})
    sheet_id = _get_sheet_id(sheet_key)
    if not sheet_id:
        return []

    date_field = cfg.get("date_field", "")

    # 注：WPS Date 字段不支持服务端 filter（GreaterThan/LessThan 返回 Unknown enum），
    #     全量拉取 + 客户端日期过滤

    try:
        resp = dbsheet_list_records(
            file_id=FILE_ID,
            sheet_id=int(sheet_id),
            page_size=100,
        )
        if resp.get("code") != 0:
            return []
    except Exception:
        return []

    raw_records = (resp.get("data") or {}).get("records", [])
    results = []

    for rec in raw_records:
        fields = rec.get("fields") or {}
        if isinstance(fields, str):
            try:
                fields = json.loads(fields)
            except Exception:
                continue

        # 客户端日期过滤
        date_val = fields.get(date_field, "")
        if date_val and not _in_recent(_normalize_date(str(date_val)), days):
            continue

        # 关键词匹配
        score = _match_score(rec, cfg, query)
        if score < 0.05:
            continue

        display = _get_display_text(rec, cfg)
        record_id = rec.get("id", "")

        # 构建简洁摘要（取内容字段前60字）
        content = display.get("内容") or display.get("跟进内容") or display.get("项目名称") or ""
        snippet = str(content)[:60].replace("\n", " ").strip()

        results.append({
            "idx": 0,  # 后续填充
            "sheet_key": sheet_key,
            "sheet_name": _sheet_key_to_name(sheet_key),
            "sheet_id": sheet_id,
            "record_id": record_id,
            "display": display,
            "snippet": snippet,
            "score": score,
        })

    return results


def _sheet_key_to_name(key: str) -> str:
    names = {
        "todo": "待办记录",
        "diary": "日记记录",
        "customer_followup": "客户跟进记录",
        "project": "项目档案",
    }
    return names.get(key, key)


def _build_display(rec: dict) -> str:
    """把一条候选记录转成易读的自然语言描述。"""
    d = rec["display"]
    parts = []
    # 类型/项目名
    if d.get("工作类型"):
        parts.append(f"[{d['工作类型']}]")
    if d.get("项目名称"):
        parts.append(f"项目：{d['项目名称']}")
    # 客户
    if d.get("关联客户"):
        parts.append(f"客户：{d['关联客户']}")
    # 内容摘要
    if rec["snippet"]:
        parts.append(rec["snippet"])
    # 日期
    date = d.get("记录时间") or d.get("跟进时间") or d.get("开始时间") or ""
    if date:
        parts.append(f"📅 {date}")
    return " | ".join(parts)


# ------------------------------------------------------------------
# 命令实现
# ------------------------------------------------------------------

def cmd_search(args) -> dict:
    """搜索近 N 天匹配关键词的记录。"""
    days = args.days or 3
    query = args.query.strip()

    all_candidates = []
    for sheet_key in TABLE_CONFIGS:
        results = _search_table(sheet_key, query, days)
        all_candidates.extend(results)

    # 按匹配度降序
    all_candidates.sort(key=lambda x: x["score"], reverse=True)

    # 填充序号
    for i, r in enumerate(all_candidates, 1):
        r["idx"] = i

    # 转换为输出格式
    candidates_out = []
    for r in all_candidates:
        d = r["display"]
        candidates_out.append({
            "idx": r["idx"],
            "sheet_name": r["sheet_name"],
            "sheet_id": r["sheet_id"],
            "record_id": r["record_id"],
            "type": d.get("工作类型", ""),
            "customer": d.get("关联客户", ""),
            "project": d.get("项目名称", ""),
            "content": r["snippet"],
            "date": d.get("记录时间") or d.get("跟进时间") or d.get("开始时间", ""),
            "source": d.get("来源", ""),
        })

    result = {
        "candidates": candidates_out,
        "needs_selection": len(candidates_out) > 1,
        "query": query,
        "total": len(candidates_out),
    }

    # 打印结构化结果（供 agent 解析）
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def cmd_complete(args) -> dict:
    """标记待办为已完成。"""
    record_id = args.record_id
    sheet_id = args.sheet_id or _get_sheet_id("todo")
    if not sheet_id:
        result = {"success": False, "message": "无法获取待办记录的 sheet_id"}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return result

    now = datetime.now().strftime("%Y/%m/%d")
    records_payload = [{
        "id": record_id,
        "fields_value": json.dumps({"状态": "已完成", "完成时间": now}, ensure_ascii=False),
    }]

    try:
        resp = batch_update_records(
            file_id=FILE_ID,
            sheet_id=int(sheet_id),
            records=records_payload,
        )
        success = resp.get("code") == 0
        result = {
            "success": success,
            "record_id": record_id,
            "sheet_id": sheet_id,
            "message": "✅ 已标记为完成" if success else f"更新失败：{resp.get('msg', '')}",
        }
    except Exception as e:
        result = {"success": False, "record_id": record_id, "sheet_id": sheet_id, "message": f"更新异常：{e}"}

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def cmd_delete(args) -> dict:
    """删除指定记录。"""
    record_id = args.record_id
    sheet_id = args.sheet_id

    try:
        resp = batch_delete_records(
            file_id=FILE_ID,
            sheet_id=int(sheet_id),
            record_ids=[record_id],
        )
        success = resp.get("code") == 0
        result = {
            "success": success,
            "record_id": record_id,
            "sheet_id": sheet_id,
            "message": "删除成功" if success else f"删除失败：{resp.get('msg', '')}",
        }
    except Exception as e:
        result = {
            "success": False,
            "record_id": record_id,
            "sheet_id": sheet_id,
            "message": f"删除异常：{e}",
        }

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def cmd_update(args) -> dict:
    """更新指定记录的字段。"""
    record_id = args.record_id
    sheet_id = args.sheet_id

    # 只传有值的字段
    fields_to_update = {}
    if args.type:
        fields_to_update["工作类型"] = args.type
    if args.customer:
        fields_to_update["关联客户"] = args.customer
    if args.project:
        fields_to_update["关联项目"] = args.project
    if args.content:
        fields_to_update["内容"] = args.content

    if not fields_to_update:
        result = {"success": False, "message": "没有要更新的字段"}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return result

    records_payload = [{
        "id": record_id,
        "fields_value": json.dumps(fields_to_update, ensure_ascii=False),
    }]

    try:
        resp = batch_update_records(
            file_id=FILE_ID,
            sheet_id=int(sheet_id),
            records=records_payload,
        )
        success = resp.get("code") == 0
        result = {
            "success": success,
            "record_id": record_id,
            "sheet_id": sheet_id,
            "updated_fields": list(fields_to_update.keys()),
            "message": "更新成功" if success else f"更新失败：{resp.get('msg', '')}",
        }
    except Exception as e:
        result = {
            "success": False,
            "record_id": record_id,
            "sheet_id": sheet_id,
            "message": f"更新异常：{e}",
        }

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


# ------------------------------------------------------------------
# CLI 入口
# ------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="kingupdate - 更新记录")
    sub = parser.add_subparsers(dest="cmd", required=True)

    # search
    p_search = sub.add_parser("search", help="搜索候选记录")
    p_search.add_argument("query", help="描述关键词")
    p_search.add_argument("--days", type=int, default=3, help="搜索近N天（默认3）")

    # complete
    p_comp = sub.add_parser("complete", help="标记待办为已完成")
    p_comp.add_argument("record_id", help="记录ID")
    p_comp.add_argument("--sheet-id", dest="sheet_id", default="", help="sheet_id（可不填，自动用待办表）")

    # delete
    p_del = sub.add_parser("delete", help="删除记录")
    p_del.add_argument("record_id", help="记录ID")
    p_del.add_argument("sheet_id", help="sheet_id")

    # update
    p_upd = sub.add_parser("update", help="更新记录")
    p_upd.add_argument("record_id", help="记录ID")
    p_upd.add_argument("sheet_id", help="sheet_id")
    p_upd.add_argument("--type", default="", help="工作类型")
    p_upd.add_argument("--customer", default="", help="关联客户")
    p_upd.add_argument("--project", default="", help="关联项目")
    p_upd.add_argument("--content", default="", help="内容")

    args = parser.parse_args()

    if args.cmd == "search":
        cmd_search(args)
    elif args.cmd == "complete":
        cmd_complete(args)
    elif args.cmd == "delete":
        cmd_delete(args)
    elif args.cmd == "update":
        cmd_update(args)


if __name__ == "__main__":
    main()
