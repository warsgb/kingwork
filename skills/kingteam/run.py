#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kingteam - 团队数据同步
将个人 KingWork 数据同步至团队：客户跟进记录 + 日报/周报
"""
import sys
import os
import json
import argparse
import subprocess
from datetime import datetime, timedelta, timezone
from pathlib import Path

BROWSE_DIR = Path(__file__).resolve().parent
KINGWORK_ROOT = BROWSE_DIR.parent.parent  # /root/.openclaw/skills/kingwork
sys.path.insert(0, str(KINGWORK_ROOT))           # kingwork_client 在这里
sys.path.insert(0, "/root/.openclaw/skills/wps365-skill")
os.chdir(KINGWORK_ROOT)

import yaml
import requests
from kingwork_client.llm import LLMClient

TZ_CST = timezone(timedelta(hours=8))

# ─── 配置加载 ─────────────────────────────────────────────────

def load_config() -> dict:
    # 读取 kingteam 自身配置
    cfg_path = BROWSE_DIR / "config.yaml"
    with open(cfg_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    # 合并 kingwork.yaml 中的配置
    kingwork_cfg_path = KINGWORK_ROOT / "config" / "kingwork.yaml"
    if kingwork_cfg_path.exists():
        with open(kingwork_cfg_path, encoding="utf-8") as f:
            kw = yaml.safe_load(f)
            for key in ["team", "personal"]:
                if key in kw:
                    if key not in cfg:
                        cfg[key] = {}
                    cfg[key].update(kw[key])
    return cfg

def today_str() -> str:
    return datetime.now(tz=TZ_CST).strftime("%Y/%m/%d")

def today_iso() -> str:
    return datetime.now(tz=TZ_CST).strftime("%Y-%m-%d")

# ─── WPS API 辅助 ──────────────────────────────────────────────

def wps_headers() -> dict:
    sid = os.environ.get("WPS_SID", "")
    return {"cookie": f"wps_sid={sid}", "Content-Type": "application/json"}

def list_records(file_id: str, sheet_id: str, page_size: int = 100) -> list:
    url = f"https://api.wps.cn/v7/dbsheet/{file_id}/sheets/{sheet_id}/records"
    resp = requests.post(url, headers=wps_headers(), json={"page_size": page_size}, timeout=15)
    data = resp.json()
    if resp.status_code != 200 or data.get("code") != 0:
        return []
    records = data.get("data", {}).get("records", [])
    result = []
    for r in records:
        fields_str = r.get("fields", "{}")
        if isinstance(fields_str, str):
            try:
                fields = json.loads(fields_str)
            except Exception:
                fields = {}
        else:
            fields = fields_str
        result.append({"id": r.get("id"), **fields})
    return result

def create_record(file_id: str, sheet_id: str, fields: dict) -> dict:
    url = f"https://api.wps.cn/v7/dbsheet/{file_id}/sheets/{sheet_id}/records/batch_create"
    payload = {
        "records": [{"fields_value": json.dumps(fields, ensure_ascii=False)}]
    }
    resp = requests.post(url, headers=wps_headers(), json=payload, timeout=15)
    data = resp.json()
    return data

def get_sheet_schema(file_id: str, sheet_id: str) -> dict:
    """调用 dbsheet 技能获取多维表结构"""
    try:
        cmd = [
            sys.executable,
            "/root/.openclaw/skills/wps365-skill/skills/dbsheet/run.py",
            "schema",
            file_id,
            "--sheet-id", str(sheet_id),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=20, cwd="/root/.openclaw/skills/wps365-skill")
        output = result.stdout
        # 提取 JSON 部分
        import re
        m = re.search(r"```json\n(.*?)\n```", output, re.DOTALL)
        if m:
            return json.loads(m.group(1))
        return {}
    except Exception as e:
        return {}

def find_folder_id_by_name(drive_id: str, folder_name: str) -> str | None:
    """
    在指定 drive 的根目录按名称查找文件夹，返回其 ID。
    """
    wps_skill = Path(__file__).parent.parent.parent.parent / "wps365-skill"
    try:
        result = subprocess.run(
            [sys.executable, str(wps_skill / "skills" / "drive" / "run.py"),
             "list", "--drive", drive_id, "--parent", "0"],
            capture_output=True, text=True, timeout=20
        )
        output = result.stdout
        # 解析 JSON 获取文件列表
        import re, json
        m = re.search(r'```json\n(.*?)\n```', output, re.DOTALL)
        if not m:
            return None
        data = json.loads(m.group(1))
        for item in data.get("items", []):
            if item.get("type") == "folder" and item.get("name") == folder_name:
                return item.get("id")
        return None
    except Exception:
        return None


def upload_doc_to_folder(drive_id: str, folder_id: str, filename: str, content: str) -> dict:
    """
    通过 drive skill 上传 Markdown 文件到指定 drive 的目标文件夹。
    使用 --drive 和 --parent-id 精确定位。
    """
    wps_skill = Path(__file__).parent.parent.parent.parent / "wps365-skill"
    import re

    # 1. 直接在目标文件夹内创建文件
    name = filename if filename.endswith(".docx") else filename + ".docx"
    create_cmd = [
        sys.executable, str(wps_skill / "skills" / "drive" / "run.py"),
        "create", name,
        "--drive", drive_id,
        "--parent-id", folder_id,
    ]
    try:
        create_result = subprocess.run(create_cmd, capture_output=True, text=True, timeout=20)
        create_output = create_result.stdout
    except Exception as e:
        return {"code": -1, "msg": str(e)}

    # 2. 解析 file_id 和 link_url
    file_id = None
    link_url = None
    for line in create_output.split("\n"):
        m = re.search(r'文件 ID.*?`([a-zA-Z0-9]{20,})`', line)
        if m and not file_id:
            file_id = m.group(1)
        m2 = re.search(r'https?://[^\s]*kdocs\.cn[^\s`\)"\'》]*', line)
        if m2 and not link_url:
            link_url = m2.group(0)

    if not file_id:
        return {"code": -1, "msg": f"创建文件失败，输出：{create_output[:200]}"}

    # 3. 写入内容（通过临时文件）
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".md", delete=False, mode="w", encoding="utf-8") as tmp:
        tmp.write(content)
        tmp_path = tmp.name

    write_cmd = [
        sys.executable, str(wps_skill / "skills" / "drive" / "run.py"),
        "write", file_id, "--file", tmp_path, "--drive", drive_id,
    ]
    try:
        write_result = subprocess.run(write_cmd, capture_output=True, text=True, timeout=30)
        write_output = write_result.stdout
    except Exception as e:
        write_output = str(e)
    finally:
        os.unlink(tmp_path)

    return {
        "code": 0,
        "file_id": file_id,
        "link_url": link_url or f"https://www.kdocs.cn/l/{file_id}",
        "write_output": (write_output or "")[:200]
    }

# ─── LLM 字段映射 ─────────────────────────────────────────────

TEAM_ALLOWED_FIELDS = {"填写日期", "工作类型", "工作内容"}

def find_team_project_by_customer(customer_name: str, team_file_id: str) -> str | None:
    """
    根据客户名称，在团队项目档案表（销售售前项目基础信息表）模糊匹配项目ID。
    匹配逻辑：客户名称是团队表中项目关联客户的子串，或反向匹配。
    返回匹配到的第一个项目ID，匹配不到返回 None。
    """
    if not customer_name:
        return None
    project_sheet_id = "3"  # 销售售前项目基础信息表
    records = list_records(team_file_id, project_sheet_id, page_size=100)
    for r in records:
        linked_customer = str(r.get("客户名称", ""))
        project_name = str(r.get("项目名称", ""))
        # 模糊匹配：客户名是项目关联客户的子串，或反过来
        if customer_name in linked_customer or linked_customer in customer_name:
            return r.get("id")
    return None

def map_record_via_llm(record: dict, team_work_type_options: list, field_mapping: dict,
                        work_type_mapping: dict) -> dict:
    """
    用 LLM 将个人记录映射为团队记录，基于实际字段选项动态匹配。
    team_work_type_options: 团队表工作类型的实际可选项列表
    field_mapping: 配置的字段映射规则（个人字段 → 团队字段）
    """
    llm = LLMClient()
    personal_json = json.dumps(record, ensure_ascii=False, indent=2)
    mapping_hint = json.dumps(work_type_mapping, ensure_ascii=False, indent=2) if work_type_mapping else "无"
    prompt = f"""你是一个字段映射助手。请将左侧「个人工作记录」映射为右侧「团队日报表」字段。

### 个人工作记录
{personal_json}

### 团队表工作类型可选值（仅能从以下列表选择，不许自定义）：
{json.dumps(team_work_type_options, ensure_ascii=False, indent=2)}

### kingwork类型 → 团队类型 参考映射（优先按此映射，LLM可根据内容做更智能的判断）：
{mapping_hint}

### 字段映射规则
1. 工作类型：优先按上述参考映射转换；如果内容明显更匹配其他选项可调整，但必须是上述列表中的值
2. 工作内容：直接使用个人记录的「{field_mapping.get('内容', '内容')}」字段
3. 填写日期：使用个人记录的「{field_mapping.get('记录时间', '记录时间')}」字段，格式统一为 YYYY/MM/DD
4. 只输出团队表存在的字段：填写日期、工作类型、工作内容

请严格仅输出以下JSON格式，不要任何额外文字：
{{"fields": {{
    "工作类型": "从团队列表选的工作类型",
    "工作内容": "工作内容",
    "填写日期": "YYYY/MM/DD"
}}}}
"""
    try:
        result = llm._call(prompt, require_json=True)
        if result and isinstance(result, dict) and "fields" in result:
            return result["fields"]
    except Exception as e:
        pass
    # LLM 失败兜底
    return {
        "工作类型": "其他",
        "工作内容": record.get(field_mapping.get("内容", "内容"), ""),
        "填写日期": record.get(field_mapping.get("记录时间", "记录时间"), "").replace("/", "-"),
    }

# ─── 客户跟进同步 ─────────────────────────────────────────────

def sync_customers(start_date: str, end_date: str, cfg: dict) -> dict:
    """读取个人日记中周期内的客户跟进记录，同步到团队日报表"""
    personal_file_id = cfg["personal"]["dbsheet_id"]
    team_file_id = cfg["team"]["dbsheet_id"]
    team_sheet_id = cfg["team"]["sheet_id"]
    field_mapping = cfg.get("team", {}).get("field_mapping", {})
    conflict_strategy = cfg.get("sync", {}).get("conflict_strategy", "skip")

    # 1. 获取团队表结构，提取工作类型可选值
    team_schema = get_sheet_schema(team_file_id, team_sheet_id)
    work_type_options = []
    for field in team_schema.get("fields", []):
        if field.get("name") == "工作类型" and field.get("type") == "SingleSelect":
            work_type_options = [opt.get("text", "") for opt in field.get("options", [])]
            break
    if not work_type_options:
        # 没有获取到可选值，使用默认兜底列表
        work_type_options = ["客户交流（线上）", "客户交流（线下）", "方案编写", "POC 支持", "培训学习", "其他"]

    # 2. 读取个人日记记录
    records = list_records(personal_file_id, "2", page_size=100)

    # 3. 筛选日期范围内 + 工作类型在配置的可同步类型列表中
    sync_types = cfg.get("sync", {}).get("sync_types", ["客户跟进"])

    # 先查团队已有记录（按填写日期去重）
    team_records = list_records(team_file_id, team_sheet_id, page_size=100)
    existing_dates = set()
    for r in team_records:
        date_val = r.get("填写日期", "")
        if date_val:
            existing_dates.add(str(date_val)[:10].replace("/", "-"))

    # 筛选个人记录（按配置的可同步类型过滤）
    matched = []
    for r in records:
        rec_date = str(r.get(field_mapping.get("记录时间", "记录时间"), ""))[:10].replace("/", "-")
        work_type = str(r.get(field_mapping.get("工作类型", "工作类型"), ""))
        if rec_date >= start_date and rec_date <= end_date and work_type in sync_types:
            matched.append(r)

    if not matched:
        return {"synced": 0, "skipped": 0, "failed": 0, "records": []}

    synced = 0
    skipped = 0
    failed = 0
    results = []

    for rec in matched:
        rec_date = str(rec.get(field_mapping.get("记录时间", "记录时间"), ""))[:10].replace("/", "-")
        # 冲突处理
        if rec_date in existing_dates and conflict_strategy == "skip":
            skipped += 1
            results.append({"id": rec.get("id"), "date": rec_date, "status": "skipped"})
            continue

        # LLM 智能映射
        mapped = map_record_via_llm(rec, work_type_options, field_mapping, cfg.get("team", {}).get("work_type_mapping", {}))
        if not mapped or not mapped.get("工作内容"):
            failed += 1
            results.append({"id": rec.get("id"), "date": rec_date, "status": "failed"})
            continue

        # 写入团队（只写入团队表存在的字段）
        safe_mapped = {k: v for k, v in mapped.items() if k in TEAM_ALLOWED_FIELDS}

        # 如果是客户跟进类型，尝试匹配团队项目档案的 项目选择
        personal_work_type = str(rec.get(field_mapping.get("工作类型", "工作类型"), ""))
        if personal_work_type == "客户跟进":
            customer = rec.get(field_mapping.get("关联客户", "关联客户"), "")
            if customer:
                project_id = find_team_project_by_customer(customer, team_file_id)
                if project_id:
                    safe_mapped["项目选择"] = [project_id]

        resp = create_record(team_file_id, team_sheet_id, safe_mapped)
        if resp.get("code") == 0:
            synced += 1
            results.append({"id": rec.get("id"), "date": rec_date, "status": "synced", "mapped": mapped})
        else:
            failed += 1
            results.append({"id": rec.get("id"), "date": rec_date, "status": "failed", "error": resp.get("msg")})

        existing_dates.add(rec_date)

    return {
        "synced": synced,
        "skipped": skipped,
        "failed": failed,
        "total_matched": len(matched),
        "work_type_options": work_type_options,
        "records": results
    }

# ─── 报告生成（调用 kingreflect） ──────────────────────────────

def generate_report(report_type: str, date_str: str) -> str:
    """调用 kingreflect 生成报告内容，输出到临时文件再读取，避免日志混入 stdout"""
    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix=".md", delete=False, mode="w", encoding="utf-8")
    tmp_path = tmp.name
    tmp.close()

    cmd = [
        sys.executable,
        str(KINGWORK_ROOT / "skills" / "kingreflect" / "run.py"),
        "--period", report_type,
        "--start", date_str,
        "--end", date_str,
        "--output", tmp_path,
    ]
    if report_type == "weekly":
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        monday = dt - timedelta(days=dt.weekday())
        sunday = monday + timedelta(days=6)
        cmd[cmd.index("--start") + 1] = monday.strftime("%Y-%m-%d")
        cmd[cmd.index("--end") + 1] = sunday.strftime("%Y-%m-%d")

    try:
        subprocess.run(cmd, capture_output=True, timeout=60, cwd=str(KINGWORK_ROOT))
        with open(tmp_path, encoding="utf-8") as f:
            content = f.read().strip()
        return content
    except Exception:
        return ""
    finally:
        os.unlink(tmp_path)

# ─── 日报同步 ─────────────────────────────────────────────────

def sync_daily(cfg: dict) -> dict:
    date_str = today_iso()
    report_content = generate_report("daily", date_str)
    if not report_content:
        return {"success": False, "message": "报告生成失败"}

    # 按名称查找日报文件夹 ID
    drive_id = cfg["team"]["drive_id"]
    folder_name = cfg["team"].get("daily_folder", "日报")
    folder_id = find_folder_id_by_name(drive_id, folder_name)
    if not folder_id:
        return {"success": False, "message": f"未找到日报文件夹「{folder_name}」，请检查配置"}

    filename = f"日报_{date_str}.md"
    resp = upload_doc_to_folder(
        drive_id,
        folder_id,
        filename,
        report_content
    )
    return {
        "success": resp.get("code") == 0,
        "filename": filename,
        "folder_url": f"https://365.kdocs.cn/ent/41000207/{drive_id}",
        "file_url": resp.get("link_url", ""),
        "report_content": report_content[:200] + "..." if len(report_content) > 200 else report_content
    }

# ─── 周报同步 ─────────────────────────────────────────────────

def sync_weekly(cfg: dict) -> dict:
    date_str = today_iso()
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    monday = dt - timedelta(days=dt.weekday())
    sunday = monday + timedelta(days=6)
    week_range = f"{monday.strftime('%Y-%m-%d')} ~ {sunday.strftime('%Y-%m-%d')}"

    report_content = generate_report("weekly", date_str)
    if not report_content:
        return {"success": False, "message": "报告生成失败"}

    filename = f"周报_{week_range}.md"
    resp = upload_doc_to_folder(
        cfg["team"]["drive_id"],
        find_folder_id_by_name(cfg["team"]["drive_id"], cfg["team"].get("weekly_folder", "周报")),
        filename,
        report_content
    )
    return {
        "success": resp.get("code") == 0,
        "filename": filename,
        "folder_url": f"https://365.kdocs.cn/ent/41000207/{cfg['team']['drive_id']}",
        "week_range": week_range,
        "report_content": report_content[:200] + "..." if len(report_content) > 200 else report_content
    }

# ─── CLI ─────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="kingteam - 团队数据同步")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_sync = sub.add_parser("sync_customers", help="同步客户跟进记录")
    p_sync.add_argument("--start", default=today_iso(), help="开始日期 YYYY-MM-DD")
    p_sync.add_argument("--end", default=today_iso(), help="结束日期 YYYY-MM-DD")

    p_daily = sub.add_parser("sync_daily", help="生成并上传日报")
    p_weekly = sub.add_parser("sync_weekly", help="生成并上传周报")
    p_all = sub.add_parser("sync_all", help="同步客户跟进 + 日报")

    args = parser.parse_args()
    cfg = load_config()

    if args.cmd == "sync_customers":
        result = sync_customers(args.start, args.end, cfg)
        print(json.dumps(result, ensure_ascii=False, indent=2))

    elif args.cmd == "sync_daily":
        result = sync_daily(cfg)
        print(json.dumps(result, ensure_ascii=False, indent=2))

    elif args.cmd == "sync_weekly":
        result = sync_weekly(cfg)
        print(json.dumps(result, ensure_ascii=False, indent=2))

    elif args.cmd == "sync_all":
        date_str = today_iso()
        r1 = sync_customers(date_str, date_str, cfg)
        r2 = sync_daily(cfg)
        print(json.dumps({"customers": r1, "daily": r2}, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
