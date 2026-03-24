#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kingreflect - 日报周报月报生成
汇总工作数据，用大模型生成报告。

用法:
  python skills/kingreflect/run.py --period daily
  python skills/kingreflect/run.py --period weekly
  python skills/kingreflect/run.py --start 2026-03-01 --end 2026-03-15
"""
import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

_root = Path(__file__).resolve().parent.parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from kingwork_client.base import today_str, now_iso, print_exec_summary
from kingwork_client.llm import LLMClient
from kingwork_client.tables import KingWorkTables

TZ_CST = timezone(timedelta(hours=8))


def parse_args():
    parser = argparse.ArgumentParser(description="生成日报周报月报")
    parser.add_argument("natural_language", nargs="?", help="自然语言输入（如：帮我写今天的日报）")
    parser.add_argument("--period", choices=["daily", "weekly", "monthly"],
                        help="报告周期")
    parser.add_argument("--start", help="开始日期 YYYY-MM-DD")
    parser.add_argument("--end", help="结束日期 YYYY-MM-DD")
    parser.add_argument("--output", "-o", help="输出到文件（.md）")
    parser.add_argument("--categories", help="包含类别（逗号分隔，如：customer,learning）")
    return parser.parse_args()


def determine_period(args) -> tuple:
    """确定报告周期，返回 (period_type, start_dt, end_dt, period_str)。"""
    now = datetime.now(tz=TZ_CST)

    if args.start and args.end:
        start = datetime.fromisoformat(args.start).replace(tzinfo=TZ_CST)
        end = datetime.fromisoformat(args.end).replace(tzinfo=TZ_CST)
        end = end.replace(hour=23, minute=59, second=59)
        period_str = f"{args.start} 至 {args.end}"
        return "custom", start, end, period_str

    # 自然语言解析（简单规则）
    if args.natural_language and not args.period:
        text = args.natural_language.lower()
        if any(w in text for w in ["月报", "月", "本月", "上个月"]):
            period = "monthly"
        elif any(w in text for w in ["周报", "周", "本周", "这周", "上周"]):
            period = "weekly"
        else:
            period = "daily"
    else:
        period = args.period or "daily"

    if period == "daily":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = now.replace(hour=23, minute=59, second=59)
        period_str = f"{now.strftime('%Y年%m月%d日')}"

    elif period == "weekly":
        # 本周一到今天
        days_since_monday = now.weekday()
        start = (now - timedelta(days=days_since_monday)).replace(hour=0, minute=0, second=0, microsecond=0)
        end = now.replace(hour=23, minute=59, second=59)
        period_str = f"{start.strftime('%Y年%m月%d日')} 至 {end.strftime('%m月%d日')}"

    else:  # monthly
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        end = now.replace(hour=23, minute=59, second=59)
        period_str = f"{now.strftime('%Y年%m月')}"

    return period, start, end, period_str


def collect_data(tables: KingWorkTables, start_dt: datetime, end_dt: datetime) -> dict:
    """收集指定周期内的工作数据。"""
    data = {}

    # 日记记录（总量统计）
    try:
        diary_records = tables.get_records_in_period("diary_records", "记录时间", start_dt, end_dt)
        data["diary_count"] = len(diary_records)
        # 按类型统计
        type_stats = {}
        for rec in diary_records:
            wt = rec["fields"].get("工作类型") or "未分类"
            type_stats[wt] = type_stats.get(wt, 0) + 1
        data["type_stats"] = type_stats
    except Exception as e:
        data["diary_count"] = 0
        data["type_stats"] = {}

    # 客户跟进记录
    try:
        followups = tables.get_records_in_period("customer_followups", "跟进时间", start_dt, end_dt)
        data["customer_followups"] = [
            {
                "customer": r["fields"].get("客户名称", ""),
                "method": r["fields"].get("跟进方式", ""),
                "content": r["fields"].get("跟进内容", "")[:200],
                "result": r["fields"].get("跟进结果", ""),
                "time": r["fields"].get("跟进时间", ""),
            }
            for r in followups
        ]
    except Exception:
        data["customer_followups"] = []

    # 学习成长记录
    try:
        learnings = tables.get_records_in_period("learning_records", "学习时间", start_dt, end_dt)
        data["learning_records"] = [
            {
                "topic": r["fields"].get("学习主题", ""),
                "type": r["fields"].get("学习类型", ""),
                "takeaway": r["fields"].get("关键收获", "")[:200],
                "duration": r["fields"].get("学习时长", 0),
            }
            for r in learnings
        ]
        data["total_learning_hours"] = sum(
            (r["fields"].get("学习时长") or 0) for r in learnings
        )
    except Exception:
        data["learning_records"] = []
        data["total_learning_hours"] = 0

    # 其他事项（横向支持 + 团队事务）
    other_matters = []
    
    try:
        supports = tables.get_records_in_period("support_records", "支持时间", start_dt, end_dt)
        for r in supports:
            other_matters.append({
                "category": "横向支持",
                "target": r["fields"].get("支持对象", ""),
                "type": r["fields"].get("支持类型", ""),
                "content": r["fields"].get("支持内容", "")[:150],
                "result": r["fields"].get("支持结果", ""),
            })
    except Exception:
        pass

    try:
        team_events = tables.get_records_in_period("team_records", "事务时间", start_dt, end_dt)
        for r in team_events:
            other_matters.append({
                "category": "团队事务",
                "type": r["fields"].get("事务类型", ""),
                "topic": r["fields"].get("事务主题", ""),
                "content": r["fields"].get("事务内容", "")[:150],
            })
    except Exception:
        pass
    
    data["other_matters"] = other_matters

    # 待办事项（未完成）
    try:
        todos = tables.get_pending_todos()
        data["pending_todos"] = [
            {
                "name": r["fields"].get("任务名称", ""),
                "priority": r["fields"].get("优先级", "中"),
                "due": r["fields"].get("到期时间", ""),
                "customer": r["fields"].get("关联客户", ""),
            }
            for r in todos[:10]  # 最多展示10条
        ]
    except Exception:
        data["pending_todos"] = []

    # 惊喜文档
    try:
        surprise_docs = tables.get_records_in_period("surprise_docs", "发现时间", start_dt, end_dt)
        data["surprise_docs"] = [
            {
                "name": r["fields"].get("文档名称", ""),
                "reason": r["fields"].get("惊喜原因", "")[:200],
                "customer": r["fields"].get("相关客户", ""),
            }
            for r in surprise_docs
        ]
    except Exception:
        data["surprise_docs"] = []

    # 惊喜沟通
    try:
        surprise_comms = tables.get_records_in_period("surprise_communications", "沟通时间", start_dt, end_dt)
        data["surprise_communications"] = [
            {
                "person": r["fields"].get("沟通对象", ""),
                "content": r["fields"].get("惊喜内容", "")[:200],
                "value": r["fields"].get("价值点", "")[:150],
            }
            for r in surprise_comms
        ]
    except Exception:
        data["surprise_communications"] = []

    # 灵感记录
    try:
        ideas = tables.get_records_in_period("idea_records", "记录时间", start_dt, end_dt)
        data["idea_records"] = [
            {
                "content": r["fields"].get("灵感内容", "")[:200],
                "category": r["fields"].get("灵感类别", ""),
                "feasibility": r["fields"].get("可行性", ""),
            }
            for r in ideas
        ]
    except Exception:
        data["idea_records"] = []

    return data


def generate_report_prompt(llm: LLMClient, period_type: str, period_str: str, data: dict) -> str:
    """生成报告提示词。"""
    work_data_str = json.dumps(data, ensure_ascii=False, indent=2)
    llm_req = llm.generate_report(period_type, period_str, work_data_str)
    if llm_req and llm_req.get("raw"):
        return llm_req["raw"]
    return None


def generate_fallback_report(period_type: str, period_str: str, data: dict) -> str:
    """当无 AI 响应时，生成基础报告。"""
    type_map = {"daily": "工作日报", "weekly": "工作周报", "monthly": "工作月报"}
    title = type_map.get(period_type, "工作报告")

    lines = [f"# {title} - {period_str}\n"]

    # 工作概览
    lines.append("## 工作概览\n")
    type_stats = data.get("type_stats", {})
    if type_stats:
        for wt, count in type_stats.items():
            lines.append(f"- {wt}：{count} 次")
    else:
        lines.append("- 暂无记录")
    lines.append("")

    # 客户跟进
    followups = data.get("customer_followups", [])
    if followups:
        lines.append(f"## 👥 客户跟进（{len(followups)} 次）\n")
        for f in followups:
            customer = f.get("customer", "未知")
            method = f.get("method", "")
            content = f.get("content", "")[:100]
            result = f.get("result", "")
            lines.append(f"- **{customer}**（{method}）：{content}")
            if result:
                lines.append(f"  - 结果：{result}")
        lines.append("")

    # 学习成长
    learnings = data.get("learning_records", [])
    total_hours = data.get("total_learning_hours", 0)
    hours_str = f"{total_hours} 小时" if total_hours > 0 else "—"
    if learnings:
        lines.append(f"## 📚 学习成长（{len(learnings)} 次，共 {hours_str}）\n")
        for l in learnings:
            topic = l.get("topic", "")
            takeaway = l.get("takeaway", "")[:100]
            lines.append(f"- **{topic}**：{takeaway}")
        lines.append("")

    # 其他事项
    other_matters = data.get("other_matters", [])
    if other_matters:
        lines.append(f"## 📌 其他事项（{len(other_matters)} 项）\n")
        for item in other_matters:
            category = item.get("category", "")
            if category == "横向支持":
                lines.append(f"- **横向支持** | {item.get('target', '')}（{item.get('type', '')}）：{item.get('content', '')}")
            else:
                lines.append(f"- **团队事务** | {item.get('topic', '')}（{item.get('type', '')}）：{item.get('content', '')}")
        lines.append("")

    # 待办事项
    todos = data.get("pending_todos", [])
    if todos:
        lines.append(f"## 📝 待办事项（{len(todos)} 项）\n")
        for t in todos:
            name = t.get("name", "")
            priority = t.get("priority", "中")
            due = t.get("due", "")
            customer = t.get("customer", "")
            due_str = f"（到期：{due[:10]}）" if due else ""
            customer_str = f"（客户：{customer}）" if customer else ""
            lines.append(f"- [{priority}] {name}{due_str}{customer_str}")
        lines.append("")

    return "\n".join(lines)


def main():
    from kingwork_client.base import debug_log
    debug_log("当前执行技能：kingreflect（报告生成）")
    args = parse_args()
    period_type, start_dt, end_dt, period_str = determine_period(args)

    print(f"\n## KingReflect - 生成{period_str}工作报告\n")

    try:
        tables = KingWorkTables()
    except Exception as e:
        print(f"❌ 初始化失败：{e}", file=sys.stderr)
        sys.exit(1)

    # 收集数据
    print("正在收集工作数据...")
    data = collect_data(tables, start_dt, end_dt)

    total_records = data.get("diary_count", 0)
    print(f"共找到 {total_records} 条工作记录\n")

    if total_records == 0:
        print(f"📭 {period_str}暂无工作记录\n")
        print("提示：使用 kingrecord 记录工作日记后，再运行此命令生成报告。")
        return

    # 生成报告
    llm = LLMClient()
    prompt = generate_report_prompt(llm, period_type, period_str, data)

    print("请根据以下提示词和数据生成工作报告（Markdown 格式）：\n")
    print(prompt)
    print("\n---")
    print("报告内容（Markdown）：")
    sys.stdout.flush()

    # 读取 AI 响应
    report_content = ""
    if not sys.stdin.isatty():
        report_content = sys.stdin.read().strip()

    if not report_content:
        # 无 AI 响应时使用基础报告
        print("\n（使用基础模板生成报告）\n")
        report_content = generate_fallback_report(period_type, period_str, data)

    # 输出报告
    print("\n" + "=" * 60)
    print(report_content)
    print("=" * 60 + "\n")

    # 保存到文件
    if args.output:
        output_path = Path(args.output)
        output_path.write_text(report_content, encoding="utf-8")
        print(f"✅ 报告已保存到：{output_path.resolve()}")

    # ── 写入 WPS 云文档（按周期分文件夹）────────────────────────
    # 确定文件夹和文件名
    if period_type == "daily":
        folder_name = "日报"
        date_str = start_dt.strftime("%Y-%m-%d")
        report_filename = f"{date_str}.md"
        title_for_doc = date_str
    elif period_type == "weekly":
        folder_name = "周报"
        # 取本周一和周日
        monday = start_dt.strftime("%Y-%m-%d")
        sunday = end_dt.strftime("%Y-%m-%d")
        week_num = start_dt.isocalendar()[1]
        year = start_dt.year
        report_filename = f"{year}年第{week_num}周({monday}~{sunday}).md"
        title_for_doc = f"{year}年第{week_num}周"
    elif period_type == "monthly":
        folder_name = "月报"
        year_month = start_dt.strftime("%Y-%m")
        report_filename = f"{year_month}.md"
        title_for_doc = year_month
    else:  # custom
        # 按日期范围命名
        date_str = start_dt.strftime("%Y-%m-%d")
        end_str = end_dt.strftime("%Y-%m-%d")
        if start_dt.date() == end_dt.date():
            folder_name = "日报"
            report_filename = f"{date_str}.md"
            title_for_doc = date_str
        else:
            folder_name = "自定义"
            report_filename = f"{date_str}_至_{end_str}.md"
            title_for_doc = f"{date_str} 至 {end_str}"

    # 写入临时文件
    import tempfile
    with tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8") as f:
        f.write(report_content)
        temp_path = f.name

    try:
        sys.path.insert(0, "/root/.openclaw/skills/wps365-skill")
        from wpsv7client.drive import create_otl_document
        from wpsv7client.drive import get_drive_id as _get_drive_id
        from wpsv7client.airpage import write_airpage_content

        drive_id = _get_drive_id("private")
        parent_path = ["我的文档", folder_name]

        # 创建智能文档（自动创建对应文件夹）
        create_resp = create_otl_document(
            drive_id=drive_id,
            file_name=report_filename,
            parent_path=parent_path,
            on_name_conflict="rename",
        )
        if create_resp.get("code") == 0:
            data_result = create_resp.get("data") or {}
            file_id = data_result.get("id", "")
            link_url = data_result.get("link_url", "")

            # 写入报告内容
            write_airpage_content(file_id, title_for_doc, report_content, pos="begin")

            print(f"\n📄 报告已写入 WPS 云文档")
            print(f"   文件夹：{folder_name}")
            print(f"   文件名：{report_filename}")
            print(f"   路径：我的文档/{folder_name}")
            print(f"   链接：{link_url}")
        else:
            print(f"\n⚠️ 云文档创建失败：{create_resp.get('msg', '')}")
    except Exception as e:
        print(f"\n⚠️ 云文档写入异常：{e}")
    finally:
        Path(temp_path).unlink(missing_ok=True)

    # 输出统一总结
    print_exec_summary([])


if __name__ == "__main__":
    main()
