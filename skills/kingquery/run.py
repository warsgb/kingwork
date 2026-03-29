#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kingquery - 数据查询和统计
查询多维表数据，提供统计和分析功能。

用法:
  python skills/kingquery/run.py stats --type customer_followup --period 30d
  python skills/kingquery/run.py timeline --customer "某某公司"
  python skills/kingquery/run.py customers
  python skills/kingquery/run.py recent --days 7
  python skills/kingquery/run.py search "关键词"
"""
import argparse
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

_root = Path(__file__).resolve().parent.parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from kingwork_client.base import today_str, debug_log, print_exec_summary
from kingwork_client.tables import KingWorkTables

TZ_CST = timezone(timedelta(hours=8))


def parse_args():
    parser = argparse.ArgumentParser(description="数据查询和统计")
    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # stats 子命令
    stats_parser = subparsers.add_parser("stats", help="统计分析")
    stats_parser.add_argument("--type", choices=["customer_followup", "todo", "learning", "all"],
                              default="all", help="统计类型")
    stats_parser.add_argument("--period", default="30d", help="时间周期（如：7d, 30d, 90d）")

    # timeline 子命令
    tl_parser = subparsers.add_parser("timeline", help="客户沟通时间线")
    tl_parser.add_argument("--customer", required=True, help="客户名称")
    tl_parser.add_argument("--days", type=int, default=90, help="最近N天")

    # customers 子命令
    subparsers.add_parser("customers", help="客户列表")

    # projects 子命令
    subparsers.add_parser("projects", help="项目列表")

    # recent 子命令
    recent_parser = subparsers.add_parser("recent", help="最近工作记录")
    recent_parser.add_argument("--days", type=int, default=7, help="最近N天")
    recent_parser.add_argument("--type", help="工作类型过滤")

    # search 子命令
    search_parser = subparsers.add_parser("search", help="关键词搜索")
    search_parser.add_argument("keyword", help="搜索关键词")

    return parser.parse_args()


def parse_period_days(period_str: str) -> int:
    """解析时间周期字符串，返回天数。"""
    period_str = period_str.lower().strip()
    if period_str.endswith("d"):
        return int(period_str[:-1])
    elif period_str.endswith("w"):
        return int(period_str[:-1]) * 7
    elif period_str.endswith("m"):
        return int(period_str[:-1]) * 30
    return 30


def cmd_stats(tables: KingWorkTables, args):
    """统计分析。"""
    days = parse_period_days(args.period)
    now = datetime.now(tz=TZ_CST)
    start = now - timedelta(days=days)
    start = start.replace(hour=0, minute=0, second=0, microsecond=0)

    print(f"\n## 工作统计 - 最近 {days} 天\n")
    print(f"统计范围：{start.strftime('%Y/%m/%d')} 至 {now.strftime('%Y/%m/%d')}\n")

    if args.type in ("customer_followup", "all"):
        followups = tables.get_records_in_period("customer_followups", "跟进时间", start, now)
        customers = {}
        methods = {}
        for r in followups:
            f = r["fields"]
            c = f.get("客户名称") or "未知"
            m = f.get("跟进方式") or "其他"
            customers[c] = customers.get(c, 0) + 1
            methods[m] = methods.get(m, 0) + 1

        print(f"### 客户跟进统计（共 {len(followups)} 次）\n")
        if customers:
            print("按客户排序：")
            for c, cnt in sorted(customers.items(), key=lambda x: -x[1]):
                print(f"  - {c}：{cnt} 次")
            print()
            print("按沟通方式：")
            for m, cnt in sorted(methods.items(), key=lambda x: -x[1]):
                print(f"  - {m}：{cnt} 次")
            print()

    if args.type in ("learning", "all"):
        learnings = tables.get_records_in_period("learning_records", "学习时间", start, now)
        total_hours = sum((r["fields"].get("学习时长") or 0) for r in learnings)
        topics = [r["fields"].get("学习主题") or "" for r in learnings]

        print(f"### 学习成长统计（共 {len(learnings)} 次，{total_hours} 小时）\n")
        for t in topics[:10]:
            if t:
                print(f"  - {t}")
        print()

    if args.type in ("todo", "all"):
        todos = tables.get_pending_todos()
        completed = []
        try:
            # 服务端过滤：状态 Equals "已完成"
            completed = tables.list_all_records("todo_records", filter_body={
                "criteria": [{"field": "状态", "operator": "equals", "values": ["已完成"]}]
            })
        except Exception:
            pass

        print(f"### 待办事项统计\n")
        print(f"  - 待完成：{len(todos)} 项")
        print(f"  - 已完成：{len(completed)} 项")
        print()


def cmd_timeline(tables: KingWorkTables, args):
    """客户沟通时间线。"""
    now = datetime.now(tz=TZ_CST)
    start = now - timedelta(days=args.days)

    followups = tables.get_records_in_period("customer_followups", "跟进时间", start, now)
    customer_records = [
        r for r in followups
        if (r["fields"].get("客户名称") or "").startswith(args.customer)
        or args.customer in (r["fields"].get("客户名称") or "")
    ]

    print(f"\n## {args.customer} - 沟通时间线（最近 {args.days} 天）\n")

    if not customer_records:
        print(f"未找到「{args.customer}」的跟进记录")
        return

    # 按时间正序
    def get_time(r):
        t = r["fields"].get("跟进时间")
        if not t:
            return datetime.min.replace(tzinfo=TZ_CST)
        try:
            if isinstance(t, (int, float)):
                return datetime.fromtimestamp(t / 1000, tz=TZ_CST)
            from dateutil import parser as dp
            dt = dp.parse(str(t))
            return dt if dt.tzinfo else dt.replace(tzinfo=TZ_CST)
        except Exception:
            return datetime.min.replace(tzinfo=TZ_CST)

    customer_records.sort(key=get_time)

    for r in customer_records:
        f = r["fields"]
        time_str = ""
        t = f.get("跟进时间")
        if t:
            try:
                if isinstance(t, (int, float)):
                    dt = datetime.fromtimestamp(t / 1000, tz=TZ_CST)
                else:
                    from dateutil import parser as dp
                    dt = dp.parse(str(t))
                time_str = dt.strftime("%Y-%m-%d %H:%M")
            except Exception:
                time_str = str(t)[:10]

        method = f.get("跟进方式") or ""
        content = (f.get("跟进内容") or "")[:150]
        result = f.get("跟进结果") or ""

        print(f"**{time_str}** [{method}]")
        if content:
            print(f"  {content}")
        if result:
            print(f"  结果：{result}")
        print()


def cmd_customers(tables: KingWorkTables):
    """客户列表。"""
    records = tables.list_all_records("customer_profiles")
    log_msg = f"查询客户列表，共 {len(records)} 家：\n"
    output_msg = f"\n## 客户档案（共 {len(records)} 家）\n"
    for rec in records:
        f = tables.get_record_fields(rec)
        name = f.get("客户名称") or "未知客户"
        status = f.get("客户状态") or "未知状态"
        follow_time = f.get("最近跟进时间") or "未跟进"
        count = f.get("跟进次数") or 0
        item = f"- {name} [{status}] 最近跟进：{str(follow_time)[:10]} 跟进次数：{count}\n"
        log_msg += item
        output_msg += item
    debug_log(log_msg)
    print(output_msg)

    if not records:
        print("暂无客户记录")
        return

    for rec in records:
        f = tables.get_record_fields(rec)
        name = f.get("客户名称") or "未知"
        c_type = f.get("客户类型") or ""
        status = f.get("客户状态") or ""
        last_time = f.get("最近跟进时间")
        count = f.get("跟进次数") or 0

        last_str = "从未跟进"
        if last_time:
            try:
                if isinstance(last_time, (int, float)):
                    dt = datetime.fromtimestamp(last_time / 1000, tz=TZ_CST)
                else:
                    from dateutil import parser as dp
                    dt = dp.parse(str(last_time))
                last_str = dt.strftime("%Y/%m/%d")
            except Exception:
                last_str = str(last_time)[:10]

        print(f"- **{name}** [{c_type}] {status}")
        print(f"  最近跟进：{last_str}  跟进次数：{count}")


def cmd_projects(tables: KingWorkTables):
    """项目列表。"""
    records = tables.list_all_records("project_profiles")
    print(f"\n## 项目档案（共 {len(records)} 个）\n")

    if not records:
        print("暂无项目记录")
        return

    for rec in records:
        f = tables.get_record_fields(rec)
        name = f.get("项目名称") or "未知"
        status = f.get("项目状态") or ""
        p_type = f.get("项目类型") or ""
        customer = f.get("关联客户") or ""
        priority = f.get("项目优先级") or ""

        print(f"- **{name}** [{status}] {p_type}")
        if customer:
            print(f"  关联客户：{customer}")
        if priority:
            print(f"  优先级：{priority}")


def cmd_recent(tables: KingWorkTables, args):
    """最近工作记录。"""
    now = datetime.now(tz=TZ_CST)
    start = now - timedelta(days=args.days)

    records = tables.get_records_in_period("diary_records", "记录时间", start, now)

    if hasattr(args, "type") and args.type:
        records = [r for r in records if r["fields"].get("工作类型") == args.type]

    log_msg = f"查询最近 {args.days} 天工作记录，共 {len(records)} 条：\n"
    output_msg = f"\n## 最近 {args.days} 天工作记录（共 {len(records)} 条）\n"

    # 按时间倒序
    def get_time(r):
        t = r["fields"].get("记录时间")
        if not t:
            return datetime.min.replace(tzinfo=TZ_CST)
        try:
            if isinstance(t, (int, float)):
                return datetime.fromtimestamp(t / 1000, tz=TZ_CST)
            from dateutil import parser as dp
            dt = dp.parse(str(t))
            return dt if dt.tzinfo else dt.replace(tzinfo=TZ_CST)
        except Exception:
            return datetime.min.replace(tzinfo=TZ_CST)

    records.sort(key=get_time, reverse=True)

    for r in records[:20]:  # 最多显示20条
        f = r["fields"]
        time_str = ""
        t = f.get("记录时间")
        if t:
            try:
                if isinstance(t, (int, float)):
                    dt = datetime.fromtimestamp(t / 1000, tz=TZ_CST)
                else:
                    from dateutil import parser as dp
                    dt = dp.parse(str(t))
                time_str = dt.strftime("%m-%d %H:%M")
            except Exception:
                time_str = str(t)[:10]

        work_type = f.get("工作类型") or ""
        content = (f.get("内容") or "")[:100]
        customer = f.get("关联客户") or ""

        item = f"{time_str} [{work_type}]{' - ' + customer if customer else ''}：{content}\n"
        log_msg += item
        output_msg += f"**{time_str}** [{work_type}]{' - ' + customer if customer else ''}\n  {content}\n\n"
    
    debug_log(log_msg)
    print(output_msg)


def cmd_search(tables: KingWorkTables, keyword: str):
    """关键词搜索。"""
    print(f"\n## 搜索：「{keyword}」\n")

    total = 0

    # 搜索客户档案（服务端 Contains 过滤）
    try:
        matched_customers = tables.list_all_records("customer_profiles", filter_body={
            "criteria": [{"field": "客户名称", "operator": "Contains", "values": [keyword]}]
        })
    except Exception:
        matched_customers = []
    if matched_customers:
        print(f"### 客户档案（{len(matched_customers)} 条）")
        for r in matched_customers:
            f = tables.get_record_fields(r)
            print(f"  - **{f.get('客户名称')}** [{f.get('客户类型', '')}]")
        print()
        total += len(matched_customers)

    # 搜索日记记录（服务端 OR + Contains 过滤）
    try:
        diary_all = tables.list_all_records("diary_records", filter_body={
            "mode": "OR",
            "criteria": [
                {"field": "内容", "operator": "Contains", "values": [keyword]},
                {"field": "关联客户", "operator": "Contains", "values": [keyword]},
            ]
        })
    except Exception:
        diary_all = []
    if diary_all:
        print(f"### 工作日记（{len(diary_all)} 条）")
        for r in diary_all[:5]:
            f = tables.get_record_fields(r)
            content = (f.get("内容") or "")[:100]
            work_type = f.get("工作类型") or ""
            print(f"  - [{work_type}] {content}")
        if len(diary_all) > 5:
            print(f"  ... 还有 {len(diary_all) - 5} 条")
        print()
        total += len(diary_all)

    if total == 0:
        print(f"未找到包含「{keyword}」的记录")
    else:
        print(f"共找到 {total} 条相关记录")


def main():
    from kingwork_client.base import debug_log
    debug_log("当前执行技能：kingquery（数据查询）")
    args = parse_args()

    try:
        tables = KingWorkTables()
    except Exception as e:
        print(f"❌ 初始化失败：{e}", file=sys.stderr)
        sys.exit(1)

    if args.command == "stats":
        cmd_stats(tables, args)
    elif args.command == "timeline":
        cmd_timeline(tables, args)
    elif args.command == "customers":
        cmd_customers(tables)
    elif args.command == "projects":
        cmd_projects(tables)
    elif args.command == "recent":
        cmd_recent(tables, args)
    elif args.command == "search":
        cmd_search(tables, args.keyword)
    else:
        print("请指定子命令。使用 --help 查看帮助。")
        sys.exit(1)
    
    # 输出统一总结
    print_exec_summary([])


if __name__ == "__main__":
    main()
