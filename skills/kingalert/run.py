#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kingalert - AI 智能提醒
查询待办事项、未跟进客户和当日日程，生成提醒清单。

用法:
  python skills/kingalert/run.py
  python skills/kingalert/run.py --todos-only
  python skills/kingalert/run.py --calendar-only
  python skills/kingalert/run.py --complete <todo_id>
  python skills/kingalert/run.py --followup <customer_name>
"""
import argparse
import json
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

_root = Path(__file__).resolve().parent.parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from kingwork_client.base import today_str, now_iso, get_alert_config, print_exec_summary
from kingwork_client.tables import KingWorkTables

TZ_CST = timezone(timedelta(hours=8))

PRIORITY_ORDER = {"高": 0, "中": 1, "低": 2, None: 3}
PRIORITY_EMOJI = {"高": "🔴", "中": "🟡", "低": "🟢"}


def parse_args():
    parser = argparse.ArgumentParser(description="AI 智能提醒")
    parser.add_argument("--todos-only", action="store_true", help="仅显示待办事项")
    parser.add_argument("--customers-only", action="store_true", help="仅显示客户跟进提醒")
    parser.add_argument("--calendar-only", action="store_true", help="仅显示当日日程")
    parser.add_argument("--complete", metavar="TODO_ID", help="标记待办为已完成")
    parser.add_argument("--followup", metavar="CUSTOMER", help="标记客户已跟进")
    parser.add_argument("--inactive-days", type=int, default=None,
                        help="未跟进天数阈值（默认取配置文件）")
    parser.add_argument("--verbose", "-v", action="store_true")
    return parser.parse_args()


def format_due_status(due_str) -> str:
    """格式化到期状态说明。"""
    if not due_str:
        return ""
    try:
        now = datetime.now(tz=TZ_CST)
        if isinstance(due_str, (int, float)):
            due_dt = datetime.fromtimestamp(due_str / 1000, tz=TZ_CST)
        else:
            from dateutil import parser as dp
            due_dt = dp.parse(str(due_str))
            if due_dt.tzinfo is None:
                due_dt = due_dt.replace(tzinfo=TZ_CST)

        delta = (due_dt.date() - now.date()).days
        if delta < 0:
            return f"⚠️  已逾期 {-delta} 天"
        elif delta == 0:
            return "⏰ 今天到期"
        elif delta == 1:
            return "📌 明天到期"
        elif delta <= 3:
            return f"📅 {delta} 天后到期"
        else:
            return f"{due_dt.strftime('%Y/%m/%d')}"
    except Exception:
        return str(due_str)[:20]


def format_due_date(due_str) -> str:
    """格式化到期日期。"""
    if not due_str:
        return "未设置"
    try:
        if isinstance(due_str, (int, float)):
            dt = datetime.fromtimestamp(due_str / 1000, tz=TZ_CST)
            return dt.strftime("%Y-%m-%d %H:%M")
        from dateutil import parser as dp
        dt = dp.parse(str(due_str))
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(due_str)[:20]


def p(msg: str = "", **kwargs):
    """带 flush 的 print，兼容日志系统"""
    print(msg, flush=True, **kwargs)


def show_todos(tables: KingWorkTables, verbose: bool = False):
    """显示待办事项提醒。"""
    todos = tables.get_pending_todos()
    if not todos:
        p("### 待办事项\n✅ 暂无待办事项\n")
        return []

    # 排序：先按逾期状态，再按优先级，再按到期时间
    def sort_key(item):
        fields = item["fields"]
        priority = fields.get("优先级")
        due = fields.get("到期时间")
        priority_order = PRIORITY_ORDER.get(priority, 3)

        if due:
            try:
                if isinstance(due, (int, float)):
                    dt = datetime.fromtimestamp(due / 1000, tz=TZ_CST)
                else:
                    from dateutil import parser as dp
                    dt = dp.parse(str(due))
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=TZ_CST)
                return (0, priority_order, dt)
            except Exception:
                pass
        return (1, priority_order, datetime.max.replace(tzinfo=TZ_CST))

    todos.sort(key=sort_key)

    p(f"### 📋 待办事项（{len(todos)}项）\n")

    for i, todo in enumerate(todos, 1):
        fields = todo["fields"]
        task_name = fields.get("任务名称") or "未命名任务"
        priority = fields.get("优先级") or "中"
        due = fields.get("到期时间")
        customer = fields.get("关联客户") or ""
        project = fields.get("关联项目") or ""
        status = fields.get("状态") or "待处理"

        emoji = PRIORITY_EMOJI.get(priority, "⚪")
        due_status = format_due_status(due)
        due_date = format_due_date(due)

        p(f"{i}. {emoji} [{priority}优先级] **{task_name}**")
        if due_date != "未设置":
            p(f"   - 到期时间：{due_date}  {due_status}")
        if customer:
            p(f"   - 关联客户：{customer}")
        if project:
            p(f"   - 关联项目：{project}")
        if verbose:
            p(f"   - 状态：{status}  ID：{todo['id']}")
        p()

    return todos


def show_customer_alerts(tables: KingWorkTables, inactive_days: int, verbose: bool = False):
    """显示客户跟进提醒。"""
    inactive_customers = tables.get_inactive_customers(days=inactive_days)
    if not inactive_customers:
        p(f"### 客户跟进提醒\n✅ 所有客户均在 {inactive_days} 天内有跟进记录\n")
        return []

    # 按未跟进天数排序（从多到少）
    inactive_customers.sort(key=lambda x: x.get("days_inactive", 0), reverse=True)

    p(f"### 👥 客户跟进提醒（{len(inactive_customers)} 家，超过 {inactive_days} 天未跟进）\n")

    for i, item in enumerate(inactive_customers, 1):
        fields = item["fields"]
        name = fields.get("客户名称") or "未知客户"
        days = item.get("days_inactive", 0)
        last_time = fields.get("最近跟进时间")
        customer_type = fields.get("客户类型") or ""
        status = fields.get("客户状态") or ""

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

        if days >= 30:
            urgency = "🔴"
        elif days >= 20:
            urgency = "🟡"
        else:
            urgency = "🟢"

        p(f"{i}. {urgency} **{name}** - 未跟进 {days} 天")
        p(f"   - 最近跟进：{last_str}")
        if customer_type:
            p(f"   - 客户类型：{customer_type}")
        if days >= 30:
            p(f"   - 建议：立即联系，关系可能趋冷")
        elif days >= 20:
            p(f"   - 建议：电话回访，了解最新进展")
        else:
            p(f"   - 建议：微信问候或发送有价值内容")
        if verbose:
            p(f"   - ID：{item['id']}")
        p()

    return inactive_customers


def show_calendar_events(verbose: bool = False):
    """显示当日日程（从 WPS 日历拉取）。"""
    import re
    today = datetime.now(tz=TZ_CST)
    start_iso = today.replace(hour=0, minute=0, second=0).strftime("%Y-%m-%dT%H:%M:%S+08:00")
    end_iso = today.replace(hour=23, minute=59, second=59).strftime("%Y-%m-%dT%H:%M:%S+08:00")

    wps_skill = "/root/.openclaw/skills/wps365-skill"

    # 1. 先拿日历列表，找到主日历
    list_cal = subprocess.run(
        [sys.executable, f"{wps_skill}/skills/calendar/run.py", "list-calendars"],
        capture_output=True, text=True, timeout=15,
    )
    primary_cal_id = None
    for line in list_cal.stdout.splitlines():
        if "primary" in line.lower() or "高波" in line:
            m = re.search(r"`([^`]+)`", line)
            if m:
                primary_cal_id = m.group(1)
                break
    if not primary_cal_id:
        m = re.search(r"`([^`]+)`", list_cal.stdout)
        if m:
            primary_cal_id = m.group(1)

    if not primary_cal_id:
        p("### 📅 今日日程\n⚠️  未找到可用日历\n")
        return []

    # 2. 查询日程
    list_ev = subprocess.run(
        [sys.executable, f"{wps_skill}/skills/calendar/run.py",
         "list-events", primary_cal_id,
         "--start", start_iso, "--end", end_iso],
        capture_output=True, text=True, timeout=15,
    )

    # 解析 JSON
    events = []
    try:
        m = re.search(r"```json\s*(.*?)\s*```", list_ev.stdout, re.DOTALL)
        if m:
            data = json.loads(m.group(1))
            events = data.get("items", []) or []
    except Exception:
        pass

    if not events:
        p(f"### 📅 今日日程\n✅ 今日暂无日程\n")
        return []

    # 按开始时间排序
    def get_start(ev):
        t = ev.get("original_start_time", {})
        if isinstance(t, dict):
            s = t.get("datetime", "")
        else:
            s = str(t)
        try:
            from dateutil import parser as dp
            dt = dp.parse(s)
            return dt if dt.tzinfo else dt.replace(tzinfo=TZ_CST)
        except Exception:
            return datetime.max.replace(tzinfo=TZ_CST)

    events.sort(key=get_start)

    p(f"### 📅 今日日程（{len(events)} 条）\n")
    for i, ev in enumerate(events, 1):
        summary = ev.get("summary", "无标题")
        org = ev.get("organizer", {})
        organizer_name = org.get("name", "") or org.get("user_id", "")
        start_t = ev.get("original_start_time", {})
        if isinstance(start_t, dict):
            start_str = start_t.get("datetime", "")[:16] if start_t.get("datetime") else ""
        else:
            start_str = str(start_t)[:16]
        online = "📹 在线会议" if ev.get("online_meeting") else ""
        reminders = "🔔 提醒已设置" if ev.get("reminders") else ""

        p(f"{i}. **{summary}**")
        if start_str:
            try:
                from dateutil import parser as dp
                dt = dp.parse(start_str)
                p(f"   - 时间：{dt.strftime('%H:%M')}")
            except Exception:
                if start_str:
                    p(f"   - 时间：{start_str[11:16]}")
        if organizer_name and organizer_name != "1711052359":
            p(f"   - 组织者：{organizer_name}")
        if online:
            p(f"   - {online}")
        if reminders:
            p(f"   - {reminders}")
        if verbose:
            p(f"   - 日程ID：{ev.get('id')}  日历ID：{ev.get('calendar_id')}")
        p()

    return events


def complete_todo(tables: KingWorkTables, todo_id: str):
    """标记待办为已完成（委托 kingupdate 完成）。"""
    import subprocess as _subprocess
    import sys as _sys
    try:
        completed = _subprocess.run(
            [sys.executable, str(Path(__file__).parent.parent / "kingupdate" / "run.py"),
             "complete", todo_id],
            capture_output=True, text=True,
            cwd=str(Path(__file__).parent.parent.parent),
        )
        if completed.returncode == 0:
            p(f"✅ 待办已标记为完成（ID：{todo_id}）")
        else:
            p(f"❌ 标记失败：{completed.stderr.strip()}", file=_sys.stderr)
            _sys.exit(1)
    except Exception as e:
        p(f"❌ 标记失败：{e}", file=_sys.stderr)
        _sys.exit(1)


def mark_customer_followup(tables: KingWorkTables, customer_name: str):
    """标记客户已跟进。"""
    try:
        updated = tables.update_customer_last_followup(customer_name)
        if updated:
            p(f"✅ 已更新「{customer_name}」的最近跟进时间为今天")
        else:
            p(f"⚠️  未找到客户「{customer_name}」的档案")
            p("   提示：可以先用 kingrecord 记录一条客户跟进，系统会自动更新客户档案")
    except Exception as e:
        p(f"❌ 更新失败：{e}", file=sys.stderr)
        sys.exit(1)


def main():
    from kingwork_client.base import debug_log
    debug_log("当前执行技能：kingalert（智能提醒）")
    args = parse_args()

    try:
        tables = KingWorkTables()
    except Exception as e:
        p(f"❌ 初始化失败：{e}", file=sys.stderr)
        p("   请确认已设置 KINGWORK_FILE_ID 环境变量并运行过 init_tables.py")
        sys.exit(1)

    # 处理操作命令
    if args.complete:
        complete_todo(tables, args.complete)
        return

    if args.followup:
        mark_customer_followup(tables, args.followup)
        return

    # 获取提醒配置
    alert_cfg = get_alert_config()
    inactive_days = args.inactive_days or alert_cfg.get("inactive_customer_days", 15)

    # 显示提醒
    p(f"\n## 工作提醒 - {today_str()}\n")

    if args.calendar_only:
        show_calendar_events(verbose=args.verbose)
        return

    if not args.customers_only:
        show_todos(tables, verbose=args.verbose)

    if not args.todos_only:
        if not args.customers_only:
            show_calendar_events(verbose=args.verbose)
        show_customer_alerts(tables, inactive_days=inactive_days, verbose=args.verbose)

    # 操作提示
    p("---")
    p("💡 操作提示：")
    p("  标记待办完成：python skills/kingalert/run.py --complete <todo_id>")
    p("  标记客户跟进：python skills/kingalert/run.py --followup <客户名称>")

    # 输出统一总结
    updated_tables = []
    if args.complete:
        updated_tables = ["todo_records"]
    elif args.followup:
        updated_tables = ["customer_profiles"]
    print_exec_summary(updated_tables)


if __name__ == "__main__":
    main()
