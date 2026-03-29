#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kingrecord - 工作日记记录
支持自然语言输入，AI 分类后写入多维表并分发到业务表。

用法:
  python skills/kingrecord/run.py "工作内容"
  python skills/kingrecord/run.py k1 "客户跟进内容"
  python skills/kingrecord/run.py --type "客户跟进" --customer "某公司" "内容"
"""
import argparse
import json
import sys
import os
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# 将 kingwork 根目录加入 path
_root = Path(__file__).resolve().parent.parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from kingwork_client.base import now_iso, today_str, weekday_cn, print_exec_summary, get_llm_config, get_wps365_root, load_config
from kingwork_client.llm import LLMClient
from kingwork_client.tables import KingWorkTables

# ── 从配置文件动态读取映射表 ──
_wt_cfg = load_config().get("work_types", {})
SHORTCUT_MAP = _wt_cfg.get("shortcuts", {})
KEYWORD_MAP = _wt_cfg.get("keywords", {})
_DEFAULT_TYPE = _wt_cfg.get("default_type", "其他")


def parse_args():
    parser = argparse.ArgumentParser(description="工作日记记录")
    parser.add_argument("content", nargs="?", help="工作内容（自然语言）")
    parser.add_argument("extra_content", nargs="?", help="当第一个参数为快捷指令时的内容")
    parser.add_argument("--type", "-t", dest="work_type", help="工作类型（跳过AI分类）")
    parser.add_argument("--customer", "-c", help="关联客户名称")
    parser.add_argument("--project", "-p", help="关联项目名称")
    parser.add_argument("--keyword", "-k", help="通过关键词指定类型")
    parser.add_argument("--verbose", "-v", action="store_true", help="详细输出")
    parser.add_argument("--dry-run", action="store_true", help="仅分析不写入")
    parser.add_argument("--list-recent", type=int, nargs="?", const=10, help="列出最近N条记录（默认10条）")
    parser.add_argument("--update", help="更新指定ID的记录，后面跟新的内容")
    return parser.parse_args()


def get_content_and_type(args):
    """解析内容和工作类型。"""
    content = args.content or ""
    work_type = args.work_type
    customer = args.customer
    project = args.project

    # 处理快捷指令
    if content in SHORTCUT_MAP:
        mapped = SHORTCUT_MAP[content]
        if mapped == "__auto__":
            # 通用记录触发词（kr/记录/记一下/写日记），不预设类型，走 AI 分类
            content = args.extra_content or ""
            if not content:
                print(f"❌ 请输入内容，例如：python run.py {args.content} \"内容\"")
                sys.exit(1)
        else:
            work_type = mapped
            content = args.extra_content or ""
            if not content:
                print(f"❌ 请输入内容，例如：python run.py {args.content} \"内容\"")
                sys.exit(1)

    # 处理关键词
    if args.keyword and not work_type:
        for kw, wt in KEYWORD_MAP.items():
            if kw in args.keyword:
                work_type = wt
                break

    return content, work_type, customer, project


# ── 枚举值校验（模块级） ─────────────────────────────────────────
VALID_METHODS = ["电话", "微信", "邮件", "上门", "会议", "其他"]
VALID_RESULTS = ["有兴趣", "需求明确", "待决策", "无需求", "已成交"]
VALID_TAGS = ["重要", "紧急", "待跟进", "已完成", "惊喜"]


def validate_enum(value: str, allowed: list, default: str) -> str:
    """将值校验为枚举值，不在枚举里则返回默认值。"""
    if value in allowed:
        return value
    return default


def validate_tags(tags_input, allowed: list) -> list:
    """校验标签列表，只保留在允许范围内的项。"""
    if not tags_input:
        return []
    if isinstance(tags_input, str):
        items = [t.strip() for t in tags_input.replace("，", ",").split(",") if t.strip()]
    else:
        items = tags_input if isinstance(tags_input, list) else []
    return [t for t in items if t in allowed]


def dispatch_to_business_table(tables: KingWorkTables, work_type: str, content: str,
                                extracted: dict, diary_id: str, verbose: bool = False):
    """根据工作类型分发到对应业务表。"""
    results = []

    if work_type == "客户跟进":
        customer = extracted.get("customer") or ""
        contact = extracted.get("contact") or ""
        method = validate_enum(extracted.get("communication_method") or "", VALID_METHODS, "其他")
        result = validate_enum(extracted.get("follow_result") or "", VALID_RESULTS, "待决策")
        next_time = extracted.get("next_followup_time")
        project = extracted.get("project") or ""
        todo = extracted.get("todo") or ""

        if customer:
            # 1. 创建客户跟进记录
            rec = tables.create_customer_followup(
                customer=customer,
                content=extracted.get("content_optimized") or content,
                method=method,
                result=result,
                next_time=next_time,
                diary_id=diary_id,
                source="手动输入",
            )
            if rec:
                results.append(f"✅ 客户跟进记录表：已创建（客户：{customer}）")

            # 2. 客户档案：不存在则创建，存在则更新跟进时间
            try:
                # 直接调用dbsheet命令创建客户档案，避免封装类依赖问题
                import subprocess
                import json
                wps_skill_path = str(get_wps365_root())
                file_id = tables.file_id
                customer_sheet_id = tables.sheet_ids['customer_profiles']  # 03客户档案 sheet_id 从配置读取
                # 先查询客户是否存在
                cmd = [
                    sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                    "list-records", file_id, str(customer_sheet_id),
                    "--page-size", "100"
                ]
                resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
                customer_exists = customer in resp
                if not customer_exists:
                    # 新建客户档案
                    customer_data = [{
                        "客户名称": customer,
                        "联系人": contact,
                        "客户类型": "潜在客户",
                        "客户状态": "跟进中",
                        "最近跟进时间": today_str(),
                        "跟进次数": 1
                    }]
                    create_cmd = [
                        sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                        "create-records", file_id, str(customer_sheet_id),
                        "--json", json.dumps(customer_data, ensure_ascii=False)
                    ]
                    subprocess.check_output(create_cmd, timeout=10)
                    results.append(f"✅ 客户档案：已新建（{customer}，联系人：{contact}）")
                else:
                    # 客户已存在，更新最近跟进时间 + 跟进次数+1
                    tables.update_customer_last_followup(customer)
                    results.append(f"✅ 客户档案：已更新最近跟进时间，跟进次数+1（{customer}）")
            except Exception as e:
                pass

        # 3. 项目档案：提取到项目则创建
        if project:
            try:
                import subprocess
                import json
                wps_skill_path = str(get_wps365_root())
                file_id = tables.file_id
                project_sheet_id = tables.sheet_ids['project_profiles']  # 04项目档案 sheet_id 从配置读取
                # 查询项目是否存在
                cmd = [
                    sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                    "list-records", file_id, str(project_sheet_id),
                    "--page-size", "100"
                ]
                resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
                project_exists = project in resp
                if not project_exists:
                    # 新建项目档案
                    project_data = [{
                        "项目名称": project,
                        "项目状态": "进行中",
                        "关联客户": customer,
                        "开始时间": today_str()
                    }]
                    create_cmd = [
                        sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                        "create-records", file_id, str(project_sheet_id),
                        "--json", json.dumps(project_data, ensure_ascii=False)
                    ]
                    subprocess.check_output(create_cmd, timeout=10)
                    results.append(f"✅ 项目档案：已新建（{project}，关联客户：{customer}）")
                else:
                    results.append(f"✅ 项目档案：已存在（{project}）")
            except Exception as e:
                pass

        # 4. 待办记录：提取到待办则创建
        if todo:
            due_date = extracted.get("due_date")
            priority = extracted.get("priority", "中")
            try:
                todo_rec = tables.create_todo_record(
                    task_name=todo,
                    customer=customer,
                    project=project,
                    priority=priority,
                    due_date=due_date,
                    status="待处理",
                    source="跟进自动生成",
                )
                if todo_rec:
                    results.append(f"✅ 待办记录表：已创建「{todo}」（优先级：{priority}，到期：{due_date}）")
            except Exception as e:
                pass

    elif work_type == "待办事项":
        task_name = extracted.get("task_name") or content[:50]
        priority = extracted.get("priority") or "中"
        due_date = extracted.get("due_date")
        customer = extracted.get("customer") or ""
        project = extracted.get("project") or ""

        rec = tables.create_todo_record(
            task_name=task_name,
            priority=priority,
            due_date=due_date,
            customer=customer if customer else None,
            project=project if project else None,
            description=extracted.get("content_optimized") or content,
        )
        if rec:
            results.append(f"✅ 待办记录表：已创建「{task_name}」（优先级：{priority}）")

    elif work_type == "学习成长":
        topic = extracted.get("learning_topic") or content[:50]
        learning_type = extracted.get("learning_type")
        key_takeaway = extracted.get("key_takeaway")
        duration = extracted.get("duration_hours")

        rec = tables.create_learning_record(
            topic=topic,
            content=extracted.get("content_optimized") or content,
            learning_type=learning_type,
            key_takeaway=key_takeaway,
            duration_hours=duration,
            diary_id=diary_id,
        )
        if rec:
            results.append(f"✅ 学习成长记录表：已创建「{topic}」")

    elif work_type == "横向支持":
        target = extracted.get("support_target") or "同事"
        support_type = extracted.get("support_type")
        customer = extracted.get("customer") or ""
        project = extracted.get("project") or ""

        rec = tables.create_support_record(
            target=target,
            content=extracted.get("content_optimized") or content,
            support_type=support_type,
            diary_id=diary_id,
        )
        if rec:
            results.append(f"✅ 横向支持记录表：已创建（支持对象：{target}）")

        # 联动更新：客户档案（如果有客户）
        if customer:
            try:
                tables.update_customer_last_followup(customer)
                results.append(f"✅ 客户档案：已更新最近跟进时间，跟进次数+1（{customer}）")
            except Exception:
                pass

        # 联动更新：项目档案（如果有项目）
        if project:
            try:
                import subprocess, json
                wps_skill_path = str(get_wps365_root())
                file_id = tables.file_id
                project_sheet_id = tables.sheet_ids["project_profiles"]
                cmd = [
                    sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                    "list-records", file_id, str(project_sheet_id), "--page-size", "100"
                ]
                resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
                project_exists = project in resp
                if not project_exists:
                    project_data = [{
                        "项目名称": project,
                        "项目状态": "进行中",
                        "关联客户": customer,
                        "开始时间": today_str(),
                    }]
                    create_cmd = [
                        sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                        "create-records", file_id, str(project_sheet_id),
                        "--json", json.dumps(project_data, ensure_ascii=False)
                    ]
                    subprocess.check_output(create_cmd, timeout=10)
                    results.append(f"✅ 项目档案：已新建（{project}，关联客户：{customer}）")
            except Exception:
                pass

    elif work_type == "团队事务":
        topic = extracted.get("task_name") or content[:50]
        event_type = extracted.get("event_type")
        participants = extracted.get("participants")
        customer = extracted.get("customer") or ""
        project = extracted.get("project") or ""

        rec = tables.create_team_record(
            topic=topic,
            content=extracted.get("content_optimized") or content,
            event_type=event_type,
            participants=participants,
            diary_id=diary_id,
        )
        if rec:
            results.append(f"✅ 团队事务记录表：已创建「{topic}」")

        # 联动更新：客户档案
        if customer:
            try:
                tables.update_customer_last_followup(customer)
                results.append(f"✅ 客户档案：已更新最近跟进时间，跟进次数+1（{customer}）")
            except Exception:
                pass

        # 联动更新：项目档案
        if project:
            try:
                import subprocess, json
                wps_skill_path = str(get_wps365_root())
                file_id = tables.file_id
                project_sheet_id = tables.sheet_ids["project_profiles"]
                cmd = [
                    sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                    "list-records", file_id, str(project_sheet_id), "--page-size", "100"
                ]
                resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
                project_exists = project in resp
                if not project_exists:
                    project_data = [{
                        "项目名称": project,
                        "项目状态": "进行中",
                        "关联客户": customer,
                        "开始时间": today_str(),
                    }]
                    create_cmd = [
                        sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                        "create-records", file_id, str(project_sheet_id),
                        "--json", json.dumps(project_data, ensure_ascii=False)
                    ]
                    subprocess.check_output(create_cmd, timeout=10)
                    results.append(f"✅ 项目档案：已新建（{project}，关联客户：{customer}）")
            except Exception:
                pass

    elif work_type == "灵感记录":
        category = extracted.get("idea_category")
        feasibility = extracted.get("feasibility")

        rec = tables.create_idea_record(
            content=extracted.get("content_optimized") or content,
            category=category,
            feasibility=feasibility,
            diary_id=diary_id,
        )
        if rec:
            results.append(f"✅ 灵感记录表：已创建")

    elif work_type == "活动接待":
        subject = extracted.get("task_name") or extracted.get("event_subject") or content[:50]
        event_type = extracted.get("event_type")
        reception_target = extracted.get("reception_target") or ""
        customer = extracted.get("customer") or ""
        project = extracted.get("project") or ""

        rec = tables.create_event_reception(
            subject=subject,
            content=extracted.get("content_optimized") or content,
            event_type=event_type,
            reception_target=reception_target,
            customer=customer,
            project=project,
            diary_id=diary_id,
        )
        if rec:
            results.append(f"✅ 活动接待记录表：已创建「{subject}」（接待对象：{reception_target}）")

        # 联动更新：客户档案
        if customer:
            try:
                tables.update_customer_last_followup(customer)
                results.append(f"✅ 客户档案：已更新最近跟进时间，跟进次数+1（{customer}）")
            except Exception:
                pass

    return results


def main():
    # 调试日志：输出当前执行的技能
    from kingwork_client.base import debug_log
    debug_log("当前执行技能：kingrecord（工作记录）")
    args = parse_args()
    wps_skill_path = str(get_wps365_root())
    file_id = os.environ.get("KINGWORK_FILE_ID", "")
    
    # 处理列出最近记录
    if args.list_recent:
        if not file_id:
            print("❌ 请先设置KINGWORK_FILE_ID环境变量")
            sys.exit(1)
        print(f"## 最近 {args.list_recent} 条记录：")
        cmd = [
            sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
            "list-records", file_id, "2", "--page-size", str(args.list_recent)
        ]
        try:
            import subprocess
            resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
            import re
            records = re.findall(r'"id": "(.*?)".*?"内容":"(.*?)"', resp)
            for i, (rid, content) in enumerate(reversed(records[-args.list_recent:]), 1):
                print(f"{i}. 🆔 {rid} | 内容：{content[:50]}..." if len(content) > 50 else f"{i}. 🆔 {rid} | 内容：{content}")
        except Exception as e:
            print(f"❌ 查询失败：{e}")
        sys.exit(0)
    
    # 处理更新记录
    if args.update:
        record_id = args.update
        new_content = args.content or ""
        if not new_content:
            print("❌ 请输入新的内容")
            sys.exit(1)
        if not file_id:
            print("❌ 请先设置KINGWORK_FILE_ID环境变量")
            sys.exit(1)
        # 先删除旧的关联记录（跟进/待办等，简化逻辑）
        # 然后重新创建新的记录
        print(f"## 正在更新记录 {record_id}...")
        # 先删除旧记录
        try:
            cmd = [
                sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                "delete-records", file_id, "2", record_id
            ]
            subprocess.check_output(cmd, timeout=10)
            print(f"✅ 旧记录已删除")
        except Exception as e:
            print(f"⚠️ 删除旧记录失败：{e}，将直接创建新记录")
        # 用新内容重新走创建流程
        args.content = new_content
    
    content, work_type, customer, project = get_content_and_type(args)

    if not content:
        print("❌ 请输入工作内容")
        sys.exit(1)

    llm = LLMClient()

    # ── 客户名称LLM语义匹配 ──────────────────────────────────────────
    def validate_customer_name(extracted_customer: str, existing_customers: list, content: str) -> str:
        """验证客户名称是否与已有客户匹配（LLM语义匹配为主，兜底校验为辅）。
        
        逻辑：
        1. 没有提取到客户名 → 返回空
        2. 已有客户列表为空 → 返回原始提取值
        3. 先让 LLM 做语义匹配（已有客户 vs 提取名称）：
           - 匹配成功 → 返回匹配到的已有客户名
           - 匹配失败 → 降级做兜底校验
        4. 兜底校验：提取的客户名必须是内容的子串（防止幻觉扩展）
           - 通过 → 返回原始提取值
           - 不通过 → 丢弃
        """
        if not extracted_customer:
            return ""
        
        if not existing_customers:
            return extracted_customer
        
        # 第一步：LLM 语义匹配（优先）
        try:
            result = llm.validate_customer(content, extracted_customer, existing_customers)
            if result:
                customer = result.get("customer")
                is_new = result.get("is_new", False)
                reason = result.get("reason", "")
                if customer:
                    if not is_new:
                        print(f"  🔗 LLM匹配客户：'{extracted_customer}' → '{customer}'（{reason}）")
                    else:
                        print(f"  🆕 LLM判断为新客户：'{customer}'（{reason}）")
                    return customer
                # LLM 判断为幻觉
                print(f"  ⚠️ LLM判断提取的'{extracted_customer}'是幻觉，已丢弃（{reason}）")
                return ""
        except Exception as e:
            print(f"  ⚠️ LLM客户匹配失败：{e}，降级到兜底校验")
        
        # 第二步：兜底校验（防止 LLM 幻觉扩展客户名）
        if extracted_customer not in content:
            print(f"  ⚠️ 兜底校验：'{extracted_customer}'不是内容子串，已丢弃")
            return ""
        
        return extracted_customer

    # 先读取现有客户和项目列表，用于匹配去重
    existing_customers = []
    existing_projects = []
    import subprocess
    import json
    wps_skill_path = str(get_wps365_root())
    file_id = os.environ.get("KINGWORK_FILE_ID", "")
    if file_id:
        try:
            # 读取客户列表
            cmd = [
                sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                "list-records", file_id, "4", "--page-size", "200"
            ]
            resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
            import re
            customer_matches = re.findall(r'"客户名称":"(.*?)"', resp)
            existing_customers = list(set(customer_matches))
            
            # 读取项目列表
            cmd = [
                sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
                "list-records", file_id, "5", "--page-size", "200"
            ]
            resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
            project_matches = re.findall(r'"项目名称":"(.*?)"', resp)
            existing_projects = list(set(project_matches))
        except Exception as e:
            pass

    # ============================================================
    # 两步式 LLM 调用：Step1 分类 → Step2 信息提取
    # ============================================================
    import time
    confidence = 0.7
    extracted = {}

    # --- Step1: 轻量分类（仅在未手动指定类型时执行）---
    if not work_type:
        classify_result = None
        for attempt in range(3):
            try:
                classify_result = llm.classify_work_type(content)
                if classify_result and classify_result.get("work_type"):
                    break
                if attempt < 2:
                    print(f"  ⚠️ 分类LLM返回异常（第{attempt+1}次），3秒后重试...")
                    time.sleep(3)
            except Exception as e:
                if attempt < 2:
                    print(f"  ⚠️ 分类LLM调用异常：{e}（第{attempt+1}次），3秒后重试...")
                    time.sleep(3)
                else:
                    print(f"  ❌ 分类LLM最终失败：{e}，降级到关键词分类")

        if classify_result and classify_result.get("work_type"):
            work_type = classify_result["work_type"]
            confidence = classify_result.get("confidence", 0.8)
            if args.verbose:
                reason = classify_result.get("reason", "")
                print(f"  📋 Step1 分类结果：{work_type}（置信度{confidence:.0%}，{reason}）")
        else:
            # 降级到关键词分类
            work_type = _keyword_classify(content)
            confidence = 0.5
            if args.verbose:
                print(f"  📋 Step1 降级到关键词分类：{work_type}")

    # --- Step2: 按类型定向信息提取 ---
    extract_result = None
    for attempt in range(3):
        try:
            extract_result = llm.extract_work_info(content, work_type, existing_customers, existing_projects)
            if extract_result is not None:
                break
            if attempt < 2:
                print(f"  ⚠️ 提取LLM返回None（第{attempt+1}次），3秒后重试...")
                time.sleep(3)
        except Exception as e:
            if attempt < 2:
                print(f"  ⚠️ 提取LLM调用异常：{e}（第{attempt+1}次），3秒后重试...")
                time.sleep(3)
            else:
                print(f"  ❌ 提取LLM最终失败：{e}，降级到规则提取")

    try:
        if extract_result and isinstance(extract_result, dict):
            extracted = extract_result
            # LLM客户名二次验证
            llm_customer = extracted.get("customer") or ""
            validated_customer = validate_customer_name(llm_customer, existing_customers, content)
            extracted["customer"] = validated_customer
        else:
            raise ValueError("提取LLM返回格式异常")
    except Exception as e:
        # LLM调用失败时，降级到规则提取
        extracted = {}
        confidence = min(confidence, 0.5)
        
        # 规则提取关键信息
        import re
        import datetime
        today = datetime.date.today()
        
        # 提取客户名称 + 联系人（增强匹配：支持开头直接是拜访/见客户等场景）
        customer_match = re.search(r"(拜访|见|沟通|和|与|跟)(.*?)(王总|李总|张总|刘总|陈总|黄总|吴总|周总|赵总|朱总|林总|总|客户)", content)
        if customer_match:
            customer_name = customer_match.group(2).strip()
            # 清理客户名称里的多余字符
            customer_name = re.sub(r"(的|沟通|交流|见面|拜访|去|再|做一次)", "", customer_name).strip()
            # 检查是否有匹配的已有客户
            matched_customer = None
            for exist_c in existing_customers:
                if customer_name in exist_c or exist_c in customer_name:
                    matched_customer = exist_c
                    break
            extracted["customer"] = matched_customer if matched_customer else customer_name
            # 提取联系人
            contact_match = re.search(r"(王总|李总|张总|刘总|陈总|黄总|吴总|周总|赵总|朱总|林总)", content)
            if contact_match:
                extracted["contact"] = contact_match.group(1)
        else:
            # 没有明确联系人的客户提取
            customer_match = re.search(r"(拜访|见|沟通|和|与|跟)(.*?)(公司|集团|股份|银行|证券|保险|科技|建投|建筑|银行)", content)
            if customer_match:
                customer_name = customer_match.group(2).strip() + customer_match.group(3).strip()
                customer_name = re.sub(r"(的|沟通|交流|见面|拜访|去|再|做一次)", "", customer_name).strip()
                # 匹配已有客户
                matched_customer = None
                for exist_c in existing_customers:
                    if customer_name in exist_c or exist_c in customer_name:
                        matched_customer = exist_c
                        break
                extracted["customer"] = matched_customer if matched_customer else customer_name
        
        # 提取项目名称（组合关键词+匹配已有项目）
        project_keywords = ["AIDoc", "混合云", "文档中心", "文档中台", "WPS365", "智能办公", "协同办公", "信创", "私有化", "云办公", "部署", "升级"]
        matched_kws = [kw for kw in project_keywords if kw in content]
        project_name = ""
        if matched_kws:
            project_name = " ".join(matched_kws) + "项目"
        if not project_name:
            project_match = re.search(r"(.*?)(项目|需求|方案|系统|平台)", content)
            if project_match:
                project_name = project_match.group(1).strip() + project_match.group(2).strip()
                if len(project_name) < 3 or len(project_name) > 20:
                    project_name = ""
        # 匹配已有项目
        if project_name:
            matched_project = None
            for exist_p in existing_projects:
                if project_name in exist_p or exist_p in project_name:
                    matched_project = exist_p
                    break
            extracted["project"] = matched_project if matched_project else project_name
        
        # 提取待办任务
        if "提交" in content or "完成" in content or "准备" in content or "出具" in content or "发送" in content or "下次去" in content or "再沟通" in content:
            todo_match = re.search(r"(提交|完成|准备|出具|发送|下次沟通|再去)(.*?)(，|。|$)", content)
            if todo_match:
                todo_content = todo_match.group(1).strip() + todo_match.group(2).strip()
                # 清理多余标点
                todo_content = re.sub(r"[，。？！]$", "", todo_content).strip()
                extracted["todo"] = todo_content
                extracted["task_name"] = todo_content
                extracted["priority"] = "高"
        
        # 提取日期
        date_map = {
            "明天": today + datetime.timedelta(days=1),
            "后天": today + datetime.timedelta(days=2),
            "下周一": today + datetime.timedelta(days=(7 - today.weekday() + 0)),
            "下周二": today + datetime.timedelta(days=(7 - today.weekday() + 1)),
            "下周三": today + datetime.timedelta(days=(7 - today.weekday() + 2)),
            "下周四": today + datetime.timedelta(days=(7 - today.weekday() + 3)),
            "下周五": today + datetime.timedelta(days=(7 - today.weekday() + 4)),
            "下周六": today + datetime.timedelta(days=(7 - today.weekday() + 5)),
            "下周日": today + datetime.timedelta(days=(7 - today.weekday() + 6)),
        }
        for date_str, date_obj in date_map.items():
            if date_str in content:
                date_iso = date_obj.strftime("%Y/%m/%d")
                extracted["next_followup_time"] = date_iso
                extracted["due_date"] = date_iso
                break
    
    # 如果是手动指定了工作类型，优先使用手动指定的类型，覆盖大模型返回的
    if args.work_type or content in SHORTCUT_MAP:
        confidence = 1.0
    
    # 双重保险：不管大模型有没有提取到日期，都用规则再提取一次，补充到extracted
    import re
    import datetime
    today = datetime.date.today()
    # 提取日期
    date_map = {
        "明天": today + datetime.timedelta(days=1),
        "后天": today + datetime.timedelta(days=2),
        "下周一": today + datetime.timedelta(days=(7 - today.weekday() + 0)),
        "下周二": today + datetime.timedelta(days=(7 - today.weekday() + 1)),
        "下周三": today + datetime.timedelta(days=(7 - today.weekday() + 2)),
        "下周四": today + datetime.timedelta(days=(7 - today.weekday() + 3)),
        "下周五": today + datetime.timedelta(days=(7 - today.weekday() + 4)),
        "下周六": today + datetime.timedelta(days=(7 - today.weekday() + 5)),
        "下周日": today + datetime.timedelta(days=(7 - today.weekday() + 6)),
        "下周1": today + datetime.timedelta(days=(7 - today.weekday() + 0)),
        "下周2": today + datetime.timedelta(days=(7 - today.weekday() + 1)),
        "下周3": today + datetime.timedelta(days=(7 - today.weekday() + 2)),
        "下周4": today + datetime.timedelta(days=(7 - today.weekday() + 3)),
        "下周5": today + datetime.timedelta(days=(7 - today.weekday() + 4)),
        "下周6": today + datetime.timedelta(days=(7 - today.weekday() + 5)),
        "下周7": today + datetime.timedelta(days=(7 - today.weekday() + 6)),
        "周1": today + datetime.timedelta(days=(7 - today.weekday() + 0)) if today.weekday() >=0 else today,
        "周2": today + datetime.timedelta(days=(7 - today.weekday() + 1)) if today.weekday() >=1 else today,
        "周3": today + datetime.timedelta(days=(7 - today.weekday() + 2)) if today.weekday() >=2 else today,
        "周4": today + datetime.timedelta(days=(7 - today.weekday() + 3)) if today.weekday() >=3 else today,
        "周5": today + datetime.timedelta(days=(7 - today.weekday() + 4)) if today.weekday() >=4 else today,
        "周6": today + datetime.timedelta(days=(7 - today.weekday() + 5)) if today.weekday() >=5 else today,
        "周日": today + datetime.timedelta(days=(7 - today.weekday() + 6)) if today.weekday() >=6 else today,
    }
    for date_str, date_obj in date_map.items():
        if date_str in content:
            date_iso = date_obj.strftime("%Y-%m-%d")
            extracted["next_followup_time"] = date_iso
            extracted["due_date"] = date_iso
            break

    # 覆盖从命令行指定的字段
    if customer:
        extracted["customer"] = customer
    if project:
        extracted["project"] = project

    # 正则提取完成后，统一做一次客户名校验（覆盖LLM失败的降级路径）
    regex_customer = extracted.get("customer") or ""
    if regex_customer:
        validated_customer = validate_customer_name(regex_customer, existing_customers, content)
        extracted["customer"] = validated_customer

    # 兜底：LLM失败且关键词也没匹配到任何类型时，默认为"其他"
    if not work_type:
        work_type = "其他"

    # 快捷指令客户跟进场景，自动提取客户名称
    if work_type == "客户跟进" and not extracted.get("customer"):
        import re
        # 提取客户名称：匹配"针对XX银行/集团/公司"或"拜访XX集团/公司"格式
        customer_match = re.search(r"(针对|关于|拜访|见|沟通|和|与|跟)(.*?)(集团|公司|股份|银行|证券|保险|科技|建投|建筑|总)", content)
        if customer_match:
            extracted_customer = customer_match.group(2).strip() + customer_match.group(3).strip()
            # LLM语义匹配验证
            validated = validate_customer_name(extracted_customer, existing_customers, content)
            extracted["customer"] = validated
        # 提取项目名称：已知项目关键词直接匹配 + "XX项目/平台/系统"格式
        project_keywords = ["AIDoc", "混合云", "文档中心", "文档中台", "WPS365", "智能办公", "协同办公", "信创", "私有化", "云办公", "部署", "升级", "一站式平台"]
        for kw in project_keywords:
            if kw in content:
                extracted["project"] = kw
                break
        if not extracted.get("project"):
            project_match = re.search(r"(参加|沟通|讨论|跟进|做|准备)(.*?)(项目|建设|需求|研讨会|方案|对接|平台|系统)", content)
            if project_match:
                extracted_project = project_match.group(2).strip() + project_match.group(3).strip()
                extracted["project"] = extracted_project

    summary = extracted.get("summary") or content[:100]
    tags = validate_tags(extracted.get("tags"), VALID_TAGS)

    if args.dry_run:
        print(f"\n[DRY RUN] 工作类型：{work_type}")
        print(f"[DRY RUN] 内容摘要：{summary}")
        print(f"[DRY RUN] 提取信息：{json.dumps(extracted, ensure_ascii=False, indent=2)}")
        return

    # 写入日记记录表
    try:
        tables = KingWorkTables()
        diary_rec = tables.create_diary_record(
            content=extracted.get("content_optimized") or summary,
            work_type=work_type,
            customer=extracted.get("customer") or customer,
            project=extracted.get("project") or project,
            tags=tags,
            note=content,
        )

        diary_id = diary_rec.get("id") if diary_rec else ""

        # 输出主记录结果
        print(f"\n✅ 已记录为【{work_type}】")
        if extracted.get("customer"):
            print(f"   客户：{extracted['customer']}")
        if extracted.get("project"):
            print(f"   项目：{extracted['project']}")
        print(f"   时间：{today_str()}")
        if confidence < 0.7:
            print(f"   ⚠️  分类置信度较低（{confidence:.0%}），请确认类型是否正确")

        # 分发到业务表
        dispatch_results = dispatch_to_business_table(
            tables, work_type, content, extracted, diary_id, verbose=args.verbose
        )

        if dispatch_results:
            print("\n## 已同步到其他表：")
            for r in dispatch_results:
                print(f"   {r}")
        
        # 输出结构化总结
        print("\n📝 本次记录完整信息：")
        print(f"   🔖 类型：{work_type}")
        if extracted.get("customer"):
            print(f"   👥 客户：{extracted['customer']}")
            if extracted.get("contact"):
                print(f"   📞 联系人：{extracted['contact']}")
        if extracted.get("project"):
            print(f"   📋 项目：{extracted['project']}")
        if extracted.get("todo"):
            print(f"   ✅ 待办任务：{extracted['todo']}")
            if extracted.get("due_date"):
                print(f"   ⏰ 到期时间：{extracted['due_date']}")
            if extracted.get("priority"):
                print(f"   ⚠️  优先级：{extracted['priority']}")
        print(f"   📅 记录时间：{today_str()}")
        # 记录本次ID，方便后续更新
        if diary_id:
            print(f"   🆔 记录ID：{diary_id}（如需修改请使用该ID）")

        # 收集更新的数据表
        updated_tables = ["diary_records"]
        if work_type == "客户跟进":
            updated_tables.extend(["customer_followups", "customer_profiles", "project_profiles", "todo_records"])
        elif work_type == "待办事项":
            updated_tables.append("todo_records")
        elif work_type == "学习成长":
            updated_tables.append("learning_records")
        elif work_type == "横向支持":
            updated_tables.append("support_records")
        elif work_type == "团队事务":
            updated_tables.append("team_records")
        elif work_type == "灵感记录":
            updated_tables.append("idea_records")
        # 去重
        updated_tables = list(set(updated_tables))
        # 输出统一总结
        print_exec_summary(updated_tables)

    except Exception as e:
        print(f"\n❌ 写入失败：{e}", file=sys.stderr)
        sys.exit(1)


def _keyword_classify(text: str) -> str:
    """基于关键词的简单分类（无 AI 时的降级方案）。"""
    for kw, wt in KEYWORD_MAP.items():
        if kw in text:
            return wt
    return _DEFAULT_TYPE


if __name__ == "__main__":
    main()
