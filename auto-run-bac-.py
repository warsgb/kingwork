#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kingauto - AI 自动分析
自动获取聊天/会议记录，大模型分析后写入多维表。

用法:
  python skills/kingauto/run.py
  python skills/kingauto/run.py --date 2026-03-19
  python skills/kingauto/run.py --source chat
"""
import argparse
import json
import subprocess
import sys
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path

# 全局表映射缓存
_sheet_map = None

def get_sheet_id(table_name: str) -> int:
    """从本地配置获取对应表的sheet_id，不存在则自动生成"""
    global _sheet_map
    current_file_id = os.environ.get("KINGWORK_FILE_ID", "")
    if not current_file_id:
        raise ValueError("未配置KINGWORK_FILE_ID环境变量")
    
    config_path = os.path.expanduser("~/.kingwork_sheet_map.json")
    # 先尝试读本地配置
    if _sheet_map is None and os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config_data = json.load(f)
                # 校验file_id是否匹配
                if config_data.get("file_id") == current_file_id:
                    _sheet_map = config_data.get("sheet_map", {})
        except Exception:
            pass
    
    # 没有映射或者不匹配，自动拉取生成
    if _sheet_map is None or table_name not in _sheet_map:
        wps_skill_path = "/root/.openclaw/skills/wps365-skill"
        cmd = [
            "python", f"{wps_skill_path}/skills/dbsheet/run.py",
            "schema", current_file_id, "--json"
        ]
        try:
            resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
            import re
            # 提取JSON
            m = re.search(r"{.*}", resp, re.DOTALL)
            if m:
                schema = json.loads(m.group(0))
                _sheet_map = {}
                for sheet in schema.get("sheets", []):
                    _sheet_map[sheet["name"]] = sheet["id"]
                # 保存到本地
                config_data = {
                    "file_id": current_file_id,
                    "sheet_map": _sheet_map,
                    "updated_at": datetime.now().strftime("%Y-%m-%d")
                }
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(config_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise ValueError(f"获取表映射失败：{e}")
    
    return _sheet_map.get(table_name)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

_root = Path(__file__).resolve().parent.parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from kingwork_client.base import (
    now_iso, today_str, get_wps365_root, get_import_mode, get_skill_call_mode, import_wpsv7client
)
from kingwork_client.llm import LLMClient
from kingwork_client.tables import KingWorkTables

TZ_CST = timezone(timedelta(hours=8))


def parse_args():
    parser = argparse.ArgumentParser(description="AI 自动分析聊天/会议记录")
    parser.add_argument("--date", help="分析指定日期（YYYY-MM-DD），默认今天")
    parser.add_argument("--start", help="开始日期")
    parser.add_argument("--end", help="结束日期")
    parser.add_argument("--source", choices=["chat", "meeting", "doc", "all"], default="all",
                        help="数据来源（chat/meeting/doc/all）")
    parser.add_argument("--no-dedup", action="store_true", help="跳过相似度去重")
    parser.add_argument("--verbose", "-v", action="store_true", help="详细输出")
    parser.add_argument("--dry-run", action="store_true", help="仅分析不写入")
    return parser.parse_args()


def get_date_range(args):
    """解析日期范围，返回 (start_dt, end_dt, date_range_str)。"""
    if args.start and args.end:
        start = datetime.fromisoformat(args.start).replace(tzinfo=TZ_CST)
        end = datetime.fromisoformat(args.end).replace(tzinfo=TZ_CST)
        end = end.replace(hour=23, minute=59, second=59)
    elif args.date:
        start = datetime.fromisoformat(args.date).replace(tzinfo=TZ_CST)
        end = start.replace(hour=23, minute=59, second=59)
    else:
        today = datetime.now(tz=TZ_CST).replace(hour=0, minute=0, second=0, microsecond=0)
        start = today
        end = today.replace(hour=23, minute=59, second=59)

    date_range_str = f"{start.strftime('%Y-%m-%d')} 至 {end.strftime('%Y-%m-%d')}"
    return start, end, date_range_str


def run_wps365_skill(skill_name: str, *args) -> dict:
    """调用 WPS365 子技能，返回解析后的结果。

    支持两种调用模式（由 skill_call_mode 配置控制）：
    - subprocess 模式（默认）：通过子进程调用
    - direct 模式：直接导入 wps365-skill 的子技能模块
    """
    call_mode = get_skill_call_mode()

    if call_mode == "direct":
        return run_wps365_skill_direct(skill_name, *args)
    else:
        return run_wps365_skill_subprocess(skill_name, *args)


def run_wps365_skill_subprocess(skill_name: str, *args) -> dict:
    """通过子进程调用 WPS365 子技能。"""
    wps365_root = get_wps365_root()
    skill_path = wps365_root / "skills" / skill_name / "run.py"
    if not skill_path.exists():
        return {"error": f"WPS365 skill not found: {skill_name}"}

    cmd = [sys.executable, str(skill_path)] + list(args)
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=60,
            cwd=str(wps365_root),
        )
        return {"stdout": result.stdout, "stderr": result.stderr, "returncode": result.returncode}
    except subprocess.TimeoutExpired:
        return {"error": "WPS365 skill timeout"}
    except Exception as e:
        return {"error": str(e)}


def run_wps365_skill_direct(skill_name: str, *args) -> dict:
    """直接导入调用 WPS365 子技能（更高效）。"""
    wps365_root = get_wps365_root()
    skill_module_path = wps365_root / "skills" / skill_name / "run.py"

    if not skill_module_path.exists():
        return {"error": f"WPS365 skill not found: {skill_name}"}

    # 将 wps365-skill 添加到路径
    if str(wps365_root) not in sys.path:
        sys.path.insert(0, str(wps365_root))

    # 动态导入子技能模块
    skill_module_name = f"skills.{skill_name}.run"
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location(skill_module_name, str(skill_module_path))
        if spec and spec.loader:
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # 调用子技能的 main 函数（如果存在）
            if hasattr(module, "main"):
                # 模拟命令行参数
                old_argv = sys.argv
                sys.argv = ["run.py"] + list(args)
                try:
                    # 捕获 stdout 输出
                    import io
                    from contextlib import redirect_stdout
                    output = io.StringIO()
                    with redirect_stdout(output):
                        module.main()
                    return {"stdout": output.getvalue(), "returncode": 0}
                finally:
                    sys.argv = old_argv
            else:
                return {"error": f"Skill {skill_name} has no main() function"}
        return {"error": f"Failed to load skill {skill_name}"}
    except Exception as e:
        return {"error": f"Direct import failed: {str(e)}"}


def get_chat_messages(start_dt: datetime, end_dt: datetime, verbose: bool = False) -> list:
    """获取时间范围内的聊天消息（包含所有群聊+单聊）。"""
    start_iso = start_dt.strftime("%Y-%m-%dT%H:%M:%S+08:00")
    end_iso = end_dt.strftime("%Y-%m-%dT%H:%M:%S+08:00")
    wps_skill_path = os.environ.get("WPS365_SKILL_PATH", "/root/.openclaw/skills/wps365-skill")
    
    processed_messages = []
    
    if verbose:
        print(f"  拉取所有会话列表...")
    
    # 1. 先拉取最近100个会话，包含群聊和单聊
    cmd = [
        "python", f"{wps_skill_path}/skills/im/run.py",
        "recent", "--page-size", "100"
    ]
    try:
        import subprocess
        resp = subprocess.check_output(cmd, timeout=10).decode("utf-8")
        # 提取会话列表
        conversations = _extract_json_list_from_output(resp, key="items")
        if verbose:
            print(f"  共找到 {len(conversations)} 个会话")
    except Exception as e:
        print(f"  ⚠️  获取会话列表失败：{e}", file=sys.stderr)
        conversations = []
    
    # 2. 遍历所有会话逐个拉取消息（群聊+单聊都走这个逻辑，保证不遗漏）
    if verbose:
        print(f"  逐个拉取会话消息：共 {len(conversations)} 个会话")
    for chat in conversations:
        chat_id = chat.get("id")
        chat_name = chat.get("name", "未知会话")
        if not chat_id:
            continue
        # 检查该会话最后消息时间是否在范围内
        last_msg_time = chat.get("last_message", {}).get("ctime")
        if last_msg_time:
            try:
                import dateutil.parser
                dt = dateutil.parser.isoparse(last_msg_time)
                if dt.timestamp() < start_dt.timestamp() or dt.timestamp() > end_dt.timestamp():
                    continue  # 该会话今天没有消息，跳过
            except:
                pass
        # 拉取该会话的消息
        result = run_wps365_skill("im", "search-messages",
                                  "--chat-ids", str(chat_id),
                                  "--keyword", "",
                                  "--start-time", start_iso,
                                  "--end-time", end_iso,
                                  "--page-size", "100")
        if result.get("error"):
            continue
        session_messages = _extract_json_list_from_output(result.get("stdout", ""), key="items")
        if verbose:
            print(f"  【{chat_name}】找到 {len(session_messages)} 条消息")
        processed_messages.extend(session_messages)
    
    # 3. 拉取单聊消息：每个单聊会话单独搜索
    p2p_chats = [c for c in conversations if c.get("type") == "p2p"]
    if verbose:
        print(f"  拉取单聊消息：共 {len(p2p_chats)} 个单聊会话")
    
    for chat in p2p_chats:
        chat_id = chat.get("id")
        chat_name = chat.get("name", "未知用户")
        if not chat_id:
            continue
        # 检查该会话最后消息时间是否在范围内
        last_msg_time = chat.get("last_message", {}).get("ctime")
        if last_msg_time:
            try:
                import dateutil.parser
                dt = dateutil.parser.isoparse(last_msg_time)
                if dt.timestamp() < start_dt.timestamp() or dt.timestamp() > end_dt.timestamp():
                    continue  # 该会话今天没有消息，跳过
            except:
                pass
        # 拉取该会话的消息
        result = run_wps365_skill("im", "search-messages",
                                  "--chat-ids", str(chat_id),
                                  "--keyword", "",
                                  "--start-time", start_iso,
                                  "--end-time", end_iso)
        if result.get("error"):
            continue
        p2p_messages = _extract_json_list_from_output(result.get("stdout", ""), key="items")
        if verbose:
            print(f"  【{chat_name}】找到 {len(p2p_messages)} 条消息")
        processed_messages.extend(p2p_messages)
    
    # 4. 按会话分组汇总聊天内容，生成总结，避免逐条插入
    final_messages = []
    MY_USER_ID = "1711052359" # 你的用户ID
    chat_groups = {} # key: chat_id, value: 会话信息+消息列表
    
    # 第一步：按会话分组
    for msg in processed_messages:
        try:
            chat_info = msg.get("chat", {})
            chat_id = chat_info.get("id", "")
            chat_type = chat_info.get("type", "unknown") # group=群聊，p2p=单聊
            chat_name = chat_info.get("name", "未知会话")
            sender_id = msg.get("message", {}).get("sender", {}).get("id", "")
            sender_name = msg.get("message", {}).get("sender", {}).get("name", "未知用户")
            mention_users = msg.get("message", {}).get("mention_users", [])
            msg_content = ""
            
            # 过滤规则：
            # 1. 群聊：只保留我发的，或者@我的消息
            if chat_type == "group":
                is_my_msg = (sender_id == MY_USER_ID)
                is_mention_me = any(mention.get("id") == MY_USER_ID for mention in mention_users)
                if not is_my_msg and not is_mention_me:
                    continue # 群聊既不是我发的也没@我，跳过
            
            # 2. 单聊：全部保留，不需要过滤
            
            # 处理不同类型的消息内容
            msg_type = msg.get("message", {}).get("type", "text")
            content = msg.get("message", {}).get("content", {})
            
            if msg_type == "text" and "text" in content:
                msg_content = content["text"].get("content", "")
            elif msg_type == "rich_text" and "rich_text" in content:
                # 提取富文本里的所有文本内容
                elements = content["rich_text"].get("elements", [])
                text_parts = []
                for ele in elements:
                    for sub_ele in ele.get("elements", []):
                        if "text_content" in sub_ele:
                            text_parts.append(sub_ele["text_content"].get("content", ""))
                        elif "doc_content" in sub_ele:
                            text_parts.append(f"[文档] {sub_ele['doc_content'].get('text', '')}：{sub_ele['doc_content'].get('file', {}).get('link_url', '')}")
                msg_content = "\n".join(text_parts)
            elif msg_type == "image":
                msg_content = "[图片]"
            elif msg_type == "file":
                msg_content = "[文件]"
            
            # 跳过空内容和无效内容（自己发的图片/文件保留，标记内容）
            invalid_keywords = ["好的", "收到", "谢谢", "感谢", "ok", "OK", "是的", "对的", "没问题", "👍", "❤️", "😂", "🙏", "表情", "动画表情", "[语音]", "[视频]", "嗯嗯", "哦哦", "哈哈", "嗯嗯好", "好哒", "对"]
            # 自己发的图片/文件保留
            is_my_msg = (sender_id == MY_USER_ID)
            if is_my_msg and msg_content in ["[图片]", "[文件]"]:
                msg_content = f"[我发的{'图片' if '图片' in msg_content else '文件'}] 工作相关内容"
            
            if not msg_content or len(msg_content.strip()) < 3 or any(kw == msg_content.strip() for kw in invalid_keywords):
                continue
            
            # 加入分组
            if chat_id not in chat_groups:
                chat_groups[chat_id] = {
                    "chat_name": chat_name,
                    "chat_type": chat_type,
                    "messages": []
                }
            chat_groups[chat_id]["messages"].append({
                "sender": sender_name,
                "content": msg_content
            })
        except Exception as e:
            continue
    
    # 第二步：按主题拆分会话内容，同一会话多个主题拆分多条记录
    # 预设主题关键词映射
    topic_map = {
        "客户跟进": ["项目", "方案", "需求", "报价", "商务", "合作", "客户", "招标", "投标"],
        "资源申请": ["资源", "申请", "权限", "账号", "包", "版本", "下载", "license"],
        "部署调试": ["部署", "调试", "安装", "配置", "环境", "测试", "私有化", "上线"],
        "问题排查": ["问题", "bug", "故障", "报错", "异常", "排查", "解决"],
        "会议安排": ["会议", "周会", "对齐", "同步", "讨论", "评审", "汇报"],
        "培训学习": ["培训", "文档", "教程", "学习", "资料", "分享"],
        "其他": []
    }
    
    for chat_id, group in chat_groups.items():
        chat_name = group["chat_name"]
        messages = group["messages"]
        if not messages:
            continue
        
        # 按主题分组消息
        topic_groups = {topic: [] for topic in topic_map.keys()}
        for msg in messages:
            content = msg["content"]
            matched_topic = "其他"
            # 匹配主题
            for topic, keywords in topic_map.items():
                if any(kw in content for kw in keywords):
                    matched_topic = topic
                    break
            topic_groups[matched_topic].append(msg)
        
        # 每个主题生成单独的汇总记录
        for topic, msgs in topic_groups.items():
            if not msgs:
                continue
            
            # 生成该主题的总结
            summary_parts = []
            todo_items = []
            for msg in msgs:
                content = msg["content"]
                summary_parts.append(f"{msg['sender']}: {content}")
                # 提取待办
                if any(kw in content for kw in ["要", "需要", "待办", "提交", "完成", "准备", "下周", "明天", "后续"]):
                    todo_items.append(content)
            
            # 拼接总结内容
            topic_title = f"[{chat_name}] {topic}："
            if len(summary_parts) == 1:
                summary_content = f"{topic_title} {summary_parts[0]}"
            else:
                summary_content = f"{topic_title}\n" + "\n".join([f"- {p}" for p in summary_parts[:8]])
                if len(summary_parts) > 8:
                    summary_content += f"\n- ...（共{len(summary_parts)}条消息）"
                if todo_items:
                    summary_content += f"\n待办：{'; '.join(todo_items)}"
            
            # 构造新的消息对象
            summary_msg = {
                "chat": {"name": chat_name, "type": group["chat_type"]},
                "message": {
                    "sender": {"name": "会话汇总", "id": MY_USER_ID},
                    "type": "text",
                    "content": {"text": {"content": summary_content}}
                },
                "content": summary_content
            }
            final_messages.append(summary_msg)
    
    if verbose:
        print(f"  总聊天记录：{len(processed_messages)} 条，过滤后有效：{len(final_messages)} 条")
    return final_messages
    return messages


def get_meetings(start_dt: datetime, end_dt: datetime, verbose: bool = False) -> list:
    """获取时间范围内的会议记录。"""
    start_iso = start_dt.strftime("%Y-%m-%dT%H:%M:%S+08:00")
    end_iso = end_dt.strftime("%Y-%m-%dT%H:%M:%S+08:00")

    if verbose:
        print(f"  获取会议记录：{start_iso} ~ {end_iso}")

    result = run_wps365_skill("meeting", "list",
                              "--start", start_iso,
                              "--end", end_iso)

    if result.get("error"):
        print(f"  ⚠️  获取会议记录失败：{result['error']}", file=sys.stderr)
        return []

    meetings = _extract_json_list_from_output(result.get("stdout", ""))
    if verbose:
        print(f"  找到 {len(meetings)} 个会议")
    return meetings

def get_documents(start_dt: datetime, end_dt: datetime, verbose: bool = False) -> list:
    """获取时间范围内的文档查看/编辑记录。"""
    if verbose:
        print(f"  获取最近访问/编辑的文档记录")

    # 直接调用命令行，绕过封装函数
    import subprocess
    cmd = [
        "python", "/root/.openclaw/skills/wps365-skill/skills/drive/run.py",
        "latest"
    ]
    try:
        stdout = subprocess.check_output(cmd, timeout=10).decode("utf-8")
    except Exception as e:
        print(f"  ⚠️  获取文档记录失败：{e}", file=sys.stderr)
        return []
    # 直接手动解析返回结果，跳过提取函数
    import re
    import json
    import dateutil.parser
    filtered_docs = []
    start_ts = start_dt.timestamp()
    end_ts = end_dt.timestamp()
    
    # 找JSON块
    m = re.search(r"```(?:json)?\s*(.*?)\s*```", stdout, re.DOTALL)
    if not m:
        m = re.search(r"{.*}", stdout, re.DOTALL)
    if m:
        try:
            data = json.loads(m.group(1))
            docs = data.get("items", [])
            for doc in docs:
                try:
                    # 外层的ctime就是访问时间
                    time_str = doc.get("ctime") or doc.get("mtime") or doc.get("file", {}).get("ctime")
                    if time_str:
                        dt = dateutil.parser.isoparse(time_str)
                        dt_ts = dt.timestamp()
                        if start_ts <= dt_ts <= end_ts:
                            # 重新构造内容，提取文件名、链接等信息
                            file_info = doc.get("file", {})
                            doc["content"] = f"查看文档：{file_info.get('name', '未知文档')}，链接：{file_info.get('link_url', '')}"
                            filtered_docs.append(doc)
                except Exception as e:
                    pass
        except Exception as e:
            print(f"  ⚠️  文档JSON解析失败：{e}")
    
    if verbose:
        print(f"  找到 {len(filtered_docs)} 个今日的文档操作记录")
    return filtered_docs


def _extract_json_list_from_output(text: str, key: str = "items") -> list:
    """从 WPS365 skill 输出中提取 JSON 数组，优先从指定key里提取。"""
    import re
    # 先移除所有markdown格式，只保留原始JSON
    cleaned_text = re.sub(r"^##.*$", "", text, flags=re.MULTILINE)
    cleaned_text = re.sub(r"^>.*$", "", cleaned_text, flags=re.MULTILINE)
    cleaned_text = re.sub(r"^\s*$", "", cleaned_text, flags=re.MULTILINE)
    
    # 尝试提取 ```json ... ``` 块
    m = re.search(r"```(?:json)?\s*(.*?)\s*```", cleaned_text, re.DOTALL)
    if m:
        try:
            data = json.loads(m.group(1))
            if isinstance(data, dict) and key in data:
                return data[key]
            elif isinstance(data, list):
                return data
        except json.JSONDecodeError:
            pass
    # 尝试直接提取JSON
    try:
        # 先找完整JSON对象
        m = re.search(r"{.*}", cleaned_text, re.DOTALL)
        if m:
            data = json.loads(m.group(0))
            if isinstance(data, dict) and key in data:
                return data[key]
        # 再找JSON数组
        m = re.search(r"\[.*\]", cleaned_text, re.DOTALL)
        if m:
            return json.loads(m.group(0))
    except json.JSONDecodeError:
        pass
    return []


def analyze_item(llm: LLMClient, item: dict, item_type: str) -> dict:
    """对单条消息/会议调用大模型分析。"""
    if item_type == "message":
        content = item.get("content") or item.get("text") or str(item)
        time_str = item.get("created_at") or item.get("time") or ""
        sender = item.get("sender_name") or item.get("sender") or ""
        chat_name = item.get("chat_name") or item.get("group_name") or ""
    else:  # meeting
        content = (
            f"会议主题：{item.get('subject', '')}\n"
            f"参与人：{item.get('participants', '')}\n"
            f"会议纪要：{item.get('summary', item.get('description', ''))}"
        )
        time_str = item.get("start_time") or item.get("created_at") or ""
        sender = ""
        chat_name = item.get("subject") or "会议"

    llm_req = llm.analyze_message(
        message_content=content,
        message_time=time_str,
        sender=sender,
        chat_name=chat_name,
    )
    return {"_prompt": llm_req["_prompt"], "_raw": item, "_type": item_type, "_content": content}


def is_duplicate(llm: LLMClient, new_content: str, existing_records: list, threshold: float = 0.7) -> bool:
    """检查新内容是否与已有记录重复（规则匹配，无需大模型）。"""
    if not existing_records:
        return False
    import difflib
    # 提取新内容的核心部分，去掉前缀的[群名] 发件人：
    new_core = new_content.split("：", maxsplit=1)[-1].strip() if "：" in new_content else new_content.strip()
    if len(new_core) < 5:
        return False
    # 只与最近20条比较
    for rec in existing_records[-20:]:
        fields = rec.get("fields", {})
        existing_content = fields.get("内容") or fields.get("跟进内容") or ""
        if not existing_content:
            continue
        # 提取已有内容的核心部分
        existing_core = existing_content.split("：", maxsplit=1)[-1].strip() if "：" in existing_content else existing_content.strip()
        if len(existing_core) < 5:
            continue
        # 计算相似度
        similarity = difflib.SequenceMatcher(None, new_core, existing_core).ratio()
        if similarity >= threshold:
            return True
    return False


def write_to_tables(tables: KingWorkTables, analysis: dict, content: str,
                    item: dict, item_type: str, dry_run: bool = False) -> list:
    """根据分析结果写入多维表。"""
    work_type = analysis.get("work_type", "")
    extracted = analysis.get("extracted_info", {})
    source = "AI自动分析"
    results = []

    # 最严格判断：必须明确提取到具体客户名称 或 具体项目名称才写入日记记录和业务表
    has_customer = False
    has_project = False
    # 排除通用模糊词，这些不算具体名称
    invalid_generic = ["客户", "项目", "公司", "集团", "甲方", "合作方", "供应商", "POC", "方案", "需求"]
    extracted_customer = extracted.get("customer", "").strip()
    extracted_project = extracted.get("project", "").strip()
    
    if extracted_customer and len(extracted_customer) >=4 and extracted_customer not in invalid_generic:
        has_customer = True
    if extracted_project and len(extracted_project) >=4 and extracted_project not in invalid_generic:
        has_project = True
    
    # dry run和正式逻辑对齐
    if dry_run:
        if has_customer or has_project:
            return [f"[DRY RUN] 将写入 {work_type}（符合要求：有具体客户/项目）"]
        else:
            return [f"[DRY RUN] 跳过 {work_type}（无具体客户/项目，仅生成惊喜记录）"]
    
    # 调试打印已关闭
    
    diary_id = ""
    if has_customer or has_project:
        # 写入日记记录（主表）
        try:
            diary_rec = tables.create_diary_record(
                content=content,
                work_type=work_type,
                customer=extracted.get("customer") or "",
                project=extracted.get("project") or "",
                tags=extracted.get("tags") or [],
                note=extracted.get("summary") or content[:100],
                source=source,
            )
            diary_id = diary_rec.get("id") if diary_rec else ""
            results.append(f"✅ 日记记录：{work_type}")
        except Exception as e:
            results.append(f"❌ 日记记录写入失败：{e}")
            diary_id = ""
    else:
        # 无客户/项目，仅生成惊喜记录，不同步业务表
        results.append(f"ℹ️  无关联客户/项目，仅生成惊喜记录")
    
    customer = extracted.get("customer") or ""
    project = extracted.get("project") or ""
    
    # 分发到业务表：只有有diary_id的时候才处理
    if diary_id:
        if work_type == "客户跟进" and customer:
            try:
                tables.create_customer_followup(
                    customer=customer,
                    content=content,
                    source=source,
                    diary_id=diary_id,
                )
                tables.update_customer_last_followup(customer)
                results.append(f"✅ 客户跟进：{customer}")
            except Exception as e:
                results.append(f"❌ 客户跟进写入失败：{e}")

        elif work_type == "学习成长":
            try:
                tables.create_learning_record(
                    topic=extracted.get("learning_topic") or content[:50],
                    content=content,
                    source=source,
                    diary_id=diary_id,
                )
                results.append("✅ 学习成长记录")
            except Exception as e:
                results.append(f"❌ 学习记录写入失败：{e}")

        elif work_type == "横向支持":
            try:
                tables.create_support_record(
                    target=extracted.get("support_target") or "同事",
                    content=content,
                    source=source,
                    diary_id=diary_id,
                )
                results.append("✅ 横向支持记录")
            except Exception as e:
                results.append(f"❌ 横向支持写入失败：{e}")

        elif work_type == "团队事务":
            try:
                tables.create_team_record(
                    topic=content[:50],
                    content=content,
                    source=source,
                    diary_id=diary_id,
                )
                results.append("✅ 团队事务记录")
            except Exception as e:
                results.append(f"❌ 团队事务写入失败：{e}")

        elif work_type == "灵感记录":
            try:
                tables.create_idea_record(
                    content=content,
                    source=source,
                    diary_id=diary_id,
                )
                results.append("✅ 灵感记录")
            except Exception as e:
                results.append(f"❌ 灵感记录写入失败：{e}")

    # 惊喜内容：直接调用dbsheet接口写入，避免封装类问题
    if extracted.get("is_surprise") and extracted.get("surprise_reason"):
        import subprocess
        import json
        import re
        wps_skill_path = "/root/.openclaw/skills/wps365-skill"
        file_id = os.environ.get("KINGWORK_FILE_ID", "")
        if not file_id:
            return results
        # 从extracted里获取项目和客户
        project = extracted.get("project", "")
        customer = extracted.get("customer", "")
        try:
            if item_type == "message":
                # 惊喜沟通记录：从配置读取sheet_id
                sheet_id = tables.sheet_ids['surprise_communications']
                # 提取沟通对象
                chat_info = item.get("chat", {})
                chat_type = chat_info.get("type", "group")
                chat_name = chat_info.get("name", "未知")
                if chat_type == "p2p":
                    communication_person = chat_name
                else:
                    communication_person = item.get("message", {}).get("sender", {}).get("name", "未知")
                
                # 自动生成价值点摘要：提取核心内容
                value_point = ""
                core_keywords = ["需求", "方案", "报价", "申请", "部署", "问题", "会议", "培训", "待办", "客户", "项目"]
                sentences = re.split(r"[。\n；，?!]", content)
                for s in sentences:
                    s = s.strip()
                    if any(kw in s for kw in core_keywords) and len(s) > 5:
                        value_point += s + "；"
                if not value_point:
                    value_point = content[:80] + ("..." if len(content) > 80 else "")
                
                # 自动打标签：最多4个
                label_pool = ["客户需求", "方案报价", "资源申请", "部署调试", "问题排查", "会议安排", "培训学习", "内部沟通", "待办事项", "行业信息", "客户跟进", "政策通知", "产品讨论"]
                matched_labels = []
                for label in label_pool:
                    # 模糊匹配关键词：前2个字符、后2个字符匹配就算命中
                    if len(label) >=2 and (label[:2] in content or label[-2:] in content):
                        matched_labels.append(label)
                        if len(matched_labels) >=4:
                            break
                # 不足4个补其他
                while len(matched_labels) <4:
                    matched_labels.append("其他")
                
                record_data = [{
                    "沟通时间": today_str(),
                    "沟通对象": communication_person,
                    "沟通方式": "单聊" if chat_type == "p2p" else "群聊",
                    "惊喜内容": content,
                    "价值点": value_point,
                    "标签": matched_labels,
                    "相关客户": customer or "",
                    "相关项目": project or "",
                    "来源": "AI自动分析",
                    "关联日记ID": diary_id,
                    "创建时间": today_str()
                }]
                cmd = [
                    "python", f"{wps_skill_path}/skills/dbsheet/run.py",
                    "create-records", file_id, str(sheet_id),
                    "--json", json.dumps(record_data, ensure_ascii=False)
                ]
                subprocess.check_output(cmd, timeout=10)
                results.append(f"✨ 惊喜沟通记录：已创建（对象：{communication_person}）")
            elif item_type == "doc":
                # 惊喜文档记录：从配置读取sheet_id
                sheet_id = tables.sheet_ids['surprise_docs']
                doc_name = item.get("file", {}).get("name") or "未知文档"
                doc_url = item.get("file", {}).get("link_url") or ""
                file_type = doc_name.split(".")[-1] if "." in doc_name else "未知"
                
                # 自动生成价值点
                doc_value_point = f"查看{file_type}类型文档：{doc_name}"
                # 自动打文档相关标签
                doc_label_pool = ["产品文档", "技术方案", "客户资料", "项目资料", "培训材料", "政策文档", "合同文件", "其他文档"]
                matched_doc_labels = []
                for label in doc_label_pool:
                    # 模糊匹配：前2个字符匹配就算命中
                    if len(label) >=2 and label[:2] in doc_name:
                        matched_doc_labels.append(label)
                        if len(matched_doc_labels) >=4:
                            break
                while len(matched_doc_labels) <4:
                    matched_doc_labels.append("其他")
                
                # 识别文档类型
                doc_type = "其他"
                type_map = {
                    "方案": "解决方案",
                    "案例": "案例研究",
                    "产品": "产品文档",
                    "报告": "行业报告",
                    "培训": "培训材料",
                    "白皮书": "产品文档",
                    "说明书": "产品文档"
                }
                for kw, t in type_map.items():
                    if kw in doc_name:
                        doc_type = t
                        break
                
                record_data = [{
                    "发现时间": today_str(),
                    "文档名称": doc_name,
                    "文档链接": doc_url,
                    "文档类型": doc_type,
                    "惊喜原因": extracted["surprise_reason"] + " | " + doc_value_point,
                    "相关客户": customer or "",
                    "相关项目": project or "",
                    "标签": matched_doc_labels,
                    "来源": "AI自动分析",
                    "关联日记ID": diary_id,
                    "创建时间": today_str()
                }]
                cmd = [
                    "python", f"{wps_skill_path}/skills/dbsheet/run.py",
                    "create-records", file_id, str(sheet_id),
                    "--json", json.dumps(record_data, ensure_ascii=False)
                ]
                subprocess.check_output(cmd, timeout=10)
                results.append(f"✨ 惊喜文档记录：已创建（{doc_name}）")
        except Exception as e:
            results.append(f"❌ 惊喜记录写入失败：{e}")

    return results


def main():
    args = parse_args()
    start_dt, end_dt, date_range_str = get_date_range(args)
    llm = LLMClient()

    print(f"\n## KingAuto - AI 自动分析")
    print(f"分析范围：{date_range_str}\n")

    # 获取数据
    all_items = []
    step = 1
    if args.source in ("chat", "all"):
        print(f"### 第{step}步：获取聊天记录")
        step +=1
        messages = get_chat_messages(start_dt, end_dt, verbose=args.verbose)
        for m in messages:
            all_items.append(("message", m))

    if args.source in ("meeting", "all"):
        print(f"### 第{step}步：获取会议记录")
        step +=1
        meetings = get_meetings(start_dt, end_dt, verbose=args.verbose)
        for m in meetings:
            all_items.append(("meeting", m))
    
    if args.source in ("doc", "all"):
        print(f"### 第{step}步：获取文档记录")
        step +=1
        docs = get_documents(start_dt, end_dt, verbose=args.verbose)
        for d in docs:
            all_items.append(("doc", d))

    print(f"\n共获取 {len(all_items)} 条记录，原始数据如下：\n")
    
    # 分类打印原始数据
    chat_items = [item for t, item in all_items if t == "message"]
    meeting_items = [item for t, item in all_items if t == "meeting"]
    doc_items = [item for t, item in all_items if t == "doc"]
    
    if chat_items:
        print("💬 聊天记录（共{}条）：".format(len(chat_items)))
        for idx, item in enumerate(chat_items, 1):
            print(f"  {idx}. {item.get('content', str(item)[:100])}")
        print()
    
    if meeting_items:
        print("📅 会议记录（共{}条）：".format(len(meeting_items)))
        for idx, item in enumerate(meeting_items, 1):
            title = item.get("subject") or item.get("title", "未知会议")
            start_time = item.get("start_time", "")
            print(f"  {idx}. {title} | 时间：{start_time}")
        print()
    
    if doc_items:
        print("📄 文档操作记录（共{}条）：".format(len(doc_items)))
        for idx, item in enumerate(doc_items, 1):
            print(f"  {idx}. {item.get('content', str(item)[:100])}")
        print()
    
    # 等待用户确认（非交互环境默认自动继续）
    import sys
    if sys.stdin.isatty():
        confirm = input("是否继续分析并写入多维表？(y/n，默认y)：").strip().lower()
        if confirm and confirm not in ["y", "yes"]:
            print("已取消操作，未写入任何数据。")
            return
    else:
        print("非交互环境，自动继续分析处理...")
    
    print("\n开始分析处理并写入...\n")

    if not all_items:
        print("没有找到需要分析的记录。")
        return

    # 获取已有日记记录（用于去重）
    try:
        tables = KingWorkTables()
        existing_records = tables.get_records_in_period(
            "diary_records", "记录时间", start_dt, end_dt
        )
        print(f"已有 {len(existing_records)} 条日记记录（用于去重）\n")
    except Exception as e:
        print(f"⚠️  获取已有记录失败：{e}", file=sys.stderr)
        tables = None
        existing_records = []

    # 逐条分析
    stats = {
        "processed": 0, 
        "skipped_dedup": 0, 
        "written": 0, 
        "errors": 0,
        "customer_followup": 0,
        "todo": 0,
        "learning": 0,
        "support": 0,
        "team": 0,
        "idea": 0,
        "surprise": 0
    }

    for i, (item_type, item) in enumerate(all_items):
        print(f"\n### 分析第 {i+1}/{len(all_items)} 条（{item_type}）")
        content = item.get("content") or item.get("subject") or str(item)[:200]

        # 自动调用大模型分析，无需交互，失败则降级为规则匹配
        llm_req = analyze_item(llm, item, item_type)
        content = llm_req["_content"]
        analysis = {"is_work_related": True, "work_type": "客户跟进", "confidence": 0.7, "extracted_info": {}}
        
        # 规则匹配逻辑，和kingrecord一致
        import re
        # 先过滤无效内容
        invalid_keywords = ["好的", "收到", "谢谢", "感谢", "ok", "OK", "是的", "对的", "没问题", "👍", "❤️", "😂", "🙏", "表情", "动画表情", "[图片]", "[文件]", "[语音]", "[视频]", "嗯嗯", "哦哦", "哈哈", "嗯嗯好", "好哒"]
        if len(content) < 5 or any(kw in content for kw in invalid_keywords):
            analysis["is_work_related"] = False
            analysis["skip_reason"] = "无效内容"
            return analysis
        
        # 判断是否工作相关
        work_keywords = ["客户", "项目", "会议", "方案", "需求", "培训", "支持", "学习", "待办", "跟进", "沟通", "拜访", "汇报", "讨论", "文档", "对接", "对齐", "同步", "方案", "报价", "招标", "投标"]
        is_work = any(kw in content for kw in work_keywords)
        if not is_work:
            analysis["is_work_related"] = False
            analysis["skip_reason"] = "非工作内容"
        else:
            # 分类
            if any(kw in content for kw in ["客户", "拜访", "沟通", "需求", "方案", "报价"]):
                analysis["work_type"] = "客户跟进"
            elif any(kw in content for kw in ["待办", "任务", "完成", "提交"]):
                analysis["work_type"] = "待办事项"
            elif any(kw in content for kw in ["学习", "培训", "研究", "文档"]):
                analysis["work_type"] = "学习成长"
            elif any(kw in content for kw in ["支持", "协助", "帮忙", "跨部门"]):
                analysis["work_type"] = "横向支持"
            elif any(kw in content for kw in ["会议", "团队", "团建", "分享"]):
                analysis["work_type"] = "团队事务"
            
            # 提取信息
            extracted = {}
            # 提取客户
            customer_match = re.search(r"(和|与|跟|拜访|见)(.*?)(公司|集团|股份|银行|证券|保险|科技|建投|建筑|总|客户)", content)
            if customer_match:
                extracted["customer"] = customer_match.group(2).strip() + customer_match.group(3).strip()
            # 提取项目
            project_match = re.search(r"(项目|需求|方案|系统|平台)(.*?)(的|，|。| )", content)
            if project_match:
                extracted["project"] = project_match.group(2).strip() + "项目"
            # 提取待办
            if any(kw in content for kw in ["提交", "完成", "准备", "出具", "发送", "需要", "要"]):
                todo_match = re.search(r"(提交|完成|准备|出具|发送|需要|要)(.*?)(，|。|$)", content)
                if todo_match:
                    extracted["todo"] = todo_match.group(1).strip() + todo_match.group(2).strip()
            # 惊喜内容识别：不过滤，所有有效记录都标记为惊喜
            extracted["is_surprise"] = True
            # 按类型自动设置惊喜原因
            if item_type == "message":
                extracted["surprise_reason"] = f"聊天记录自动采集：{content[:30]}..."
            elif item_type == "meeting":
                extracted["surprise_reason"] = f"会议记录自动采集：{content[:30]}..."
            elif item_type == "doc":
                extracted["surprise_reason"] = f"文档操作自动采集：{content[:30]}..."
            analysis["extracted_info"] = extracted

        if not analysis.get("is_work_related", False):
            if args.verbose:
                print(f"  ⏭️  非工作相关，跳过")
            continue

        stats["processed"] += 1

        # 相似度去重
        if not args.no_dedup and tables:
            if is_duplicate(llm, content, existing_records):
                print(f"  🔄 与已有记录重复，跳过")
                stats["skipped_dedup"] += 1
                continue

        # 写入多维表
        if tables:
            results = write_to_tables(
                tables, analysis, content, item, item_type, dry_run=args.dry_run
            )
            for r in results:
                print(f"  {r}")
            stats["written"] += 1
            
            # 按类型统计数量
            work_type = analysis.get("work_type", "")
            if work_type == "客户跟进":
                stats["customer_followup"] += 1
            elif work_type == "待办事项":
                stats["todo"] += 1
            elif work_type == "学习成长":
                stats["learning"] += 1
            elif work_type == "横向支持":
                stats["support"] += 1
            elif work_type == "团队事务":
                stats["team"] += 1
            elif work_type == "灵感记录":
                stats["idea"] += 1
            
            # 统计惊喜内容
            if analysis.get("extracted_info", {}).get("is_surprise"):
                stats["surprise"] += 1

    # 汇总
    print(f"\n## 分析完成")
    print(f"- 处理：{stats['processed']} 条")
    print(f"- 去重跳过：{stats['skipped_dedup']} 条")
    print(f"- 写入：{stats['written']} 条")
    if stats["errors"]:
        print(f"- 错误：{stats['errors']} 条")
    
    # 输出结构化执行摘要
    if stats["written"] > 0:
        print("\n📝 本次执行摘要：")
        print(f"   📅 分析范围：{date_range_str}")
        print(f"   ✅ 总写入记录：{stats['written']} 条")
        if stats["customer_followup"] > 0:
            print(f"   👥 客户跟进记录：{stats['customer_followup']} 条")
        if stats["todo"] > 0:
            print(f"   ✅ 待办任务：{stats['todo']} 条")
        if stats["learning"] > 0:
            print(f"   📚 学习成长记录：{stats['learning']} 条")
        if stats["support"] > 0:
            print(f"   🤝 横向支持记录：{stats['support']} 条")
        if stats["team"] > 0:
            print(f"   👪 团队事务记录：{stats['team']} 条")
        if stats["idea"] > 0:
            print(f"   💡 灵感记录：{stats['idea']} 条")
        if stats["surprise"] > 0:
            print(f"   ✨ 惊喜内容：{stats['surprise']} 条")
        print(f"   🔗 查看多维表：https://www.kdocs.cn/l/cmWhTjnp3SWB")


if __name__ == "__main__":
    main()
