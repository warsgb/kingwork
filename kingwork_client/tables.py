# -*- coding: utf-8 -*-
"""
KingWork 多维表操作封装。
通过调用 wpsv7client 的 dbsheet 函数实现记录的增删改查。
"""
import json
import sys
from typing import Optional, List, Dict, Any
from pathlib import Path

from .base import import_wpsv7client, get_file_id, get_sheet_ids, now_iso, today_str, get_enum_config, debug_log

# 子表ID到名称映射（便于日志可读性）
SHEET_KEY_NAME_MAP = {
    "diary_records": "日记记录表",
    "todo_records": "待办记录表",
    "customer_profiles": "客户档案表",
    "project_profiles": "项目档案表",
    "customer_followups": "客户跟进记录表",
    "learning_records": "学习成长记录表",
    "support_records": "横向支持记录表",
    "team_records": "团队事务记录表",
    "idea_records": "灵感记录表",
    "surprise_docs": "20惊喜文档记录表",
    "surprise_communications": "21惊喜沟通记录表",
    "surprise_meetings": "22惊喜会议记录表",
    "event_reception": "活动接待记录表"
}


class KingWorkTables:
    """封装 KingWork 多维表操作。"""

    def __init__(self, file_id: str = None, sheet_ids: dict = None):
        # 使用优化的导入函数（优先直接导入，失败则回退到路径导入）
        wpsv7client = import_wpsv7client()
        self._get_schema = wpsv7client.dbsheet_get_schema
        self._list_records = wpsv7client.dbsheet_list_records
        self._create_records = wpsv7client.dbsheet_batch_create_records
        self._update_records = wpsv7client.dbsheet_batch_update_records
        self._delete_records = wpsv7client.dbsheet_batch_delete_records
        self._create_sheet = wpsv7client.dbsheet_create_sheet
        self._create_view = wpsv7client.dbsheet_create_view

        self.file_id = file_id or get_file_id()
        self.sheet_ids = sheet_ids or get_sheet_ids()
        self.enum_config = get_enum_config()

    def _sid(self, key: str) -> str:
        """根据 key 获取 sheet_id，找不到则报错。"""
        sid = self.sheet_ids.get(key)
        if not sid:
            raise ValueError(f"未找到数据表 '{key}' 的 sheet_id，请先运行 init_tables.py")
        return str(sid)

    def _check(self, resp: dict) -> dict:
        """检查 API 响应是否成功。"""
        if resp.get("code") != 0:
            msg = resp.get("msg") or resp.get("message") or "未知错误"
            raise RuntimeError(f"WPS API 错误: {msg}")
        return resp.get("data") or resp

    # ─── 通用增删改查 ─────────────────────────────────────────────

    def list_records(self, sheet_key: str, page_size: int = 100, page_token: str = None) -> List[dict]:
        """列举记录，返回 record 列表。"""
        sid = self._sid(sheet_key)
        kwargs = {"page_size": page_size}
        if page_token:
            kwargs["page_token"] = page_token
        resp = self._list_records(self.file_id, sid, **kwargs)
        data = self._check(resp)
        return data.get("records") or []

    def list_all_records(self, sheet_key: str, filter_body: dict = None) -> List[dict]:
        """遍历所有页，返回全量记录（支持服务端筛选）。"""
        all_records = []
        page_token = None
        while True:
            sid = self._sid(sheet_key)
            kwargs = {"page_size": 100}
            if page_token:
                kwargs["page_token"] = page_token
            if filter_body:
                kwargs["filter_body"] = filter_body
            resp = self._list_records(self.file_id, sid, **kwargs)
            data = self._check(resp)
            records = data.get("records") or []
            all_records.extend(records)
            page_token = data.get("page_token")
            if not page_token or not records:
                break
        return all_records

    def create_record(self, sheet_key: str, fields: dict) -> Optional[dict]:
        """创建单条记录，返回新记录。"""
        results = self.create_records(sheet_key, [fields])
        return results[0] if results else None

    def create_records(self, sheet_key: str, records: List[dict]) -> List[dict]:
        """批量创建记录，records 为字段字典列表，返回新记录列表。"""
        sid = self._sid(sheet_key)
        # 调试日志输出
        sheet_name = SHEET_KEY_NAME_MAP.get(sheet_key, sheet_key)
        debug_log(f"执行写入操作：")
        debug_log(f"  多维表FILE_ID：{self.file_id}")
        debug_log(f"  子表KEY：{sheet_key} | 子表名称：{sheet_name} | 子表ID：{sid}")
        debug_log(f"  写入数据：{records[:3]}{'（省略多余' + str(len(records)-3) + '条）' if len(records) >3 else ''}")
        # 枚举值校验与自动修正
        enum_fields = self.enum_config.get(sheet_key, {})
        if enum_fields:
            for record in records:
                for field_name, field_value in list(record.items()):
                    if field_name not in enum_fields:
                        continue
                    allowed = enum_fields[field_name]
                    if isinstance(field_value, list):
                        # MultiSelect：只保留在允许列表中的项，全部无效时用默认值
                        valid_items = [v for v in field_value if v in allowed]
                        if not valid_items:
                            record[field_name] = allowed[0]
                    elif field_value not in allowed:
                        # SingleSelect：直接不匹配则替换默认值
                        record[field_name] = allowed[0]
        # 将字段字典包装为 API 要求的格式
        payload = [{"fields_value": json.dumps(r, ensure_ascii=False)} for r in records]
        resp = self._create_records(self.file_id, sid, payload)
        data = self._check(resp)
        return data.get("records") or []

    def update_record(self, sheet_key: str, record_id: str, fields: dict) -> bool:
        """更新单条记录。"""
        return self.update_records(sheet_key, [{"id": record_id, "fields": fields}])

    def update_records(self, sheet_key: str, updates: List[dict]) -> bool:
        """批量更新记录。updates 为 [{"id": "...", "fields": {...}}] 列表。"""
        sid = self._sid(sheet_key)
        # 枚举值校验与自动修正
        enum_fields = self.enum_config.get(sheet_key, {})
        if enum_fields:
            for update in updates:
                fields = update["fields"]
                for field_name, field_value in list(fields.items()):
                    if field_name not in enum_fields:
                        continue
                    allowed = enum_fields[field_name]
                    if isinstance(field_value, list):
                        valid_items = [v for v in field_value if v in allowed]
                        if not valid_items:
                            fields[field_name] = allowed[0]
                    elif field_value not in allowed:
                        fields[field_name] = allowed[0]
        payload = [
            {"id": u["id"], "fields_value": json.dumps(u["fields"], ensure_ascii=False)}
            for u in updates
        ]
        resp = self._update_records(self.file_id, sid, payload)
        self._check(resp)
        return True

    def get_record_fields(self, record: dict) -> dict:
        """从 record 解析出字段字典。"""
        raw = record.get("fields")
        if raw is None:
            return {}
        if isinstance(raw, dict):
            return raw
        try:
            return json.loads(raw) if isinstance(raw, str) else {}
        except json.JSONDecodeError:
            return {}

    # ─── 业务级操作 ──────────────────────────────────────────────

    def create_diary_record(self, content: str, work_type: str, **kwargs) -> Optional[dict]:
        """创建日记记录。"""
        fields = {
            "记录时间": kwargs.get("record_time", today_str()),
            "内容": content,
            "工作类型": work_type,
            "创建时间": today_str(),
            "来源": kwargs.get("source", "手动输入"),
        }
        if kwargs.get("customer"):
            fields["关联客户"] = kwargs["customer"]
        if kwargs.get("project"):
            fields["关联项目"] = kwargs["project"]
        if kwargs.get("tags"):
            fields["标签"] = kwargs["tags"]
        if kwargs.get("note"):
            fields["备注"] = kwargs["note"]
        return self.create_record("diary_records", fields)

    def create_todo_record(self, task_name: str, **kwargs) -> Optional[dict]:
        """创建待办记录。"""
        fields = {
            "任务名称": task_name,
            "优先级": kwargs.get("priority", "中"),
            "状态": "待处理",
            "创建时间": today_str(),
        }
        if kwargs.get("description"):
            fields["任务描述"] = kwargs["description"]
        if kwargs.get("due_date"):
            fields["到期时间"] = kwargs["due_date"]
        if kwargs.get("customer"):
            fields["关联客户"] = kwargs["customer"]
        if kwargs.get("project"):
            fields["关联项目"] = kwargs["project"]
        return self.create_record("todo_records", fields)

    def create_customer_followup(self, customer: str, content: str, **kwargs) -> Optional[dict]:
        """创建客户跟进记录。"""
        fields = {
            "跟进时间": kwargs.get("followup_time", today_str()),
            "客户名称": customer,
            "跟进方式": kwargs.get("method", "其他"),
            "跟进内容": content,
            "来源": kwargs.get("source", "手动输入"),
            "创建时间": today_str(),
        }
        if kwargs.get("result"):
            fields["跟进结果"] = kwargs["result"]
        if kwargs.get("next_time"):
            fields["下次跟进时间"] = kwargs["next_time"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        if kwargs.get("meeting_link"):
            fields["会议链接"] = kwargs["meeting_link"]
        return self.create_record("customer_followups", fields)

    def create_learning_record(self, topic: str, content: str, **kwargs) -> Optional[dict]:
        """创建学习成长记录。"""
        fields = {
            "学习时间": kwargs.get("learn_time", today_str()),
            "学习主题": topic,
            "学习内容": content,
            "来源": kwargs.get("source", "手动输入"),
            "创建时间": today_str(),
        }
        if kwargs.get("learning_type"):
            fields["学习类型"] = kwargs["learning_type"]
        if kwargs.get("duration_hours") is not None:
            fields["学习时长"] = kwargs["duration_hours"]
        if kwargs.get("resource"):
            fields["学习资源"] = kwargs["resource"]
        if kwargs.get("key_takeaway"):
            fields["关键收获"] = kwargs["key_takeaway"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        return self.create_record("learning_records", fields)

    def create_support_record(self, target: str, content: str, **kwargs) -> Optional[dict]:
        """创建横向支持记录。"""
        fields = {
            "支持时间": kwargs.get("support_time", today_str()),
            "支持对象": target,
            "支持内容": content,
            "来源": kwargs.get("source", "手动输入"),
            "创建时间": today_str(),
        }
        if kwargs.get("support_type"):
            fields["支持类型"] = kwargs["support_type"]
        if kwargs.get("result"):
            fields["支持结果"] = kwargs["result"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        return self.create_record("support_records", fields)

    def create_team_record(self, topic: str, content: str, **kwargs) -> Optional[dict]:
        """创建团队事务记录。"""
        fields = {
            "事务时间": kwargs.get("event_time", today_str()),
            "事务主题": topic,
            "事务内容": content,
            "来源": kwargs.get("source", "手动输入"),
            "创建时间": today_str(),
        }
        if kwargs.get("event_type"):
            fields["事务类型"] = kwargs["event_type"]
        if kwargs.get("participants"):
            fields["参与人员"] = kwargs["participants"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        return self.create_record("team_records", fields)

    def create_idea_record(self, content: str, **kwargs) -> Optional[dict]:
        """创建灵感记录。"""
        fields = {
            "记录时间": kwargs.get("record_time", today_str()),
            "灵感内容": content,
            "来源": kwargs.get("source", "手动输入"),
            "创建时间": today_str(),
        }
        if kwargs.get("category"):
            fields["灵感类别"] = kwargs["category"]
        if kwargs.get("feasibility"):
            fields["可行性"] = kwargs["feasibility"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        return self.create_record("idea_records", fields)

    def create_event_reception(self, subject: str, content: str, **kwargs) -> Optional[dict]:
        """创建活动接待记录。"""
        fields = {
            "活动时间": kwargs.get("event_time", today_str()),
            "主题": subject,
            "内容": content,
            "来源": kwargs.get("source", "手动输入"),
            "创建时间": today_str(),
        }
        if kwargs.get("reception_target"):
            fields["接待对象"] = kwargs["reception_target"]
        if kwargs.get("event_type"):
            fields["活动类型"] = kwargs["event_type"]
        if kwargs.get("customer"):
            fields["关联客户"] = kwargs["customer"]
        if kwargs.get("project"):
            fields["关联项目"] = kwargs["project"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        if kwargs.get("remark"):
            fields["备注"] = kwargs["remark"]
        return self.create_record("event_reception", fields)

    def create_surprise_doc(self, doc_name: str, reason: str, **kwargs) -> Optional[dict]:
        """创建惊喜文档记录。"""
        fields = {
            "发现时间": kwargs.get("found_time", now_iso()),
            "文档名称": doc_name,
            "惊喜原因": reason,
            "来源": kwargs.get("source", "手动输入"),
            "创建时间": now_iso(),
        }
        if kwargs.get("doc_link"):
            fields["文档链接"] = kwargs["doc_link"]
        if kwargs.get("doc_type"):
            fields["文档类型"] = kwargs["doc_type"]
        if kwargs.get("customer"):
            fields["相关客户"] = kwargs["customer"]
        if kwargs.get("project"):
            fields["相关项目"] = kwargs["project"]
        if kwargs.get("tags"):
            fields["标签"] = kwargs["tags"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        return self.create_record("surprise_docs", fields)

    def create_surprise_communication(self, person: str, content: str, value_point: str, **kwargs) -> Optional[dict]:
        """创建惊喜沟通记录。"""
        fields = {
            "沟通时间": kwargs.get("comm_time", now_iso()),
            "沟通对象": person,
            "惊喜内容": content,
            "价值点": value_point,
            "来源": kwargs.get("source", "AI自动分析"),
            "创建时间": now_iso(),
        }
        if kwargs.get("method"):
            fields["沟通方式"] = kwargs["method"]
        if kwargs.get("customer"):
            fields["相关客户"] = kwargs["customer"]
        if kwargs.get("project"):
            fields["相关项目"] = kwargs["project"]
        if kwargs.get("tags"):
            fields["标签"] = kwargs["tags"]
        if kwargs.get("diary_id"):
            fields["关联日记ID"] = kwargs["diary_id"]
        return self.create_record("surprise_communications", fields)

    def create_surprise_meeting(self, meeting_name: str, summary: str, participants: str = "", **kwargs) -> Optional[dict]:
        """创建惊喜会议记录。"""
        fields = {
            "发现时间": kwargs.get("found_time", today_str()),
            "会议名称": meeting_name,
            "参会人清单": participants,
            "会议摘要": summary,
            "会议链接": kwargs.get("meeting_url", ""),
            "会议ID": kwargs.get("meeting_id", ""),
            "来源": kwargs.get("source", "AI自动分析"),
            "创建时间": now_iso(),
        }
        if kwargs.get("customer"):
            fields["相关客户"] = kwargs["customer"]
        if kwargs.get("project"):
            fields["相关项目"] = kwargs["project"]
        if kwargs.get("tags"):
            fields["标签"] = kwargs["tags"]
        return self.create_record("surprise_meetings", fields)

    def update_customer_last_followup(self, customer_name: str) -> bool:
        """更新客户档案中的最近跟进时间。"""
        try:
            records = self.list_all_records("customer_profiles")
            for rec in records:
                fields = self.get_record_fields(rec)
                if fields.get("客户名称") == customer_name:
                    # 更新跟进次数和最近跟进时间
                    current_count = fields.get("跟进次数", 0) or 0
                    self.update_record("customer_profiles", rec["id"], {
                        "最近跟进时间": today_str(),
                        "跟进次数": int(current_count) + 1,
                    })
                    return True
        except Exception:
            pass
        return False

    def get_pending_todos(self) -> List[dict]:
        """获取未完成的待办事项，优先使用服务端筛选。"""
        # 服务端筛选：状态 NOT "已完成" AND NOT "已取消"
        filter_body = {
            "mode": "AND",
            "criteria": [
                {"field": "状态", "operator": "NotEqu", "values": ["已完成"]},
                {"field": "状态", "operator": "NotEqu", "values": ["已取消"]},
            ]
        }
        try:
            records = self.list_all_records("todo_records", filter_body=filter_body)
            return [{"id": rec["id"], "fields": self.get_record_fields(rec)} for rec in records]
        except Exception:
            pass

        # 降级：全量拉回 + 客户端过滤
        records = self.list_all_records("todo_records")
        result = []
        for rec in records:
            fields = self.get_record_fields(rec)
            status = fields.get("状态", "")
            if status not in ("已完成", "已取消"):
                result.append({"id": rec["id"], "fields": fields})
        return result

    def get_inactive_customers(self, days: int = 15) -> List[dict]:
        """获取超过 N 天未跟进的客户，优先使用服务端筛选。"""
        from datetime import datetime, timedelta, timezone
        tz_cst = timezone(timedelta(hours=8))
        cutoff = datetime.now(tz=tz_cst) - timedelta(days=days)
        cutoff_iso = cutoff.isoformat()

        # 服务端筛选：排除已成交/已流失，且最近跟进时间 < cutoff 或为空
        filter_body = {
            "mode": "AND",
            "criteria": [
                {"field": "客户状态", "operator": "NotEqu", "values": ["成交"]},
                {"field": "客户状态", "operator": "NotEqu", "values": ["流失"]},
            ]
        }
        try:
            records = self.list_all_records("customer_profiles", filter_body=filter_body)
        except Exception:
            # 服务端筛选失败，降级全量拉取
            records = self.list_all_records("customer_profiles")

        result = []
        for rec in records:
            fields = self.get_record_fields(rec)
            status = fields.get("客户状态", "")
            if status in ("成交", "流失"):
                continue
            last_time = fields.get("最近跟进时间")
            if not last_time:
                result.append({"id": rec["id"], "fields": fields, "days_inactive": 999})
                continue
            try:
                if isinstance(last_time, (int, float)):
                    dt = datetime.fromtimestamp(last_time / 1000, tz=tz_cst)
                else:
                    from dateutil import parser as dp
                    dt = dp.parse(str(last_time))
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=tz_cst)
                if dt < cutoff:
                    delta = datetime.now(tz=tz_cst) - dt
                    result.append({"id": rec["id"], "fields": fields, "days_inactive": delta.days})
            except Exception:
                result.append({"id": rec["id"], "fields": fields, "days_inactive": 999})
        return result

    def get_records_in_period(self, sheet_key: str, time_field: str, start_dt, end_dt) -> List[dict]:
        """获取指定时间范围内的记录，优先使用服务端筛选，失败时降级为客户端过滤。"""
        from datetime import datetime, timezone, timedelta
        tz_cst = timezone(timedelta(hours=8))

        # WPS Date 字段不支持服务端 filter（GreaterThan/LessThan 返回 Unknown enum value），
        # 只能全量拉取 + 客户端内存过滤。
        # 若 WPS 后续支持 Date filter，可在此恢复服务端筛选。
        records = self.list_all_records(sheet_key)
        result = []
        for rec in records:
            fields = self.get_record_fields(rec)
            t = fields.get(time_field)
            if not t:
                continue
            try:
                if isinstance(t, (int, float)):
                    dt = datetime.fromtimestamp(t / 1000, tz=tz_cst)
                else:
                    from dateutil import parser as dp
                    dt = dp.parse(str(t))
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=tz_cst)
                if start_dt <= dt <= end_dt:
                    result.append({"id": rec["id"], "fields": fields})
            except Exception:
                pass
        return result
