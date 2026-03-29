"""
kingbrowse 核心检索模块
- 基于 WPS 多维表 server-side 模糊搜索
- 支持多文件、多表配置
- 自动解析 Attachment 字段获取链接
"""
import re
import sys
import os
import json
import time
import logging
from typing import Optional

# kingwork + wps365 路径（动态获取，兼容 Mac/Linux/Windows）
from pathlib import Path
KINGWORK_ROOT = str(Path(__file__).resolve().parent.parent.parent)
sys.path.insert(0, KINGWORK_ROOT)

from kingwork_client.base import get_wps365_root as _get_wps365_root
WPS365_ROOT = str(_get_wps365_root())
if WPS365_ROOT not in sys.path:
    sys.path.insert(0, WPS365_ROOT)
os.chdir(KINGWORK_ROOT)

import kingwork_client.tables as kt
from wpsv7client import dbsheet_get_schema, dbsheet_list_records

log = logging.getLogger("kingbrowse")

# ------------------------------------------------------------------
# 工具函数
# ------------------------------------------------------------------

def _fuzzy_match(text: str, keyword: str) -> bool:
    """简单模糊匹配：keyword 是否出现在 text 中（不区分大小写）。"""
    if not text or not keyword:
        return False
    return keyword.lower() in str(text).lower()


def _resolve_link(field_value, field_type: str) -> Optional[str]:
    """
    从字段值中提取可访问的链接。
    - URL 类型：直接返回字符串
    - Attachment 类型：取数组第一个元素的 linkUrl
    - 其他：返回 None
    """
    if not field_value:
        return None

    if field_type == "Url":
        # Url 类型可能是字符串，也可能是 [{'address': '...', 'displayText': '...'}]
        if isinstance(field_value, str):
            return field_value.strip() or None
        if isinstance(field_value, list) and field_value:
            first = field_value[0]
            if isinstance(first, dict):
                return first.get("address") or first.get("url") or first.get("link") or str(first)
        return None

    if field_type in ("Attachment", "Attachments"):
        try:
            # Attachment 是 JSON 数组字符串或列表
            if isinstance(field_value, str):
                items = json.loads(field_value)
            else:
                items = field_value or []
            for item in items:
                link = item.get("linkUrl") or item.get("url") or ""
                if link:
                    return link
        except Exception:
            pass
        return None

    # MultiLineText / SingleLineText 等：如果内容像 URL，直接返回
    if isinstance(field_value, str):
        text = field_value.strip()
        if text and ("http://" in text or "https://" in text):
            # 提取第一个 URL
            for part in text.split():
                if part.startswith("http://") or part.startswith("https://"):
                    return part
            return text
        return None

    return None


def _get_all_text(value) -> str:
    """将任意字段值转为可搜索的纯文本字符串。"""
    if not value:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, list):
        texts = []
        for item in value:
            if isinstance(item, dict):
                # Attachment 数组：拼接文件名
                fn = item.get("fileName", "")
                if fn:
                    texts.append(fn)
                link = item.get("linkUrl", "")
                if link:
                    texts.append(link)
            else:
                texts.append(str(item))
        return " ".join(texts)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def _score_result(text: str, keyword: str) -> float:
    """
    计算匹配质量得分 0.0~1.0。
    """
    text_lower = str(text).lower()
    kw_lower = keyword.lower()

    # 完全匹配
    if text_lower == kw_lower:
        return 1.0
    # 开头匹配
    if text_lower.startswith(kw_lower):
        return 0.95
    # 包含匹配：按长度给分（越短越精确）
    if kw_lower in text_lower:
        ratio = len(kw_lower) / max(len(text_lower), 1)
        return 0.5 + ratio * 0.4  # 0.5~0.9
    return 0.0


# ------------------------------------------------------------------
# Schema 缓存（进程内）
# ------------------------------------------------------------------
# file_id -> {sheet_name -> {field_name -> field_type}}
_SCHEMA_CACHE: dict = {}
_SHEET_ID_CACHE: dict = {}  # file_id -> {sheet_name -> sheet_id}


def get_schema(file_id: str) -> dict:
    """获取并缓存指定文件的 schema。"""
    if file_id in _SCHEMA_CACHE:
        return _SCHEMA_CACHE[file_id]

    resp = dbsheet_get_schema(file_id)
    sheets = resp.get("data", {}).get("sheets", [])
    schema = {}
    id_map = {}
    for sheet in sheets:
        sname = sheet.get("name", "")
        sid = sheet.get("id", "")
        id_map[sname] = str(sid)
        field_map = {}
        for f in sheet.get("fields", []):
            field_map[f["name"]] = f.get("field_type", "Unknown")
        schema[sname] = field_map

    _SCHEMA_CACHE[file_id] = schema
    _SHEET_ID_CACHE[file_id] = id_map
    log.info(f"已缓存文件 {file_id} 的 schema（{len(schema)} 个表）")
    return schema


def get_sheet_id(file_id: str, sheet_name: str) -> Optional[str]:
    """获取 sheet_name 对应的 sheet_id。"""
    get_schema(file_id)  # 确保缓存已填充
    return _SHEET_ID_CACHE.get(file_id, {}).get(sheet_name)


# ------------------------------------------------------------------
# 单表搜索
# ------------------------------------------------------------------

def search_sheet(
    file_id: str,
    file_name: str,
    sheet_name: str,
    keyword: str,
    search_fields: list[str],
    name_field: str,
    link_fields: list[str],
    top_k: int = 10,
) -> list[dict]:
    """
    在单张表中模糊搜索 keyword。
    返回匹配结果列表。
    """
    results = []
    sheet_id = get_sheet_id(file_id, sheet_name)
    if not sheet_id:
        log.warning(f"未找到 sheet: {sheet_name}（file: {file_id}）")
        return results

    schema = _SCHEMA_CACHE.get(file_id, {}).get(sheet_name, {})

    # WPS server-side 过滤：第一个 search_field 用 Contains
    # 多个字段需要客户端过滤（WPS filter mode=OR 不支持跨字段 OR）
    primary_field = search_fields[0] if search_fields else name_field

    filter_body = {
        "mode": "OR",
        "criteria": [
            {"field": f, "operator": "contains", "values": [keyword]}
            for f in search_fields
        ]
    }

    for attempt in range(2):
        try:
            resp = dbsheet_list_records(
                file_id=file_id,
                sheet_id=int(sheet_id),
                filter_body={"filter": filter_body},
                page_size=100,
            )
            if resp.get("code") != 0:
                log.warning(f"查询失败 [{sheet_name}]: {resp.get('msg', '')}")
                return results
            break
        except Exception as e:
            if attempt == 0:
                time.sleep(0.5)
                continue
            log.warning(f"查询异常 [{sheet_name}]: {e}")
            return results

    raw_records = (resp.get("data") or {}).get("records", [])
    log.debug(f"[{sheet_name}] 粗筛返回 {len(raw_records)} 条")

    for rec in raw_records:
        fields = rec.get("fields") or {}
        if isinstance(fields, str):
            try:
                fields = json.loads(fields)
            except Exception:
                continue

        # ---- 名称 ----
        raw_name = fields.get(name_field, "")
        name = _get_all_text(raw_name).split("\n")[0].strip()[:100]
        if not name:
            continue

        # ---- 链接 ----
        doc_link = ""
        for lf in link_fields:
            ft = schema.get(lf, "Unknown")
            lv = fields.get(lf)
            if lv:
                resolved = _resolve_link(lv, ft)
                if resolved:
                    doc_link = resolved
                    break

        # ---- 匹配度评分 & snippet ----
        best_score = 0.0
        best_match_field = ""
        snippet_parts = []

        for sf in search_fields:
            fv = fields.get(sf, "")
            fv_text = _get_all_text(fv)
            score = _score_result(fv_text, keyword)
            if score > best_score:
                best_score = score
                best_match_field = sf
            if score > 0.3 and len(snippet_parts) < 2:
                # 取匹配字段的前50字作为摘要
                snippet_text = fv_text.strip()[:80].replace("\n", " ")
                if snippet_text:
                    snippet_parts.append(f"[{sf}] {snippet_text}")

        if best_score < 0.15:
            continue  # 匹配度过低，跳过

        # Relevance 等级
        if best_score >= 0.9:
            relevance = "高"
        elif best_score >= 0.6:
            relevance = "中"
        else:
            relevance = "低"

        snippet = (snippet_parts[0][len(best_match_field)+3:] if snippet_parts else name)

        results.append({
            "file_id": file_id,
            "file_name": file_name,
            "sheet_name": sheet_name,
            "record_id": rec.get("id", ""),
            "name": name,
            "link": doc_link,
            "snippet": snippet[:100],
            "match_field": best_match_field,
            "match_score": round(best_score, 2),
            "relevance": relevance,
        })

    # 按匹配度降序，取 top_k
    results.sort(key=lambda x: x["match_score"], reverse=True)
    return results[:top_k]


# ------------------------------------------------------------------
# LLM 关键词扩展
# ------------------------------------------------------------------

def _expand_keywords(keyword: str) -> list[str]:
    """
    将用户query交给LLM，生成3~5个精准搜索词。
    返回搜索词列表。
    """
    import kingwork_client.llm as kllm
    prompt = (
        "你是一个售前资料库搜索助手。请将用户的查询拆分成3到5个搜索词，"
        "优先从查询中提取已有词汇，其次补充同类相关词汇，覆盖不同角度（行业、产品、场景、客户类型等）。\n"
        f"用户查询：「{keyword}」\n"
        "输出格式：词语1 | 词语2 | 词语3（直接返回词语，用 | 分隔，不要解释，不要加引号）\n"
        "示例：输入「制造业典型场景资料」→ 输出「制造业 | 典型场景 | 制造业案例 | 典型客户」"
    )
    try:
        llm = kllm.LLMClient()
        # require_json=False 直接返回原始文本，不用 JSON 解析
        result = llm._call(prompt, require_json=False)
        raw = ""
        if isinstance(result, dict):
            raw = result.get("raw", "") or ""
        elif result is not None:
            raw = str(result)

        # 如果 raw 为空或包含 "None" literal，按标点分词
        if not raw.strip() or raw.strip().lower() in ("none", "null"):
            parts = [p.strip() for p in re.split(r'[,，、\s]+', keyword) if p.strip()]
            return parts[:5] if parts else [keyword]

        # 按 | 分隔
        terms = [t.strip() for t in raw.split("|") if t.strip() and t.strip().lower() not in ("none", "null")]
        if not terms:
            return [keyword]
        return terms[:5]
    except Exception:
        # 出错时按空格/逗号分词
        parts = [p.strip() for p in re.split(r'[,，、\s]+', keyword) if p.strip()]
        return parts[:5] if parts else [keyword]


# ------------------------------------------------------------------
# 主搜索入口
# ------------------------------------------------------------------

def search(query: str, top_k_per_source: int = 5) -> list[dict]:
    """
    在所有已配置的来源中搜索 query。
    自动调用LLM扩展关键词，分别搜索后合并结果。
    返回合并去重后的结果列表，按匹配度降序。
    """
    import yaml

    # 1. LLM 扩展关键词
    keywords = _expand_keywords(query)
    log.info(f"LLM扩展关键词：{keywords}")

    config_path = os.path.join(os.path.dirname(__file__), "config.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    all_results = []
    seen = set()  # 全局去重：(file_id, record_id, keyword)
    sources = config.get("browse_sources", [])

    for source in sources:
        file_id = source.get("file_id", "")
        file_name = source.get("file_name", file_id)
        sheet_configs = source.get("sheets", [])

        if not sheet_configs:
            # 未指定 sheets，全文件搜索
            schema = get_schema(file_id)
            sheet_configs = [
                {
                    "sheet_name": sname,
                    "search_fields": list(fmap.keys()),
                    "name_field": list(fmap.keys())[0] if fmap else "名称",
                    "link_fields": [
                        fname for fname, ftype in fmap.items()
                        if ftype in ("Url", "Attachment", "Attachments")
                    ],
                }
                for sname, fmap in schema.items()
            ]

        for sheet_cfg in sheet_configs:
            sname = sheet_cfg.get("sheet_name")
            if not sname:
                continue

            search_fields = sheet_cfg.get("search_fields", [])
            name_field = sheet_cfg.get("name_field", search_fields[0] if search_fields else "")
            link_fields = sheet_cfg.get("link_fields", [])

            # 用每个扩展关键词分别搜索
            for kw in keywords:
                try:
                    sheet_results = search_sheet(
                        file_id=file_id,
                        file_name=file_name,
                        sheet_name=sname,
                        keyword=kw,
                        search_fields=search_fields,
                        name_field=name_field,
                        link_fields=link_fields,
                        top_k=top_k_per_source,
                    )
                    # 追加去重（避免同一关键词搜出重复记录）
                    for r in sheet_results:
                        r["matched_keyword"] = kw  # 记录是哪个词匹配到的
                        key = (r["file_id"], r["record_id"])
                        if key not in seen:
                            seen.add(key)
                            all_results.append(r)
                except Exception as e:
                    log.warning(f"搜索 [{file_name} > {sname}]（关键词={kw}）出错: {e}")

    # 按匹配度降序
    all_results.sort(key=lambda x: x["match_score"], reverse=True)
    return all_results
