#!/usr/bin/env python3
"""
处理用户直接发送的 URL 灵感记录：
1. web_fetch 抓取页面内容
2. LLM 生成摘要和标签
3. WPS 云文档创建智能文档
4. 写入灵感记录表
"""
import re
import sys
import os
import json
import subprocess

# kingwork 根目录（动态获取，兼容 Mac/Linux/Windows）
from pathlib import Path
KINGWORK_ROOT = str(Path(__file__).resolve().parent.parent.parent)
sys.path.insert(0, KINGWORK_ROOT)

from kingwork_client.base import get_wps365_root as _get_wps365_root
WPS365_ROOT = str(_get_wps365_root())
if WPS365_ROOT not in sys.path:
    sys.path.insert(0, WPS365_ROOT)
os.chdir(KINGWORK_ROOT)

import kingwork_client.llm as kingwork_llm
import kingwork_client.tables as kingwork_tables
from wpsv7client import create_otl_document, get_drive_id
from wpsv7client.airpage import write_airpage_content
LLMClient = kingwork_llm.LLMClient


def extract_urls(text: str) -> list:
    """从文本中提取所有 URL。"""
    url_pattern = re.compile(
        r'https?://[^\s<>"{}|\\^`\[\]]+',
        re.IGNORECASE
    )
    return list(dict.fromkeys(url_pattern.findall(text)))


def summarize_url_content(llm, url: str, content: str) -> dict:
    """用大模型从 URL 内容中提取摘要和标签。"""
    prompt = (
        f"你是一个信息分析助手。请根据以下从网页提取的内容，生成简洁的摘要和标签。\n\n"
        f"来源URL：{url}\n\n"
        f"内容：\n{content[:4000]}\n\n"
        f"请返回 JSON 格式：\n"
        f'{{"summary": "核心内容摘要，100字以内", "tags": ["标签1", "标签2", "标签3", "标签4"]}}\n'
        f"标签从以下选项中选择最相关的4个：产品创意,销售策略,流程优化,技术方案,AI应用,客户需求,行业洞察,竞品信息,合作机会\n"
        f"如果内容与这些选项都不相关，可以自己生成相关标签但不超过4个。"
    )
    result = llm._call(prompt) or {}
    raw = ""
    if isinstance(result, dict):
        raw = result.get("raw", "")
    else:
        raw = str(result)
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            pass
    return {"summary": content[:150], "tags": []}


def process_url(url: str, content_text: str = "", title: str = "") -> dict:
    """
    处理单个 URL：抓取内容 → 生成摘要 → 创建 WPS 文档 → 写入灵感记录表。
    返回处理结果 dict。
    """
    from datetime import datetime

    llm = LLMClient()
    tables = kingwork_tables.KingWorkTables()
    wps_skill_path = WPS365_ROOT

    file_id = os.environ.get("KINGWORK_FILE_ID", "cbMwPNjcGRwD")
    sheet_id_num = tables.sheet_ids.get("idea_records", 10)

    today = datetime.now().strftime("%Y/%m/%d")

    results = {
        "url": url,
        "title": title,
        "summary": "",
        "tags": [],
        "doc_url": "",
        "doc_link_id": "",
        "idea_record_id": "",
        "success": False,
        "error": "",
    }

    # 1. 抓取内容（如果没传入）
    if not content_text:
        # 尝试通过 subprocess 调用 web_fetch
        import urllib.request
        headers = {"User-Agent": "Mozilla/5.0 (compatible; KingWorkBot/1.0)"}
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=15) as resp:
                html = resp.read().decode("utf-8", errors="ignore")
            # 去除 JS/CSS
            html = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.I|re.S)
            html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.I|re.S)
            text = re.sub(r'<[^>]+>', ' ', html)
            text = re.sub(r'\s+', ' ', text).strip()
            content_text = text[:5000]
            if not title:
                tm = re.search(r'<title[^>]*>([^<]+)</title>', html, re.I)
                title = tm.group(1).strip() if tm else ""
        except Exception as e:
            results["error"] = f"内容获取失败：{e}"
            return results

    # 2. LLM 生成摘要和标签
    summary_data = summarize_url_content(llm, url, content_text)
    summary = summary_data.get("summary", content_text[:150])
    tags = (summary_data.get("tags") or [])[:4]

    results["summary"] = summary
    results["tags"] = tags
    if title:
        results["title"] = title

    # 3. 创建 WPS 云文档（完整路径 myDocs/灵感收藏）
    drive_id = get_drive_id("private")
    doc_title = (title[:80] if title else f"灵感收藏 {url[:30]}").strip()
    doc_url = ""
    doc_link_id = ""

    for attempt in range(3):
        try:
            doc_resp = create_otl_document(
                drive_id=drive_id,
                file_name=doc_title,
                parent_path=["我的文档", "灵感收藏"],
                on_name_conflict="rename",
            )
            doc_data = doc_resp.get("data") or {}
            doc_url = doc_data.get("link_url") or ""
            doc_link_id = doc_data.get("link_id") or ""
            if doc_url:
                break  # 成功获取 link_url，跳出重试
        except Exception as e:
            if attempt < 2:
                import time
                time.sleep(1)  # 等待1秒后重试
                continue
            results["error"] = f"WPS文档创建失败：{e}"
    else:
        if not results["error"]:
            results["error"] = "WPS文档创建失败（重试后仍无 link_url）"

    results["doc_url"] = doc_url
    results["doc_link_id"] = doc_link_id

    if not doc_url:
        results["error"] = (results["error"] + "; WPS文档创建无链接" if results["error"] else "WPS文档创建无链接")

    # 4. 向智能文档写入内容
    if doc_link_id:
        try:
            body_text = content_text[:4000]
            doc_content_md = (
                f"# {title or doc_title}\n\n"
                f"来源：[{url}]({url})\n\n"
                f"## 摘要\n{summary}\n\n"
                f"## 正文\n{body_text}\n\n"
                f"## 标签\n{', '.join(tags)}\n"
            )
            write_airpage_content(doc_link_id, title or doc_title, doc_content_md, pos="begin")
        except Exception as e:
            results["error"] = (results["error"] + f"; 文档写入失败：{e}") if results["error"] else f"文档写入失败：{e}"

    # 5. 写入灵感记录表
    record = {
        "记录时间": today,
        "灵感内容": summary,
        "灵感类别": "其他",
        "来源": "AI自动分析",
        "URL链接地址": url,
        "WPS文档地址": doc_url,
        "标签": tags,
        "创建时间": today,
    }

    try:
        cmd = [
            sys.executable, str(Path(wps_skill_path) / "skills" / "dbsheet" / "run.py"),
            "create-records", file_id, str(sheet_id_num),
            "--json", json.dumps([record], ensure_ascii=False)
        ]
        out = subprocess.check_output(cmd, timeout=15)
        # 解析返回的记录ID
        try:
            out_str = out.decode("utf-8", errors="ignore")
            id_m = re.search(r'"id"\s*:\s*(\d+)', out_str)
            if id_m:
                results["idea_record_id"] = id_m.group(1)
        except Exception:
            pass
    except Exception as e:
        results["error"] = (results["error"] + f"; 记录写入失败：{e}") if results["error"] else f"记录写入失败：{e}"

    results["success"] = not results["error"]
    return results


if __name__ == "__main__":
    # 支持命令行调用：python url_idea.py <url>
    if len(sys.argv) < 2:
        print("用法: python url_idea.py <url>")
        sys.exit(1)

    url = sys.argv[1]
    result = process_url(url)
    print(json.dumps(result, ensure_ascii=False, indent=2))
