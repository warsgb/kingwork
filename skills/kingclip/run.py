#!/usr/bin/env python3
"""
kingclip - 灵感收藏技能
处理用户发送的「灵感 + URL」，自动抓取内容、生成摘要和标签、
在 WPS 云文档归档、写入灵感记录表。

用法：
    python run.py process <url>
"""
from __future__ import annotations
import re
import sys
import os
import json
import time
import subprocess
import argparse
from datetime import datetime

# ------------------------------------------------------------------
# 路径配置（动态获取，兼容不同部署环境）
# ------------------------------------------------------------------
KINGWORK_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WPS365_ROOT = os.environ.get(
    "WPS365_SKILL_PATH",
    os.path.abspath(os.path.join(KINGWORK_ROOT, "..", "..", "official", "wps365-skill"))
)
IDEA_SHEET_ID = "10"
# 灵感收藏文件夹（我的文档/灵感收藏）parent_id
INSPIRATION_FOLDER_ID = "fq26KPJnmxMSVZhxtHc6rxGMzxMnUjsLf"

sys.path.insert(0, KINGWORK_ROOT)
sys.path.insert(0, WPS365_ROOT)
os.chdir(KINGWORK_ROOT)

# 从配置文件动态读取 file_id
from kingwork_client.base import get_file_id as _get_file_id
KINGWORK_FILE_ID = _get_file_id()


# ------------------------------------------------------------------
# 工具函数
# ------------------------------------------------------------------

def extract_urls(text: str) -> list:
    """从文本中提取所有 URL。"""
    pattern = re.compile(
        r'https?://[^\s<>"{}|\\^`\[\]]+',
        re.IGNORECASE
    )
    return list(dict.fromkeys(pattern.findall(text)))


def extract_text_from_html(html: str) -> tuple[str, str]:
    """
    从 HTML 中提取正文文本和标题。
    返回 (title, body_text)。
    """
    # 去除 script / style
    html = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.I | re.S)
    html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.I | re.S)

    # 标题
    tm = re.search(r'<title[^>]*>([^<]+)</title>', html, re.I)
    title = tm.group(1).strip() if tm else ""

    # HTML → 文本
    text = re.sub(r'<[^>]+>', ' ', html)
    text = re.sub(r'\s+', ' ', text).strip()
    return title, text


def fetch_url(url: str) -> tuple[str, str]:
    """
    通过 urllib 获取 URL 内容。
    返回 (title, body_text)，title 从 <title> 标签提取，body 为去标签后的正文。
    """
    try:
        from urllib.request import Request, urlopen
        headers_req = {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,*/*",
            "Accept-Language": "zh-CN,zh;q=0.9",
        }
        req = Request(url, headers=headers_req)
        with urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        # 提取 title
        title_m = re.search(r'<title[^>]*>([^<]+)</title>', html, re.I)
        title = title_m.group(1).split("_")[0].split("|")[0].strip() if title_m else ""
        # 去除 script/style
        html = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.I | re.S)
        html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.I | re.S)
        text = re.sub(r'<[^>]+>', ' ', html)
        text = re.sub(r'\s+', ' ', text).strip()
        return title, text
    except Exception as e:
        return "", str(e)


def summarize_content(llm, url: str, content: str) -> dict:
    """
    用大模型从页面内容中提取摘要和标签。
    """
    prompt = (
        f"你是一个信息分析助手。请根据以下从网页提取的内容，生成简洁的摘要和标签。\n\n"
        f"来源URL：{url}\n\n"
        f"内容：\n{content[:4000]}\n\n"
        f"请返回 JSON 格式：\n"
        f'{{"summary": "核心内容摘要，100字以内", "tags": ["标签1", "标签2", "标签3", "标签4"]}}\n'
        f"标签从以下选项中选择最相关的4个（如果不相关可自己生成）：\n"
        f"产品创意, 销售策略, 流程优化, 技术方案, AI应用, 客户需求, 行业洞察, 竞品信息, 合作机会\n"
    )
    result = llm._call(prompt) or {}
    if isinstance(result, dict):
        # LLM 返回 {summary, tags} 直接可用（_call 已解析 JSON）
        if "summary" in result or "tags" in result:
            return {
                "summary": result.get("summary") or content[:150],
                "tags": result.get("tags") or [],
            }
        # 尝试从 raw 字段解析
        raw = result.get("raw", "")
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if m:
            try:
                parsed = json.loads(m.group(0))
                return {
                    "summary": parsed.get("summary") or content[:150],
                    "tags": parsed.get("tags") or [],
                }
            except Exception:
                pass
    return {"summary": content[:150], "tags": []}


# ------------------------------------------------------------------
# WPS 文档标签
# ------------------------------------------------------------------
# 缓存：标签名 → WPS label_id
_WPS_LABELS_CACHE: dict[str, str] = {}
_WPS_LABELS_INITED = False


def _init_wps_labels_cache():
    """从 WPS 拉取现有标签列表填充缓存。"""
    global _WPS_LABELS_CACHE, _WPS_LABELS_INITED
    if _WPS_LABELS_INITED:
        return
    try:
        from wpsv7client import list_drive_labels
        resp = list_drive_labels(page_size=500)
        for item in resp.get("data", {}).get("items", []):
            name = item.get("name", "").strip()
            lid = item.get("id", "")
            if name and lid:
                _WPS_LABELS_CACHE[name] = lid
    except Exception:
        pass
    _WPS_LABELS_INITED = True


def _get_or_create_wps_label(tag_name: str) -> str | None:
    """
    根据标签名称获取或创建 WPS 文档标签，返回 label_id。
    """
    _init_wps_labels_cache()
    if tag_name in _WPS_LABELS_CACHE:
        return _WPS_LABELS_CACHE[tag_name]

    try:
        from wpsv7client import create_drive_label
        resp = create_drive_label(name=tag_name.strip())
        if resp.get("code") == 0:
            label_id = resp.get("data", {}).get("id", "")
            if label_id:
                _WPS_LABELS_CACHE[tag_name] = label_id
                return label_id
    except Exception:
        pass
    return None


def _apply_wps_doc_labels(doc_link_id: str, tags: list[str]) -> list[str]:
    """
    将标签列表应用到 WPS 云文档，返回成功打上的标签名称列表。
    """
    if not doc_link_id or not tags:
        return []
    applied = []
    for tag in tags:
        label_id = _get_or_create_wps_label(tag)
        if not label_id:
            continue
        try:
            from wpsv7client import batch_add_drive_label_objects
            resp = batch_add_drive_label_objects(label_id, [doc_link_id])
            if resp.get("code") == 0:
                applied.append(tag)
        except Exception:
            pass
    return applied


# ------------------------------------------------------------------
# 核心处理
# ------------------------------------------------------------------
# 核心处理
# ------------------------------------------------------------------

def process_url(url: str) -> dict:
    """
    处理单个 URL：
    1. 抓取内容
    2. LLM 生成摘要和标签
    3. WPS 云文档创建并写入内容
    4. 写入灵感记录表

    返回结果 dict。
    """
    from wpsv7client import create_otl_document, get_drive_id
    from wpsv7client.airpage import write_airpage_content

    results = {
        "url": url,
        "title": "",
        "summary": "",
        "tags": [],
        "doc_url": "",
        "doc_link_id": "",
        "success": False,
        "error": "",
    }

    today = datetime.now().strftime("%Y/%m/%d")

    # ── 1. 抓取内容 ──────────────────────────────────────────────
    title, body_text = fetch_url(url)
    if not body_text:
        results["error"] = f"页面获取失败：{title or str(url)}"
        return results

    body_text = body_text[:5000]
    results["title"] = title

    # ── 2. LLM 摘要和标签 ─────────────────────────────────────────
    import kingwork_client.llm as kllm
    llm = kllm.LLMClient()
    sd = summarize_content(llm, url, body_text)
    results["summary"] = sd.get("summary", body_text[:150])
    results["tags"] = (sd.get("tags") or [])[:4]

    # ── 3. 创建 WPS 智能文档 ──────────────────────────────────────
    drive_id = get_drive_id("private")
    doc_title = (title[:80] if title else f"灵感 {url[8:38]}").strip()
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
                results["doc_url"] = doc_url
                results["doc_link_id"] = doc_link_id
                break
        except Exception as e:
            if attempt < 2:
                time.sleep(1)
                continue
            results["error"] = f"WPS文档创建失败：{e}"
    else:
        if not results["error"]:
            results["error"] = "WPS文档创建失败（重试后仍无 link_url）"

    # ── 4. 写入文档内容 ───────────────────────────────────────────
    if doc_link_id:
        doc_content_md = (
            f"# {title or doc_title}\n\n"
            f"来源：[{url}]({url})\n\n"
            f"## 摘要\n{results['summary']}\n\n"
            f"## 正文\n{body_text[:4000]}\n\n"
            f"## 标签\n{', '.join(results['tags'])}\n"
        )
        try:
            write_airpage_content(doc_link_id, title or doc_title, doc_content_md, pos="begin")
        except Exception as e:
            results["error"] = (results["error"] + f"; 文档写入失败：{e}"
                               if results["error"] else f"文档写入失败：{e}")

    # ── 4.5 打 WPS 文档标签 ─────────────────────────────────────
    if doc_link_id and results["tags"]:
        applied = _apply_wps_doc_labels(doc_link_id, results["tags"])
        if applied:
            results["tags_applied"] = applied

    # ── 5. 写入灵感记录表 ─────────────────────────────────────────
    # 标签：优先使用 WPS 打标签成功的，fallback 到 LLM 提取的
    final_tags = results.get("tags_applied") or results.get("tags") or []
    record = {
        "记录时间": today,
        "灵感内容": results["summary"],
        "灵感类别": "其他",
        "来源": "AI自动分析",
        "URL链接地址": url,
        "WPS文档地址": doc_url,
        "标签": final_tags,
        "创建时间": today,
    }

    try:
        cmd = [
            sys.executable, f"{WPS365_ROOT}/skills/dbsheet/run.py",
            "create-records", KINGWORK_FILE_ID, IDEA_SHEET_ID,
            "--json", json.dumps([record], ensure_ascii=False)
        ]
        subprocess.check_output(cmd, timeout=15)
    except subprocess.CalledProcessError as e:
        err = e.output.decode("utf-8", errors="ignore") if e.output else str(e)
        results["error"] = (results["error"] + f"; 记录写入失败：{err}"
                           if results["error"] else f"记录写入失败：{err}")

    results["success"] = not results["error"]
    return results


# ------------------------------------------------------------------
# CLI 入口
# ------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="kingclip - 灵感收藏")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("process", help="处理单个 URL 灵感")
    p.add_argument("url", help="要处理的 URL")
    p.add_argument("--title", default="", help="手动指定标题（可选）")
    p.add_argument("--content", default="", help="手动指定正文内容（可选，跳过自动抓取）")
    p.add_argument("--skip-fetch", action="store_true", help="跳过网页抓取（content已直接提供）")

    args = parser.parse_args()

    if args.cmd == "process":
        # 如果已预抓取内容，直接走 LLM → 写文档 → 写表
        if args.content:
            import kingwork_client.llm as kllm
            llm = kllm.LLMClient()
            sd = summarize_content(llm, args.url, args.content)
            summary = sd.get("summary", args.content[:150])
            tags = (sd.get("tags") or [])[:4]
            title = args.title or ""

            print(f"\n{'='*50}")
            print(f"🔗 URL：{args.url}")
            if title:
                print(f"📄 标题：{title}")
            print(f"💡 摘要：{summary}")
            print(f"🏷️  标签（多维表）：{applied_tags if applied_tags else tags}")
            print(f"🏷️  LLM提取标签：{tags}")

            # 把预抓取内容注入到 process_url（通过修改全局逻辑太耦合，直接在这里处理）
            # 直接调用关键步骤：创建文档 + 写入表
            from wpsv7client import create_otl_document, get_drive_id
            from wpsv7client.airpage import write_airpage_content
            from datetime import datetime

            today = datetime.now().strftime("%Y/%m/%d")
            drive_id = get_drive_id("private")
            doc_title = (title[:80] if title else f"灵感 {args.url[8:38]}").strip()
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
                        break
                except Exception:
                    if attempt < 2:
                        time.sleep(1)
                        continue

            if doc_link_id:
                try:
                    body = args.content[:4000]
                    write_airpage_content(
                        doc_link_id, title or doc_title,
                        f"# {title or doc_title}\n\n"
                        f"来源：[{args.url}]({args.url})\n\n"
                        f"## 摘要\n{summary}\n\n"
                        f"## 正文\n{body}\n\n"
                        f"## 标签\n{', '.join(tags)}\n",
                        pos="begin"
                    )
                except Exception as e:
                    print(f"⚠️  文档内容写入失败：{e}")

            # 打 WPS 文档标签
            applied_tags = []
            if doc_link_id and tags:
                applied_tags = _apply_wps_doc_labels(doc_link_id, tags)
                if applied_tags:
                    print(f" 🏷️ WPS文档标签：{applied_tags}")

            if doc_url:
                print(f"📂 WPS文档：{doc_url}")
            else:
                print(f"⚠️  WPS文档：创建失败")

            # 写入灵感记录表
            record = {
                "记录时间": today,
                "灵感内容": summary,
                "灵感类别": "其他",
                "来源": "AI自动分析",
                "URL链接地址": args.url,
                "WPS文档地址": doc_url,
                "标签": applied_tags if applied_tags else tags,
                "创建时间": today,
            }
            try:
                cmd = [
                    sys.executable, f"{WPS365_ROOT}/skills/dbsheet/run.py",
                    "create-records", KINGWORK_FILE_ID, IDEA_SHEET_ID,
                    "--json", json.dumps([record], ensure_ascii=False)
                ]
                subprocess.check_output(cmd, timeout=15)
                print(f"📊 状态：{'✅ 成功' if doc_url else '⚠️  部分成功'}")
            except Exception as e:
                print(f"📊 状态：❌ 失败 - 记录写入：{e}")
            return

        result = process_url(args.url)

        print(f"\n{'='*50}")
        print(f"🔗 URL：{result['url']}")
        if result["title"]:
            print(f"📄 标题：{result['title']}")
        print(f"💡 摘要：{result['summary'] or '(获取失败)'}")
        applied = result.get("tags_applied", [])
        llm_tags = result.get("tags", [])
        print(f"🏷️  LLM标签：{llm_tags or '(无)'}")
        if applied:
            print(f"🏷️  WPS文档标签：{applied}")
        if result["doc_url"]:
            print(f"📂 WPS文档：{result['doc_url']}")
        else:
            print(f"⚠️  WPS文档：创建失败")
        print(f"📊 状态：{'✅ 成功' if result['success'] else '❌ 失败'}")
        if result["error"]:
            print(f"💬 备注：{result['error']}")

        sys.exit(0 if result["success"] else 1)


if __name__ == "__main__":
    main()
