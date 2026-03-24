#!/usr/bin/env python3
"""
kingbrowse CLI 入口
用法：
    python run.py <关键词> [--top N]
    python run.py 端云一体 --top 5
"""
import sys
import os
import argparse

BROWSE_DIR = os.path.dirname(os.path.abspath(__file__))
# kingbrowse 目录在 .../kingwork/skills/kingbrowse/
# run.py 在 .../kingwork/skills/kingbrowse/run.py
# 需要把 skills/ 加入 path，这样 "import kingbrowse" 才能找到 kingbrowse/ 目录
SKILLS_DIR = BROWSE_DIR
KINGWORK_ROOT = os.path.dirname(os.path.dirname(SKILLS_DIR))
sys.path.insert(0, os.path.dirname(SKILLS_DIR))  # .../kingwork/skills/ -> 找到 kingbrowse/ 目录
sys.path.insert(0, KINGWORK_ROOT)
sys.path.insert(0, "/root/.openclaw/skills/wps365-skill")
os.chdir(KINGWORK_ROOT)

from kingbrowse.browse import search


def format_result(r: dict, idx: int) -> str:
    """格式化单条结果为易读文本。"""
    link = r.get("link") or ""
    name = r.get("name") or "（无标题）"
    snippet = r.get("snippet") or ""
    matched_kw = r.get("matched_keyword", "")
    snippet_display = f"\n   📝 {snippet}" if snippet else ""

    if link:
        link_markdown = f"[🔗 {link}]({link})"
    else:
        link_markdown = "（无链接）"

    kw_tag = f"（命中词：`{matched_kw}`）" if matched_kw else ""

    return (
        f"{idx}. **{name}**\n"
        f"   📋 来源：{r['file_name']} > {r['sheet_name']}\n"
        f"   {link_markdown}\n"
        f"   🏷️  匹配：{r['match_field']}（{r['relevance']}）{kw_tag}{snippet_display}"
    )


def main():
    parser = argparse.ArgumentParser(description="kingbrowse - 资料检索")
    parser.add_argument("keyword", nargs="+", help="搜索关键词")
    parser.add_argument("--top", type=int, default=10, help="每表最多返回条数（默认10）")
    parser.add_argument("--json", action="store_true", help="输出 JSON 格式")

    args = parser.parse_args()
    keyword = " ".join(args.keyword).strip()

    if not keyword:
        print("关键词不能为空", file=sys.stderr)
        sys.exit(1)

    print(f"\n🔍 正在搜索：「{keyword}」...\n", end="", flush=True)

    results = search(keyword, top_k_per_source=args.top)

    if not results:
        print("未找到相关资料，试试其他关键词？")
        sys.exit(0)

    print(f"共找到 {len(results)} 条相关资料：\n")

    if args.json:
        import json
        print(json.dumps(results, ensure_ascii=False, indent=2))
        return

    for i, r in enumerate(results, 1):
        print(format_result(r, i))
        print()


if __name__ == "__main__":
    main()
