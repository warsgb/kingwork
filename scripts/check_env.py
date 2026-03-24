#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
环境检测脚本 - 检查 KingWork 与 wps365-skill 的集成状态。

用法:
  python scripts/check_env.py
"""
import os
import sys
from pathlib import Path

# 添加项目根目录到路径
_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from kingwork_client.base import (
    get_import_mode,
    get_skill_call_mode,
    get_wps365_root,
    try_import_wpsv7client,
    get_config,
)


def check_wps_sid():
    """检查 WPS_SID 环境变量。"""
    sid = os.environ.get("WPS_SID")
    if sid:
        return True, f"WPS_SID: {sid[:10]}..." if len(sid) > 10 else f"WPS_SID: {sid}"
    return False, "WPS_SID 未设置"


def check_kingwork_file_id():
    """检查 KINGWORK_FILE_ID 环境变量。"""
    file_id = os.environ.get("KINGWORK_FILE_ID")
    if file_id:
        return True, f"KINGWORK_FILE_ID: {file_id[:20]}..."
    return False, "KINGWORK_FILE_ID 未设置"


def check_wps365_skill_path():
    """检查 WPS365_SKILL_PATH 环境变量。"""
    path = os.environ.get("WPS365_SKILL_PATH")
    if path:
        path_obj = Path(path)
        if path_obj.exists():
            return True, f"WPS365_SKILL_PATH: {path} (存在)"
        else:
            return False, f"WPS365_SKILL_PATH: {path} (目录不存在)"
    return False, "WPS365_SKILL_PATH 未设置"


def check_direct_import():
    """检查 wpsv7client 是否可以直接导入。"""
    available = try_import_wpsv7client()
    if available:
        return True, "wpsv7client 可以直接导入"
    return False, "wpsv7client 无法直接导入"


def check_wps365_root():
    """检查 wps365-skill 根目录。"""
    root = get_wps365_root()
    exists = root.exists()
    skills_dir = root / "skills"
    skills_exists = skills_dir.exists() if exists else False

    if exists and skills_exists:
        return True, f"wps365-skill 根目录: {root}"
    elif exists:
        return False, f"wps365-skill 根目录存在但无 skills 子目录: {root}"
    else:
        return False, f"wps365-skill 根目录不存在: {root}"


def check_config():
    """检查配置文件。"""
    config = get_config()
    import_mode = config.get("import_mode", "auto")
    skill_call_mode = config.get("skill_call_mode", "subprocess")
    wps365_path = config.get("wps365_skill_path", "")

    return {
        "import_mode": import_mode,
        "skill_call_mode": skill_call_mode,
        "wps365_skill_path": wps365_path or "(未配置)",
    }


def main():
    print("\n" + "=" * 60)
    print("KingWork 环境检测")
    print("=" * 60 + "\n")

    # 检查环境变量
    print("### 环境变量检查\n")

    checks = [
        ("WPS_SID", check_wps_sid),
        ("KINGWORK_FILE_ID", check_kingwork_file_id),
        ("WPS365_SKILL_PATH", check_wps365_skill_path),
    ]

    for name, check_func in checks:
        ok, msg = check_func()
        status = "✅" if ok else "❌"
        print(f"  {status} {name}: {msg}")

    # 检查导入能力
    print("\n### WPS365 导入检查\n")

    direct_ok, direct_msg = check_direct_import()
    print(f"  {'✅' if direct_ok else '⚠️'} {direct_msg}")

    root_ok, root_msg = check_wps365_root()
    print(f"  {'✅' if root_ok else '❌'} {root_msg}")

    # 检查导入模式
    print("\n### 导入模式\n")

    import_mode = get_import_mode()
    skill_call_mode = get_skill_call_mode()

    print(f"  当前导入模式: {import_mode}")
    print(f"  子技能调用模式: {skill_call_mode}")

    # 检查配置
    print("\n### 配置文件\n")

    config = check_config()
    for key, value in config.items():
        print(f"  {key}: {value}")

    # 总结
    print("\n" + "=" * 60)

    issues = []
    if not os.environ.get("WPS_SID"):
        issues.append("WPS_SID 未设置")
    if not os.environ.get("KINGWORK_FILE_ID"):
        issues.append("KINGWORK_FILE_ID 未设置")
    if not direct_ok and not root_ok:
        issues.append("wps365-skill 不可用")

    if issues:
        print("⚠️  发现问题:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("✅ 环境配置正常")

    print("=" * 60 + "\n")

    return 0 if not issues else 1


if __name__ == "__main__":
    sys.exit(main())
