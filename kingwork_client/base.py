# -*- coding: utf-8 -*-
"""
KingWork 基础配置和工具函数。
"""
import os
import sys
import json
import yaml
import logging
import logging.handlers
from pathlib import Path
from datetime import datetime, timezone

# 全局日志初始化：同时输出到控制台和文件
def init_kingwork_logging():
    logger = logging.getLogger("kingwork")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    
    # 日志格式
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    
    # 1. 控制台输出Handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 2. 文件输出Handler，自动轮转（100MB/文件，最多保留5个备份）
    log_path = "/var/log/kingwork_debug.log"
    try:
        # 尝试写/var/log目录
        file_handler = logging.handlers.RotatingFileHandler(
            log_path, maxBytes=100*1024*1024, backupCount=5, encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except PermissionError:
        # 无权限则写入用户目录
        user_log_dir = Path.home() / ".openclaw" / "skills" / "kingwork" / "logs"
        user_log_dir.mkdir(parents=True, exist_ok=True)
        log_path = str(user_log_dir / "kingwork_debug.log")
        file_handler = logging.handlers.RotatingFileHandler(
            log_path, maxBytes=100*1024*1024, backupCount=5, encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    # 重定向print和stderr到日志
    class LoggerWriter:
        def __init__(self, level):
            self.level = level
            self.buffer = []
        def write(self, message):
            if message.strip():
                self.buffer.append(message)
                if message.endswith('\n'):
                    full_msg = ''.join(self.buffer).rstrip('\n')
                    logger.log(self.level, full_msg)
                    self.buffer = []
        def flush(self):
            if self.buffer:
                full_msg = ''.join(self.buffer).rstrip('\n')
                logger.log(self.level, full_msg)
                self.buffer = []
    
    sys.stdout = LoggerWriter(logging.INFO)
    sys.stderr = LoggerWriter(logging.ERROR)
    
    return logger

# 自动初始化日志
logger = init_kingwork_logging()

# 全局调试日志配置
_cfg = None
def _load_config_once():
    global _cfg
    if _cfg is None:
        config_path = KINGWORK_ROOT / "config" / "kingwork.yaml"
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                _cfg = yaml.safe_load(f) or {}
        else:
            _cfg = {}
    return _cfg

# 初始化日志
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("/var/log/kingwork_debug.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("kingwork")

def debug_log(msg: str):
    """输出调试日志，同时写入控制台和/var/log/kingwork_debug.log"""
    cfg = _load_config_once()
    if cfg.get("debug", {}).get("enable_debug_log", False):
        logger.debug(msg)


# KingWork 根目录
KINGWORK_ROOT = Path(__file__).resolve().parent.parent

# WPS365 Skill 根目录（从配置或环境变量获取）
def get_wps365_root() -> Path:
    cfg = get_config()
    env_path = os.environ.get("WPS365_SKILL_PATH")
    if env_path:
        return Path(env_path)
    skill_path = cfg.get("wps365_skill_path", "")
    if skill_path and not skill_path.startswith("${"):
        return Path(skill_path)
    # 默认：相邻目录
    return KINGWORK_ROOT.parent / "wps365-skill"


def get_config() -> dict:
    """加载 kingwork.yaml 配置。"""
    config_path = KINGWORK_ROOT / "config" / "kingwork.yaml"
    if not config_path.exists():
        return {}
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def get_file_id() -> str:
    """获取多维表文件 ID（优先环境变量）。"""
    env_id = os.environ.get("KINGWORK_FILE_ID")
    if env_id:
        return env_id
    cfg = get_config()
    file_id = cfg.get("file_id", "")
    if file_id and not file_id.startswith("${"):
        return file_id
    raise ValueError(
        "未配置多维表文件 ID。\n"
        "请设置环境变量 KINGWORK_FILE_ID，\n"
        "或运行 python scripts/init_tables.py 初始化。"
    )


def get_sheet_ids() -> dict:
    """获取数据表 ID 映射。"""
    cfg = get_config()
    return cfg.get("sheet_ids", {})


def get_analysis_config() -> dict:
    cfg = get_config()
    return cfg.get("analysis", {
        "similarity_threshold": 0.7,
        "enable_surprise_extraction": True,
        "batch_size": 10,
    })


def get_llm_config() -> dict:
    """获取 LLM 配置（从 kingwork.yaml 的 llm 节）。"""
    cfg = get_config()
    return cfg.get("llm", {
        "endpoint": "http://localhost:8080/v1/chat/completions",
        "model": "doubao-seed-2.0-pro",
        "temperature": 0.1,
        "api_key": "",
    })


def get_alert_config() -> dict:
    cfg = get_config()
    return cfg.get("alert", {
        "inactive_customer_days": 15,
        "overdue_warning_days": 0,
    })


def save_config_sheet_ids(sheet_id_map: dict):
    """将初始化后的 sheet_id 映射写回配置文件。"""
    config_path = KINGWORK_ROOT / "config" / "kingwork.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    cfg["sheet_ids"] = sheet_id_map
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(cfg, f, allow_unicode=True, default_flow_style=False)


def save_config_file_id(file_id: str):
    """将 file_id 写回配置文件。"""
    config_path = KINGWORK_ROOT / "config" / "kingwork.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    cfg["file_id"] = file_id
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(cfg, f, allow_unicode=True, default_flow_style=False)


def load_prompts() -> dict:
    """加载提示词模板。"""
    prompts_path = KINGWORK_ROOT / "config" / "prompts.yaml"
    if not prompts_path.exists():
        return {}
    with open(prompts_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return data.get("prompts", {})


def load_tables_schema() -> list:
    """加载数据表 Schema。"""
    tables_path = KINGWORK_ROOT / "config" / "tables.yaml"
    if not tables_path.exists():
        return []
    with open(tables_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return data.get("tables", [])

def get_enum_config() -> dict:
    """加载枚举字段配置。"""
    enum_config_path = KINGWORK_ROOT / "config" / "fields_enum.yaml"
    if not enum_config_path.exists():
        return {}
    with open(enum_config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def now_iso() -> str:
    """返回当前时间 ISO 格式（东8区）。"""
    from datetime import timedelta
    tz_cst = timezone(timedelta(hours=8))
    return datetime.now(tz=tz_cst).isoformat()


def today_str() -> str:
    """返回今日日期字符串 YYYY/MM/DD（和WPS多维表格式一致）。"""
    from datetime import timedelta
    tz_cst = timezone(timedelta(hours=8))
    return datetime.now(tz=tz_cst).strftime("%Y/%m/%d")


def weekday_cn() -> str:
    """返回今天是周几（中文）。"""
    from datetime import timedelta
    tz_cst = timezone(timedelta(hours=8))
    days = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    return days[datetime.now(tz=tz_cst).weekday()]


# 导入缓存
_wps_client_cache = None
_import_mode = None  # "direct" | "path" | None (unknown)


def get_import_mode() -> str:
    """检测并返回当前使用的导入模式（支持配置覆盖）。

    优先级：
    1. 配置指定（import_mode: "direct" 或 "path"）
    2. 自动检测（直接导入成功则用 direct，否则用 path）
    """
    global _import_mode
    cfg = get_config()
    cfg_mode = cfg.get("import_mode", "").strip().lower()

    # 优先使用配置指定的模式（排除 "auto"）
    if cfg_mode in ("direct", "path"):
        _import_mode = cfg_mode
        return _import_mode

    # 自动检测（如果还未检测过）
    if _import_mode is None:
        if try_import_wpsv7client():
            _import_mode = "direct"
        else:
            _import_mode = "path"

    return _import_mode


def get_skill_call_mode() -> str:
    """获取子技能调用模式。"""
    cfg = get_config()
    mode = cfg.get("skill_call_mode", "subprocess").strip().lower()
    if mode not in ("subprocess", "direct"):
        mode = "subprocess"
    return mode


def try_import_wpsv7client() -> bool:
    """尝试直接导入 wpsv7client（假设已作为包安装）。"""
    try:
        import wpsv7client
        return True
    except ImportError:
        return False


def add_wps365_to_path():
    """将 wps365-skill 根目录加入 sys.path。"""
    wps365_root = get_wps365_root()
    if str(wps365_root) not in sys.path:
        sys.path.insert(0, str(wps365_root))


def import_wpsv7client():
    """导入 wpsv7client，优先直接导入，失败则回退到路径导入。"""
    global _wps_client_cache

    # 优先尝试直接导入
    try:
        import wpsv7client
        return wpsv7client
    except ImportError:
        pass

    # 回退到路径导入
    add_wps365_to_path()
    import wpsv7client
    return wpsv7client


# 复用 wps365 的 WpsV7Client
def get_wps_client():
    """获取 WpsV7Client 实例（带缓存）。"""
    global _wps_client_cache

    if _wps_client_cache is not None:
        return _wps_client_cache

    wpsv7client = import_wpsv7client()
    _wps_client_cache = wpsv7client.WpsV7Client()
    return _wps_client_cache


def reset_wps_client_cache():
    """重置 WPS 客户端缓存（主要用于测试）。"""
    global _wps_client_cache, _import_mode
    _wps_client_cache = None
    _import_mode = None


def get_wps365_functions():
    """获取 wps365 的 dbsheet 函数（带缓存）。"""
    wpsv7client = import_wpsv7client()
    return {
        "dbsheet_get_schema": wpsv7client.dbsheet_get_schema,
        "dbsheet_list_records": wpsv7client.dbsheet_list_records,
        "dbsheet_batch_create_records": wpsv7client.dbsheet_batch_create_records,
        "dbsheet_batch_update_records": wpsv7client.dbsheet_batch_update_records,
        "dbsheet_batch_delete_records": wpsv7client.dbsheet_batch_delete_records,
        "dbsheet_create_sheet": wpsv7client.dbsheet_create_sheet,
        "dbsheet_create_view": wpsv7client.dbsheet_create_view,
    }


class KingWorkConfig:
    """KingWork 配置容器，供各 skill 使用。"""

    def __init__(self):
        self._cfg = get_config()
        self.file_id = get_file_id()
        self.sheet_ids = get_sheet_ids()
        self.analysis = get_analysis_config()
        self.alert = get_alert_config()
        self.prompts = load_prompts()
        self.wps365_root = get_wps365_root()
        # 多维表访问链接
        self.dbt_link = "https://www.kdocs.cn/l/cbMwPNjcGRwD"


# 数据表名称映射（key -> 中文名称）
SHEET_NAME_MAP = {
    "diary_records": "01日记记录",
    "todo_records": "02待办记录",
    "customer_profiles": "03客户档案",
    "project_profiles": "04项目档案",
    "customer_followups": "05客户跟进记录",
    "learning_records": "06学习成长记录",
    "support_records": "07横向支持记录",
    "team_records": "08团队事务记录",
    "idea_records": "09灵感记录",
    "surprise_docs": "10惊喜文档记录",
    "surprise_communications": "11惊喜沟通记录"
}


def print_exec_summary(updated_tables: list = None):
    """
    统一输出执行总结，包含更新的数据表和多维表访问链接。
    :param updated_tables: 更新的数据表key列表，如 ["diary_records", "todo_records"]
    """
    cfg = KingWorkConfig()
    print("\n" + "="*60)
    print("✅ 操作完成")
    if updated_tables and len(updated_tables) > 0:
        print("\n📝 更新的数据表：")
        for table_key in updated_tables:
            table_name = SHEET_NAME_MAP.get(table_key, table_key)
            print(f"  - {table_name}")
    print(f"\n🔗 多维表访问地址：<{cfg.dbt_link}>")
    print("="*60 + "\n")
