# -*- coding: utf-8 -*-
"""KingWork 客户端库"""
from .base import (
    KingWorkConfig,
    get_config,
    get_import_mode,
    get_skill_call_mode,
    try_import_wpsv7client,
    get_wps_client,
    import_wpsv7client,
    add_wps365_to_path,
    get_wps365_root,
    reset_wps_client_cache,
)
from .tables import KingWorkTables
from .llm import LLMClient

__all__ = [
    "KingWorkConfig",
    "get_config",
    "get_import_mode",
    "get_skill_call_mode",
    "try_import_wpsv7client",
    "get_wps_client",
    "import_wpsv7client",
    "add_wps365_to_path",
    "get_wps365_root",
    "reset_wps_client_cache",
    "KingWorkTables",
    "LLMClient",
]
