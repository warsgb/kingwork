# -*- coding: utf-8 -*-
"""
KingWork LLM 客户端封装
支持 MiniMax API，兼容 OpenAI 格式
"""
import json
import time
import requests
import re
from datetime import datetime
from pathlib import Path

KINGWORK_ROOT = Path(__file__).resolve().parent.parent


class KingWorkBase(requests.Session):
    def __init__(self):
        super().__init__()
        self.config = self._load_config()

    def _load_config(self) -> dict:
        cfg_file = KINGWORK_ROOT / "config" / "kingwork.yaml"
        if not cfg_file.exists():
            return {}
        import yaml
        with open(cfg_file, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    def _load_prompts(self) -> dict:
        prompt_file = KINGWORK_ROOT / "config" / "prompts.yaml"
        if not prompt_file.exists():
            return {}
        import yaml
        with open(prompt_file, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
            return data.get("prompts", {})

    def save_prompts(self):
        prompt_file = KINGWORK_ROOT / "config" / "prompts.yaml"
        import yaml
        with open(prompt_file, encoding="utf-8") as f:
            yaml.dump({"prompts": self.prompts}, f, allow_unicode=True)


class KingWorkLLMClient(KingWorkBase):
    def __init__(self):
        super().__init__()
        self.prompts = self._load_prompts()
        llm_cfg = self.config.get("llm", {})
        self.endpoint = llm_cfg.get("endpoint", "https://api.minimaxi.com/v1/chat/completions")
        self.api_key = llm_cfg.get("api_key", "")
        self.model = llm_cfg.get("model", "MiniMax-M2.7")
        self.temperature = float(llm_cfg.get("temperature", 0.1))
        self.max_tokens = int(llm_cfg.get("max_tokens", 2000))
        self.timeout = int(llm_cfg.get("timeout", 30))
        self.max_retries = int(llm_cfg.get("max_retries", 3))

    def get_prompt(self, name: str) -> str:
        return self.prompts.get(name, "")

    def update_system_prompt(self, name: str, prompt: str):
        self.prompts[name] = prompt
        self.save_prompts()

    def generate(self, user_message: str, system_prompt: str = "") -> str:
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": user_message})
        return self._call(messages)

    def generate_report(self, period_type: str, period_str: str, work_data_str: str) -> dict:
        """生成报告，返回 dict（兼容 generate_report_prompt 的 llm_req.get("raw") 逻辑）。"""
        title_map = {"daily": "日报", "weekly": "周报", "monthly": "月报"}
        title = title_map.get(period_type, period_type)
        prompt = self.get_prompt("report_generation").format(
            report_type=period_type, period=period_str,
            report_title=title,
            current_datetime=datetime.now().strftime("%Y-%m-%d %H:%M"),
            work_data=work_data_str)
        content = self._call(prompt, require_json=False)
        if isinstance(content, str):
            return {"raw": content}
        return content

    def classify(self, text: str) -> dict:
        result = self._call(text, require_json=True)
        return result if isinstance(result, dict) else {}

    def similar(self, text1: str, text2: str) -> float:
        prompt = self.get_prompt("similarity_check").format(text1=text1, text2=text2)
        result = self._call(prompt, require_json=True)
        try:
            return float(result.get("similarity", 0.0))
        except (ValueError, TypeError):
            return 0.0

    def analyze_communication(self, content: str) -> dict:
        prompt = self.get_prompt("communication_analysis").format(content=content)
        result = self._call(prompt, require_json=True)
        return result if isinstance(result, dict) else {}

    def validate_customer(self, content: str, extracted: str, existing: list) -> dict:
        prompt = self.get_prompt("fuzzy_match").format(
            content=content, extracted_customer=extracted,
            existing_customers=json.dumps(existing, ensure_ascii=False))
        result = self._call(prompt, require_json=True)
        return result if isinstance(result, dict) else {}

    def classify_work(self, content: str, existing_customers: list, existing_projects: list) -> dict:
        """工作日记分类与信息提取（供 kingrecord 使用）。

        Args:
            content: 用户输入的工作内容
            existing_customers: 已有客户名称列表（用于匹配校验）
            existing_projects: 已有项目名称列表（用于匹配校验）

        Returns:
            dict，包含 work_type, confidence, extracted_info
        """
        prompt = self.get_prompt("work_classification").format(
            user_input=content,
            existing_customers=json.dumps(existing_customers, ensure_ascii=False),
            existing_projects=json.dumps(existing_projects, ensure_ascii=False),
        )
        result = self._call(prompt, require_json=True)
        if isinstance(result, dict):
            # 确保 extracted_info 存在
            if "extracted_info" not in result:
                result["extracted_info"] = {}
        else:
            result = {"work_type": "其他", "confidence": 0.0, "extracted_info": {}}
        return result

    def _call(self, prompt: str, require_json: bool = False) -> str | dict | list:
        """实际 API 调用，支持模型 fallback。"""
        fallback_models = self.config.get("llm", {}).get("model_fallback") or [self.model]
        fallback_models = [m for m in fallback_models if m]

        for fi, model_name in enumerate(fallback_models):
            for attempt in range(self.max_retries):
                try:
                    body = {
                        "model": model_name,
                        "messages": [{"role": "user", "content": prompt}],
                        "max_tokens": self.max_tokens,
                        "temperature": self.temperature,
                    }
                    resp = requests.post(
                        self.endpoint,
                        headers={"Authorization": f"Bearer {self.api_key}"},
                        json=body,
                        timeout=self.timeout,
                    )
                    if resp.status_code == 429:
                        wait = float(resp.headers.get("Retry-After", 2 ** attempt))
                        print(f"  ⚠️ Rate limit（{model_name}），{wait:.1f}s 后重试...")
                        time.sleep(wait)
                        continue
                    if resp.status_code >= 500:
                        print(f"  ⚠️ 服务端错误 {resp.status_code}（{model_name}），重试...")
                        time.sleep(2 ** attempt)
                        continue
                    data = resp.json()
                    if data.get("error"):
                        raise Exception(data["error"])
                    if fi > 0:
                        print(f"  🔄 切换至 {model_name} 成功")
                    return self._parse_response(data, require_json)
                except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
                    print(f"  ⚠️ 连接错误（{model_name}）：{e}，重试...")
                    time.sleep(2 ** attempt)
                    continue
                except Exception as e:
                    if fi < len(fallback_models) - 1:
                        print(f"  ⚠️ LLM 错误（{model_name}）：{e}，切换模型...")
                        break
                    raise
            if fi < len(fallback_models) - 1:
                print(f"  🔄 模型 {model_name} 不可用，切换到下一个...")
        raise Exception("LLM 调用失败（所有模型均不可用）")

    def _parse_response(self, data: dict, require_json: bool) -> str | dict | list:
        try:
            content = data["choices"][0]["messages"][0]["content"]
        except (KeyError, IndexError):
            try:
                content = data["choices"][0]["message"]["content"]
            except (KeyError, IndexError):
                raise Exception(f"API 返回格式异常：{data}")
        content = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL).strip()
        if not require_json:
            return content
        for pattern in [r"```json\s*(.*?)\s*```", r"```\s*(.*?)\s*```", r"(\{.*\}|\[.*\])"]:
            m = re.search(pattern, content, re.DOTALL)
            if m:
                try:
                    return json.loads(m.group(1).strip())
                except json.JSONDecodeError:
                    pass
        raise Exception(f"无法解析 JSON：{content[:100]}")


class LLMClient(KingWorkLLMClient):
    """兼容性别名"""

    def analyze_message(self, message_content: str, message_time: str = "", sender: str = "", chat_name: str = "") -> dict:
        """分析单条消息/会议内容，提取结构和关键信息。"""
        prompt = self.get_prompt("content_analysis").format(
            message_content=message_content,
            message_time=message_time,
            sender=sender,
            chat_name=chat_name,
        )
        result = self._call(prompt, require_json=True)
        if isinstance(result, dict):
            return result
        return {}
