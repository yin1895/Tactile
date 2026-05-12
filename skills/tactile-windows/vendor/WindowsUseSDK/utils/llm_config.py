from __future__ import annotations

import ast
import json
import os
import re
import time


DEFAULT_LLM_MODEL = os.getenv("TACTILE_MODEL", "gpt-5.5")
LLM_TIMEOUT_SECONDS = float(os.getenv("TACTILE_LLM_TIMEOUT", "600"))
LLM_MAX_RETRIES = int(os.getenv("TACTILE_LLM_MAX_RETRIES", "3"))
LLM_RETRY_DELAY_SECONDS = float(os.getenv("TACTILE_LLM_RETRY_DELAY", "5"))
PROXY_ENV_KEYS = (
    "ALL_PROXY",
    "all_proxy",
    "HTTP_PROXY",
    "http_proxy",
    "HTTPS_PROXY",
    "https_proxy",
)


def _env_key(provider: str, suffix: str) -> str:
    normalized = re.sub(r"[^A-Za-z0-9]+", "_", provider).strip("_").upper()
    return f"TACTILE_{normalized}_{suffix}"


def _client_config(provider: str | None = None) -> tuple[str, str | None]:
    api_key = os.getenv("TACTILE_OPENAI_API_KEY")
    base_url = os.getenv("TACTILE_OPENAI_BASE_URL")
    if provider:
        api_key = os.getenv(_env_key(provider, "API_KEY")) or api_key
        base_url = os.getenv(_env_key(provider, "BASE_URL")) or base_url
    if not api_key:
        raise RuntimeError("missing LLM API key; set TACTILE_OPENAI_API_KEY")
    return api_key, base_url


def _drop_unsupported_socks_proxy_env() -> dict[str, str]:
    try:
        import socksio  # noqa: F401
        return {}
    except ImportError:
        pass

    removed: dict[str, str] = {}
    for key in PROXY_ENV_KEYS:
        value = os.environ.get(key)
        if value and value.lower().startswith(("socks://", "socks4://", "socks5://", "socks5h://")):
            removed[key] = value
            os.environ.pop(key, None)
    return removed


def _restore_env(values: dict[str, str]) -> None:
    for key, value in values.items():
        os.environ[key] = value


def call_llm(
    prompt: str,
    model_name: str = DEFAULT_LLM_MODEL,
    provider: str | None = None,
    image_base64: str | list[str] | None = None,
    temperature: float = 0,
    top_p: float = 1,
) -> str:
    from openai import OpenAI

    api_key, base_url = _client_config(provider)
    removed_proxy_env = _drop_unsupported_socks_proxy_env()
    try:
        client = OpenAI(
            api_key=api_key,
            base_url=base_url,
            timeout=LLM_TIMEOUT_SECONDS,
            max_retries=0,
        )
    finally:
        _restore_env(removed_proxy_env)

    content: str | list[dict[str, object]] = prompt
    if isinstance(image_base64, str):
        image_base64 = [image_base64]
    if image_base64:
        content = [{"type": "text", "text": prompt}]
        for image in image_base64:
            content.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image}"}})

    messages = [{"role": "user", "content": content}]
    response = None
    for attempt in range(1, LLM_MAX_RETRIES + 1):
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=messages,
                stream=False,
                temperature=temperature,
                top_p=top_p,
            )
            break
        except Exception:
            if attempt >= LLM_MAX_RETRIES:
                raise
            time.sleep(LLM_RETRY_DELAY_SECONDS * attempt)

    if response is None or not response.choices:
        raise RuntimeError("LLM response did not include choices")
    return response.choices[0].message.content or ""


def extract_and_convert_dict(text: str):
    def find_balanced_braces(value: str) -> list[tuple[int, int]]:
        candidates: list[tuple[int, int]] = []
        stack: list[int] = []
        for index, char in enumerate(value):
            if char == "{":
                stack.append(index)
            elif char == "}" and stack:
                start = stack.pop()
                candidates.append((start, index + 1))
        candidates.sort(key=lambda pair: -(pair[1] - pair[0]))
        return candidates

    def collapse_paren_string_concat(value: str) -> str:
        string_literal = r'"[^"\\]*(?:\\.[^"\\]*)*"'
        pattern = re.compile(
            r"\(\s*(" + string_literal + r"(?:\s*" + string_literal + r")*)\s*\)",
            re.DOTALL,
        )

        def replace(match: re.Match[str]) -> str:
            inner = match.group(1)
            parts = re.findall(r'"((?:[^"\\]|\\.)*)"', inner, re.DOTALL)
            joined = "".join(parts).replace("\n", "\\n").replace("\r", "\\r")
            return '"' + joined + '"'

        return pattern.sub(replace, value)

    parsers = (json.loads, ast.literal_eval)
    for start, end in find_balanced_braces(text):
        candidate = text[start:end]
        for parser in parsers:
            try:
                result = parser(candidate)
                if isinstance(result, dict):
                    return result
            except Exception:
                pass

        collapsed = collapse_paren_string_concat(candidate)
        if collapsed == candidate:
            continue
        for parser in parsers:
            try:
                result = parser(collapsed)
                if isinstance(result, dict):
                    return result
            except Exception:
                pass
    return None
