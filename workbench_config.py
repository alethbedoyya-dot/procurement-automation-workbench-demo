"""Load private runtime configuration for browser automation.

The repository intentionally contains no organization-specific web address.  A
developer supplies their own authorized endpoint through ``config.local.json``
or environment variables before running browser automation.
"""

import json
import os
from pathlib import Path
from urllib.parse import urlparse


CONFIG_FILE_NAME = "config.local.json"
HOME_URL_ENV = "PROCUREMENT_HOME_URL"
SEARCH_URL_ENV = "PROCUREMENT_SEARCH_URL"


class ConfigurationError(RuntimeError):
    """Raised when required private runtime settings are missing or invalid."""


def _read_config_file(config_path):
    if not config_path.exists():
        return {}

    try:
        with config_path.open("r", encoding="utf-8") as config_file:
            data = json.load(config_file)
    except json.JSONDecodeError as error:
        raise ConfigurationError(
            f"配置文件不是有效 JSON：{config_path.name}。"
        ) from error

    if not isinstance(data, dict):
        raise ConfigurationError(f"配置文件顶层必须是对象：{config_path.name}。")
    return data


def _validate_url(value, setting_name):
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"缺少 {setting_name}。")

    url = value.strip()
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ConfigurationError(f"{setting_name} 必须是完整的 http 或 https URL。")
    return url


def load_web_urls(*, config_path=None, environ=None):
    """Return ``(home_url, search_url)`` without exposing private defaults.

    Environment variables override the values in the local JSON configuration.
    ``config_path`` and ``environ`` are injectable so the behavior is easy to
    test without touching a developer's real configuration.
    """
    environment = os.environ if environ is None else environ
    local_path = (
        Path(config_path)
        if config_path is not None
        else Path(__file__).with_name(CONFIG_FILE_NAME)
    )
    config = _read_config_file(local_path)
    web_config = config.get("web", {})
    if not isinstance(web_config, dict):
        raise ConfigurationError("配置项 web 必须是对象。")

    home_url = environment.get(HOME_URL_ENV) or web_config.get("home_url")
    search_url = environment.get(SEARCH_URL_ENV) or web_config.get("search_url")
    if not home_url or not search_url:
        raise ConfigurationError(
            "缺少网页地址配置。请复制 config.example.json 为 config.local.json，"
            f"或设置 {HOME_URL_ENV} 与 {SEARCH_URL_ENV}。"
        )

    return (
        _validate_url(home_url, "web.home_url"),
        _validate_url(search_url, "web.search_url"),
    )
