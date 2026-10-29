"""Load only this module's settings without changing other tools' environment."""

import os
from pathlib import Path

from dotenv import dotenv_values


class ConfigurationError(Exception):
    """An initialization error whose message contains no configuration values."""


def load_environment():
    path = Path(__file__).resolve().parent.parent / ".env"
    if not path.is_file():
        return
    try:
        # Keep literal password characters, including ${...}, and accept a UTF-8 BOM.
        values = dotenv_values(path, encoding="utf-8-sig", interpolate=False)
    except (OSError, UnicodeError) as exc:
        raise ConfigurationError("无法读取项目根目录的 .env，请检查文件权限并使用 UTF-8 编码保存") from exc
    for name, value in values.items():
        if name.startswith("JIRA_CREATOR_") and value is not None:
            os.environ.setdefault(name, value)
