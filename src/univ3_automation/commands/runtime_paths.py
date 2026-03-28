from __future__ import annotations

import json
import os
from functools import lru_cache
from pathlib import Path
from typing import Any

DEFAULT_SHEET_ID = "1fCu0cwgh5Tynw2veeEx9zAJMepwWDRFhWdCGLDRUzT0"


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[3]


def _expand_env(value: str) -> str:
    expanded = os.path.expandvars(value)
    if expanded.startswith("~"):
        return str(Path(expanded).expanduser())
    return expanded


def _resolve_path(value: str, base: Path) -> Path:
    candidate = Path(_expand_env(value.replace("\\", "/")))
    if candidate.is_absolute():
        return candidate
    return (base / candidate).resolve()


@lru_cache(maxsize=1)
def load_config() -> dict[str, Any]:
    root = _repo_root()
    config_dir = root / "config"
    local_path = config_dir / "local.json"
    if not local_path.exists():
        raise FileNotFoundError(f"Config file not found: {local_path}")
    return json.loads(local_path.read_text(encoding="utf-8-sig"))


def reload_config() -> dict[str, Any]:
    load_config.cache_clear()
    return load_config()


def get_base_dir() -> str:
    env_override = os.environ.get("UNIV3_BASE_DIR")
    if env_override:
        return str(_resolve_path(env_override, _repo_root()))

    config = load_config()
    base_value = config.get("paths", {}).get("base_dir", "")
    if base_value:
        return str(_resolve_path(str(base_value), _repo_root()))

    # Legacy fallback: parent of the commands directory
    return str(Path(__file__).resolve().parents[2])


def _get_directories() -> dict[str, Any]:
    config = load_config()
    return config.get("paths", {}).get("directories", {})


def get_dir(key: str, default_subdir: str | None = None) -> str:
    base = Path(get_base_dir())
    directories = _get_directories()
    configured = directories.get(key)
    if configured:
        return str(_resolve_path(str(configured), base))
    if default_subdir:
        return str((base / default_subdir).resolve())
    return str(base)


def get_downloads_dir() -> str:
    return get_dir("downloads", "Downloads")


def get_kakao_downloads_dir() -> str:
    # Prefer explicit kakao_downloads key but keep backward compatibility.
    directories = _get_directories()
    if "kakao_downloads" in directories:
        return get_dir("kakao_downloads")
    if "kakaoDownloads" in directories:
        return get_dir("kakaoDownloads")
    return str(Path("D:/Data/Documents/카카오톡 받은 파일"))


def get_screenshot_dirs() -> list[str]:
    directories = _get_directories()
    configured = directories.get("screenshots", [])
    base = Path(get_base_dir())
    resolved: list[str] = []
    if isinstance(configured, list):
        for entry in configured:
            resolved.append(str(_resolve_path(str(entry), base)))
    if not resolved:
        resolved = [
            str(Path("D:/Data/Pictures/Screenshots")),
            str(Path.home() / "Pictures" / "Screenshots"),
            str(Path.home() / "OneDrive" / "Pictures" / "Screenshots"),
            str(Path.home() / "Desktop"),
        ]
    return resolved


def get_env_file_path() -> Path:
    env_override = os.environ.get("UNIV3_ENV_FILE")
    if env_override:
        return _resolve_path(env_override, _repo_root())

    config = load_config()
    configured = config.get("email", {}).get("env_file")
    if configured:
        return _resolve_path(str(configured), _repo_root())

    return _repo_root() / ".env"


def get_google_sheet_id(default: str = DEFAULT_SHEET_ID) -> str:
    config = load_config()
    return str(config.get("google", {}).get("sheet_id", default))


def get_google_credentials_candidates() -> list[Path]:
    config = load_config()
    root = _repo_root()
    base = Path(get_base_dir())

    candidates: list[Path] = []

    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if env_path:
        candidates.append(_resolve_path(env_path, root))

    configured = config.get("google", {}).get("credentials_path")
    if configured:
        candidates.append(_resolve_path(str(configured), root))
        candidates.append(_resolve_path(str(configured), base))

    candidates.extend(
        [
            base / "코드" / "credentials.json",
            base / "credentials.json",
            root / "credentials.json",
            root / "config" / "secrets" / "google-service-account.json",
        ]
    )

    seen: set[str] = set()
    unique: list[Path] = []
    for candidate in candidates:
        try:
            exists = candidate.exists()
        except OSError:
            exists = False

        if exists:
            try:
                key = str(candidate.resolve())
            except OSError:
                key = str(candidate)
        else:
            key = str(candidate)

        if key in seen:
            continue
        seen.add(key)
        unique.append(candidate)

    return unique




def get_birthday_selector_columns() -> dict[str, str]:
    config = load_config()
    columns = config.get("birthday_selector", {})
    return {
        "name_column": str(columns.get("name_column", "G")),
        "gender_column": str(columns.get("gender_column", "E")),
        "grade_column": str(columns.get("grade_column", "F")),
        "birth_year_column": str(columns.get("birth_year_column", "AM")),
        "birth_month_column": str(columns.get("birth_month_column", "AN")),
        "birth_day_column": str(columns.get("birth_day_column", "AO")),
    }

def get_email_recipients(default: list[str] | None = None) -> list[str]:
    if default is None:
        default = []
    config = load_config()
    recipients = config.get("email", {}).get("recipients")
    if isinstance(recipients, list) and recipients:
        return [str(item) for item in recipients]
    return default

