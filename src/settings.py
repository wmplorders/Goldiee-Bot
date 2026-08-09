from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = PROJECT_ROOT / "config.json"
EXAMPLE_CONFIG_PATH = PROJECT_ROOT / "config.example.json"
ENV_PATH = PROJECT_ROOT / ".env"


def load_config() -> dict[str, Any]:
    path = CONFIG_PATH if CONFIG_PATH.exists() else EXAMPLE_CONFIG_PATH
    with path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
    _apply_env_overrides(config, _load_env())
    return _resolve_paths(config)


def _resolve_paths(config: dict[str, Any]) -> dict[str, Any]:
    for key in ("downloads", "processed", "logs"):
        config["paths"][key] = str((PROJECT_ROOT / config["paths"][key]).resolve())
    for key in ("state", "processing_log"):
        config["paths"][key] = str((PROJECT_ROOT / config["paths"][key]).resolve())
    config["whatsapp"]["edge_profile_dir"] = str(
        (PROJECT_ROOT / config["whatsapp"]["edge_profile_dir"]).resolve()
    )
    return config


def _load_env() -> dict[str, str]:
    values: dict[str, str] = {}
    if ENV_PATH.exists():
        for raw_line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
    values.update({key: value for key, value in os.environ.items() if key.startswith("TALLY_")})
    return values


def _env_bool(value: str) -> bool:
    return value.strip().casefold() in {"1", "true", "yes", "y", "on"}


def _apply_env_overrides(config: dict[str, Any], env: dict[str, str]) -> None:
    tally = config.setdefault("tally", {})
    mapping = {
        "TALLY_DATA_PATH": "data_path",
        "TALLY_COMPANY_FOLDER": "company_folder",
        "TALLY_COMPANY_NAME": "company_name",
        "TALLY_HTTP_URL": "http_url",
    }
    for env_key, config_key in mapping.items():
        if env.get(env_key):
            tally[config_key] = env[env_key]
    for env_key, config_key in {
        "TALLY_REQUIRE_EXACT_COMPANY": "require_exact_company",
        "TALLY_CLOSE_OTHER_COMPANIES": "close_other_companies",
    }.items():
        if env_key in env:
            tally[config_key] = _env_bool(env[env_key])
