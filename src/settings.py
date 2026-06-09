from __future__ import annotations

import json
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = PROJECT_ROOT / "config.json"
EXAMPLE_CONFIG_PATH = PROJECT_ROOT / "config.example.json"


def load_config() -> dict[str, Any]:
    path = CONFIG_PATH if CONFIG_PATH.exists() else EXAMPLE_CONFIG_PATH
    with path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
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
