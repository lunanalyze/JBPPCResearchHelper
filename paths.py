from __future__ import annotations

import os
import shutil
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
APP_DATA_DIR = Path(os.environ["LOCALAPPDATA"]) / "PPCResearchAutomation"
CONFIG_DIR = APP_DATA_DIR / "config"
RUNS_DIR = APP_DATA_DIR / "runs"
RESOURCES_DIR = APP_DATA_DIR / "resources"


def ensure_app_dirs() -> None:
    for path in [APP_DATA_DIR, CONFIG_DIR, RUNS_DIR, RESOURCES_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def copy_default_resource(filename: str, default_text: str | None = None, overwrite: bool = False) -> Path:
    ensure_app_dirs()
    target = RESOURCES_DIR / filename
    if target.exists() and not overwrite:
        return target
    source = APP_DIR / filename
    if source.exists():
        shutil.copy2(source, target)
    elif default_text is not None:
        target.write_text(default_text, encoding="utf-8")
    return target
