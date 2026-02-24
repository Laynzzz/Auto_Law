"""Config loader & validator for the billing system."""

import json
import sys
import warnings
from pathlib import Path


def _get_base_dir() -> Path:
    """Return the project root directory.

    When running as a PyInstaller bundle (frozen), the base directory is the
    folder containing the .exe.  In normal dev mode it is the repo root
    (one level up from the ``src/`` package).
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE_DIR = _get_base_dir()
CONFIG_PATH = BASE_DIR / "config" / "config.json"


def load_config(path: Path = CONFIG_PATH) -> dict:
    """Load and validate config.json."""
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        config = json.load(f)

    _validate(config)
    return config


def get_data_root(config: dict | None = None) -> Path:
    """Return the shared data root path.

    If shared_root is set in config, returns that path.
    Otherwise returns the local project root (backward compatible).
    """
    if config is None:
        config = load_config()
    shared = config.get("shared_root", "")
    if shared:
        return Path(shared)
    return CONFIG_PATH.parent.parent


def _validate(config: dict) -> None:
    """Validate required config structure."""
    if "firms" not in config or not config["firms"]:
        raise ValueError("Config must contain a non-empty 'firms' list")

    for i, firm in enumerate(config["firms"]):
        for key in ("name", "initials"):
            if key not in firm:
                raise ValueError(f"Firm #{i} missing required key '{key}'")

    if "paths" not in config:
        raise ValueError("Config must contain a 'paths' section")

    # Warn (not crash) if shared_root is set but path doesn't exist
    shared = config.get("shared_root", "")
    if shared and not Path(shared).exists():
        warnings.warn(
            f"shared_root path does not exist: {shared}\n"
            "Falling back to local project root.",
            stacklevel=2,
        )

    # Validate that declared paths exist — use shared_root when set,
    # otherwise fall back to the directory containing config/.
    if shared and Path(shared).exists():
        project_root = Path(shared)
    else:
        project_root = CONFIG_PATH.parent.parent
    for label, rel in config["paths"].items():
        p = project_root / rel
        if not p.exists():
            raise ValueError(f"Path '{label}' does not exist: {p}")


def get_firm(name: str, config: dict | None = None) -> dict:
    """Look up a firm by name (case-insensitive)."""
    if config is None:
        config = load_config()

    target = name.lower()
    for firm in config["firms"]:
        if firm["name"].lower() == target:
            return firm

    available = [f["name"] for f in config["firms"]]
    raise KeyError(f"Firm '{name}' not found. Available: {available}")
