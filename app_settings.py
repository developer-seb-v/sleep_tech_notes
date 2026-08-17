"""
Small persistent app settings (currently just the document logo path).

Settings live next to this script in app_settings.json, so they survive
across runs without needing a database.
"""

import json
import os

SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_settings.json")


def _load():
    if not os.path.isfile(SETTINGS_PATH):
        return {}
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save(settings):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


def get_logo_path():
    """Return the configured logo file path, or None if not set."""
    return _load().get("logo_path")


def set_logo_path(path):
    """Persist `path` as the logo to use on saved tech notes."""
    settings = _load()
    settings["logo_path"] = path
    _save(settings)


def clear_logo_path():
    """Remove the configured logo, if any."""
    settings = _load()
    settings.pop("logo_path", None)
    _save(settings)
