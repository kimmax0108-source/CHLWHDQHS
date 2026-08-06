from __future__ import annotations

import json
import shutil
from copy import deepcopy
from pathlib import Path
from typing import Any, Literal

PresetKind = Literal["sites", "items"]


class PresetStore:
    """JSON-backed site/item preset storage.

    The bundled defaults are copied into the user's application data directory on
    first launch. Subsequent edits never modify the bundled file.
    """

    def __init__(self, data_dir: Path, defaults_path: Path) -> None:
        self.data_dir = Path(data_dir)
        self.defaults_path = Path(defaults_path)
        self.path = self.data_dir / "presets.json"
        self.data_dir.mkdir(parents=True, exist_ok=True)
        if not self.path.exists():
            shutil.copy2(self.defaults_path, self.path)
        self._data = self._read()

    def _read(self) -> dict[str, Any]:
        try:
            with self.path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            with self.defaults_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
            self._write(data)
        data.setdefault("version", 1)
        data.setdefault("settings", {})
        data.setdefault("sites", [])
        data.setdefault("items", [])
        return data

    def _write(self, data: dict[str, Any] | None = None) -> None:
        payload = self._data if data is None else data
        temp = self.path.with_suffix(".tmp")
        with temp.open("w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
        temp.replace(self.path)

    def reload(self) -> None:
        self._data = self._read()

    def all(self, kind: PresetKind) -> list[dict[str, Any]]:
        return deepcopy(self._data.get(kind, []))

    def names(self, kind: PresetKind) -> list[str]:
        return [str(item.get("name", "")) for item in self._data.get(kind, [])]

    def find(self, kind: PresetKind, name: str) -> dict[str, Any] | None:
        for item in self._data.get(kind, []):
            if item.get("name") == name:
                return deepcopy(item)
        return None

    def upsert(self, kind: PresetKind, preset: dict[str, Any]) -> None:
        name = str(preset.get("name", "")).strip()
        if not name:
            raise ValueError("프리셋 이름이 비어 있습니다.")
        preset = deepcopy(preset)
        preset["name"] = name
        items = self._data.setdefault(kind, [])
        for index, existing in enumerate(items):
            if existing.get("name") == name:
                items[index] = preset
                break
        else:
            items.append(preset)
        items.sort(key=lambda item: str(item.get("name", "")))
        self._write()

    def delete(self, kind: PresetKind, name: str) -> bool:
        items = self._data.setdefault(kind, [])
        original = len(items)
        self._data[kind] = [item for item in items if item.get("name") != name]
        changed = len(self._data[kind]) != original
        if changed:
            self._write()
        return changed

    def setting(self, key: str, default: Any = None) -> Any:
        return deepcopy(self._data.setdefault("settings", {}).get(key, default))

    def set_setting(self, key: str, value: Any) -> None:
        self._data.setdefault("settings", {})[key] = value
        self._write()

    def reset_defaults(self) -> None:
        shutil.copy2(self.defaults_path, self.path)
        self.reload()
