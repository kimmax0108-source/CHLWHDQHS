from __future__ import annotations

import json
from pathlib import Path

from purchase_request_app.preset_store import PresetStore


def test_preset_crud(tmp_path: Path) -> None:
    defaults = tmp_path / "defaults.json"
    defaults.write_text(
        json.dumps({"version": 1, "settings": {}, "sites": [], "items": []}, ensure_ascii=False),
        encoding="utf-8",
    )
    store = PresetStore(tmp_path / "data", defaults)
    store.upsert("sites", {"name": "테스트현장", "short_name": "테스트"})
    assert store.find("sites", "테스트현장")["short_name"] == "테스트"
    store.set_setting("last_site", "테스트현장")
    assert store.setting("last_site") == "테스트현장"
    assert store.delete("sites", "테스트현장") is True
    assert store.find("sites", "테스트현장") is None
