from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from material_claim_manager.storage import ClaimOverrideStore


def test_management_values_are_saved_outside_original_workbook(tmp_path):
    workbook = tmp_path / "원본.xlsx"
    workbook.write_bytes(b"not modified")
    before = workbook.read_bytes()
    store = ClaimOverrideStore(str(workbook))
    store.set_classification("row1", "가설", "잡자재", reason="예산분류")
    store.exclude("row2", reason="타 공종")
    store.set("row3", "2026-07", "2026-06")

    assert workbook.read_bytes() == before
    assert store.path.exists()
    reloaded = ClaimOverrideStore(str(workbook))
    assert reloaded.classification("row1") == "가설"
    assert reloaded.is_excluded("row2")
    assert reloaded.get("row3") == "2026-07"


def test_automatic_backups_keep_recent_five_and_restore(tmp_path):
    workbook = tmp_path / "원본.xlsx"
    workbook.write_bytes(b"immutable workbook")
    original = workbook.read_bytes()
    store = ClaimOverrideStore(str(workbook))
    for index in range(8):
        store.set("row", f"2026-{(index % 12) + 1:02d}", "2026-01", reason=str(index))
    assert workbook.read_bytes() == original
    assert len(store.available_backups()) == 5

    manual = tmp_path / "manual.json"
    store.create_manual_backup(manual)
    store.set_classification("row", "가설", "잡자재")
    assert store.classification("row") == "가설"
    store.restore_backup(manual)
    assert store.classification("row") is None
