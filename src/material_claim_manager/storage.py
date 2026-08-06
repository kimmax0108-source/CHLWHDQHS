from __future__ import annotations

from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
from typing import Any, Optional


def _read_json(path: Path, fallback: dict[str, Any]) -> dict[str, Any]:
    if not path.exists():
        return fallback
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        return payload if isinstance(payload, dict) else fallback
    except (OSError, json.JSONDecodeError, TypeError):
        return fallback


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    temporary.replace(path)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


class ClaimOverrideStore:
    """원본 통합문서는 수정하지 않고 청구 관리값만 별도 JSON에 저장한다.

    v2.0에서는 저장 직전 관리 JSON을 자동 백업하며 최근 5개를 유지한다.
    원본 Excel은 읽기만 하며 이 클래스는 원본 경로에 쓰지 않는다.
    """

    BACKUP_LIMIT = 5

    def __init__(self, workbook_path: str):
        self.source_path = Path(workbook_path).resolve()
        self.path = self.source_path.with_name(f".{self.source_path.stem}_청구관리.json")
        self.backup_dir = self.source_path.with_name(
            f".{self.source_path.stem}_청구관리_백업"
        )
        self._claim_months: dict[str, str] = {}
        self._classifications: dict[str, str] = {}
        self._excluded: dict[str, dict[str, Any]] = {}
        self._management_notes: dict[str, str] = {}
        self._history: list[dict[str, Any]] = []
        self._source_sha256 = _sha256(self.source_path) if self.source_path.exists() else ""
        self.load()

    def load(self) -> None:
        payload = _read_json(self.path, {})
        # v1 호환: 예전 .청구이월.json이 있으면 이월정보만 승계한다.
        if not payload:
            old_path = self.path.with_name(
                self.path.name.replace("_청구관리.json", "_청구이월.json")
            )
            payload = _read_json(old_path, {})
        self._claim_months = dict(payload.get("claim_months", payload.get("overrides", {})))
        self._classifications = dict(payload.get("classifications", {}))
        self._excluded = dict(payload.get("excluded", {}))
        self._management_notes = dict(payload.get("management_notes", {}))
        history = payload.get("history", [])
        self._history = list(history) if isinstance(history, list) else []

    def _payload(self) -> dict[str, Any]:
        return {
            "version": 3,
            "workbook": {
                "name": self.source_path.name,
                "size": self.source_path.stat().st_size if self.source_path.exists() else 0,
                "sha256": self._source_sha256,
            },
            "claim_months": self._claim_months,
            "classifications": self._classifications,
            "excluded": self._excluded,
            "management_notes": self._management_notes,
            "history": self._history[-5000:],
        }

    def _backup_existing(self) -> Optional[Path]:
        if not self.path.exists():
            return None
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup = self.backup_dir / f"{self.path.stem}_{stamp}.json"
        shutil.copy2(self.path, backup)
        backups = sorted(
            self.backup_dir.glob("*.json"), key=lambda item: item.stat().st_mtime, reverse=True
        )
        for stale in backups[self.BACKUP_LIMIT :]:
            try:
                stale.unlink()
            except OSError:
                pass
        return backup

    def save(self) -> None:
        payload = self._payload()
        new_text = json.dumps(payload, ensure_ascii=False, indent=2)
        current_text = ""
        if self.path.exists():
            try:
                current_text = self.path.read_text(encoding="utf-8")
            except OSError:
                current_text = ""
        if current_text == new_text:
            return
        self._backup_existing()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        temporary.write_text(new_text, encoding="utf-8")
        temporary.replace(self.path)

    def create_manual_backup(self, destination: str | Path) -> Path:
        """현재 관리 JSON을 지정 위치에 복사한다."""
        if not self.path.exists():
            self.save()
        target = Path(destination)
        if target.is_dir() or not target.suffix:
            target.mkdir(parents=True, exist_ok=True)
            target = target / f"{self.source_path.stem}_청구관리_백업.json"
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.path, target)
        return target

    def restore_backup(self, backup_path: str | Path) -> None:
        source = Path(backup_path)
        payload = _read_json(source, {})
        if not payload or not any(
            key in payload for key in ("claim_months", "overrides", "classifications", "history")
        ):
            raise ValueError("올바른 청구관리 백업 JSON이 아닙니다.")
        self._backup_existing()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, self.path)
        self.load()

    def available_backups(self) -> list[Path]:
        if not self.backup_dir.exists():
            return []
        return sorted(
            self.backup_dir.glob("*.json"), key=lambda item: item.stat().st_mtime, reverse=True
        )

    def _record(
        self,
        action: str,
        fingerprint: str,
        *,
        old_value: Any = None,
        new_value: Any = None,
        reason: str = "",
        row_info: Optional[dict[str, Any]] = None,
    ) -> None:
        self._history.append(
            {
                "at": datetime.now().isoformat(timespec="seconds"),
                "action": action,
                "fingerprint": fingerprint,
                "old": old_value,
                "new": new_value,
                "reason": reason.strip(),
                "row": row_info or {},
            }
        )

    def get(self, fingerprint: str) -> Optional[str]:
        return self._claim_months.get(fingerprint)

    def set(
        self,
        fingerprint: str,
        claim_month: str,
        intake_month: str,
        *,
        reason: str = "",
        row_info: Optional[dict[str, Any]] = None,
        autosave: bool = True,
    ) -> None:
        old = self._claim_months.get(fingerprint, intake_month)
        if claim_month == intake_month:
            self._claim_months.pop(fingerprint, None)
        else:
            self._claim_months[fingerprint] = claim_month
        self._record(
            "claim_month",
            fingerprint,
            old_value=old,
            new_value=claim_month,
            reason=reason,
            row_info=row_info,
        )
        if autosave:
            self.save()

    def remove(
        self,
        fingerprint: str,
        *,
        intake_month: str = "",
        reason: str = "",
        row_info: Optional[dict[str, Any]] = None,
        autosave: bool = True,
    ) -> None:
        old = self._claim_months.pop(fingerprint, intake_month or None)
        self._record(
            "claim_month_reset",
            fingerprint,
            old_value=old,
            new_value=intake_month,
            reason=reason,
            row_info=row_info,
        )
        if autosave:
            self.save()

    def classification(self, fingerprint: str) -> Optional[str]:
        return self._classifications.get(fingerprint)

    def set_classification(
        self,
        fingerprint: str,
        classification: str,
        original_trade: str,
        *,
        reason: str = "",
        row_info: Optional[dict[str, Any]] = None,
        autosave: bool = True,
    ) -> None:
        old = self._classifications.get(fingerprint, original_trade)
        if not classification or classification == original_trade:
            self._classifications.pop(fingerprint, None)
        else:
            self._classifications[fingerprint] = classification
        self._record(
            "classification",
            fingerprint,
            old_value=old,
            new_value=classification or original_trade,
            reason=reason,
            row_info=row_info,
        )
        if autosave:
            self.save()

    def reset_classification(
        self,
        fingerprint: str,
        original_trade: str,
        *,
        reason: str = "",
        row_info: Optional[dict[str, Any]] = None,
        autosave: bool = True,
    ) -> None:
        old = self._classifications.pop(fingerprint, original_trade)
        self._record(
            "classification_reset",
            fingerprint,
            old_value=old,
            new_value=original_trade,
            reason=reason,
            row_info=row_info,
        )
        if autosave:
            self.save()

    def is_excluded(self, fingerprint: str) -> bool:
        value = self._excluded.get(fingerprint)
        return bool(value and value.get("excluded", True))

    def exclusion_reason(self, fingerprint: str) -> str:
        value = self._excluded.get(fingerprint) or {}
        return str(value.get("reason", ""))

    def exclude(
        self,
        fingerprint: str,
        *,
        reason: str = "",
        row_info: Optional[dict[str, Any]] = None,
        autosave: bool = True,
    ) -> None:
        self._excluded[fingerprint] = {
            "excluded": True,
            "reason": reason.strip(),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        self._record(
            "exclude",
            fingerprint,
            old_value=False,
            new_value=True,
            reason=reason,
            row_info=row_info,
        )
        if autosave:
            self.save()

    def restore_excluded(
        self,
        fingerprint: str,
        *,
        reason: str = "",
        row_info: Optional[dict[str, Any]] = None,
        autosave: bool = True,
    ) -> None:
        old = self._excluded.pop(fingerprint, None)
        self._record(
            "exclude_reset",
            fingerprint,
            old_value=bool(old),
            new_value=False,
            reason=reason,
            row_info=row_info,
        )
        if autosave:
            self.save()

    def management_note(self, fingerprint: str) -> str:
        return self._management_notes.get(fingerprint, "")

    def set_management_note(
        self,
        fingerprint: str,
        note: str,
        *,
        row_info: Optional[dict[str, Any]] = None,
        autosave: bool = True,
    ) -> None:
        old = self._management_notes.get(fingerprint, "")
        if note.strip():
            self._management_notes[fingerprint] = note.strip()
        else:
            self._management_notes.pop(fingerprint, None)
        self._record(
            "management_note",
            fingerprint,
            old_value=old,
            new_value=note.strip(),
            row_info=row_info,
        )
        if autosave:
            self.save()

    @property
    def history(self) -> list[dict[str, Any]]:
        return list(self._history)


class PresetStore:
    """통합문서를 바꿔도 유지되는 사용자 조회 프리셋."""

    def __init__(self):
        if os.name == "nt" and os.getenv("APPDATA"):
            root = Path(os.environ["APPDATA"]) / "MaterialDocumentStandardization"
        else:
            root = Path.home() / ".config" / "material-document-standardization"
        self.path = root / "claim_presets.json"
        self._presets: dict[str, dict[str, Any]] = {}
        self.load()

    def load(self) -> None:
        payload = _read_json(self.path, {})
        self._presets = dict(payload.get("presets", {}))

    def save(self) -> None:
        _write_json(self.path, {"version": 2, "presets": self._presets})

    def names(self) -> list[str]:
        return sorted(self._presets, key=str.casefold)

    def get(self, name: str) -> Optional[dict[str, Any]]:
        value = self._presets.get(name)
        return dict(value) if value else None

    def set(self, name: str, value: dict[str, Any]) -> None:
        self._presets[name] = dict(value)
        self.save()

    def delete(self, name: str) -> None:
        self._presets.pop(name, None)
        self.save()
