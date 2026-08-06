from pathlib import Path


def test_integrated_required_files_exist() -> None:
    root = Path(__file__).resolve().parents[1]
    required = [
        root / "main.py",
        root / "material_document_standardization.spec",
        root / "templates" / "purchase_request_3sheet_template.xlsx",
        root / "templates" / "expense_statement_template.xlsx",
        root / ".github" / "workflows" / "build-windows.yml",
        root / ".github" / "workflows" / "ci.yml",
        root / ".github" / "dependabot.yml",
        root / "assets" / "purchase_set.svg",
        root / "assets" / "expense_large.svg",
        root / "assets" / "database.svg",
        root / "src" / "material_claim_manager" / "ui.py",
        root / "src" / "material_claim_manager" / "excel_io.py",
        root / "examples" / "sample_standard_material_ledger.xlsx",
    ]
    assert all(path.exists() for path in required)


def test_archive_paths_are_ascii() -> None:
    root = Path(__file__).resolve().parents[1]
    ignored = {".pytest_cache", "__pycache__", "build", "dist"}
    for path in root.rglob("*"):
        if not path.is_file() or any(part in ignored for part in path.relative_to(root).parts):
            continue
        assert path.relative_to(root).as_posix().isascii()
