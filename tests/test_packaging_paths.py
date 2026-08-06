from pathlib import Path


def test_project_paths_are_ascii() -> None:
    root = Path(__file__).resolve().parents[1]
    for path in root.rglob("*"):
        relative = path.relative_to(root).as_posix()
        assert relative.isascii(), f"Non-ASCII packaged path: {relative}"


def test_required_resources_exist() -> None:
    root = Path(__file__).resolve().parents[1]
    assert (root / "templates" / "purchase_request_3sheet_template.xlsx").exists()
    assert (root / ".github" / "workflows" / "build-windows.yml").exists()
