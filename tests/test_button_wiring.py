from __future__ import annotations

import ast
from pathlib import Path


UI_FILES = (
    Path("src/purchase_request_app/ui.py"),
    Path("src/expense_statement_app/ui.py"),
    Path("src/material_claim_manager/ui.py"),
    Path("src/material_document_app/launcher.py"),
)
INHERITED_QT_HANDLERS = {
    "accept",
    "close",
    "reject",
    "showMaximized",
    "showMinimized",
    "showNormal",
}


def test_direct_qt_signal_handlers_are_implemented() -> None:
    missing: list[str] = []
    connection_count = 0

    for path in UI_FILES:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for class_node in (node for node in ast.walk(tree) if isinstance(node, ast.ClassDef)):
            methods = {
                node.name
                for node in class_node.body
                if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
            }
            for call in ast.walk(class_node):
                if not (
                    isinstance(call, ast.Call)
                    and isinstance(call.func, ast.Attribute)
                    and call.func.attr == "connect"
                    and call.args
                ):
                    continue
                handler = call.args[0]
                if not (
                    isinstance(handler, ast.Attribute)
                    and isinstance(handler.value, ast.Name)
                    and handler.value.id == "self"
                ):
                    continue
                connection_count += 1
                if handler.attr not in methods and handler.attr not in INHERITED_QT_HANDLERS:
                    missing.append(
                        f"{path}:{call.lineno} {class_node.name}.{handler.attr}"
                    )

    assert connection_count >= 80
    assert not missing, "연결되었지만 구현되지 않은 버튼 핸들러: " + ", ".join(missing)
