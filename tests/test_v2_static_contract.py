from __future__ import annotations

import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_all_application_modules_parse_and_have_no_unimplemented_handlers() -> None:
    for path in (ROOT / "src").rglob("*.py"):
        source = path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(path))
        assert "NotImplementedError" not in source
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                assert not (len(node.body) == 1 and isinstance(node.body[0], ast.Pass)), f"empty handler: {path}:{node.name}"


def test_integrated_version_and_dependencies() -> None:
    assert (ROOT / "VERSION").read_text(encoding="utf-8").strip() == "2.0.0"
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert 'version = "2.0.0"' in pyproject
    assert 'openpyxl==3.1.5' in pyproject
    assert 'xlrd==2.0.2' in pyproject
    spec = (ROOT / "material_document_standardization.spec").read_text(encoding="utf-8")
    assert 'collect_submodules("material_claim_manager")' in spec


def test_purchase_ui_amount_and_layout_contract() -> None:
    ui_source = (ROOT / "src/purchase_request_app/ui.py").read_text(encoding="utf-8")
    assert "자동입력값 다시 불러오기" not in ui_source
    assert "간이 미리보기" not in ui_source
    assert 'self.budget_label.setText(f"₩ {format_money(data.budget_amount)}")' in ui_source
    assert 'self.contract_label.setText(f"₩ {format_money(data.contract_amount)}")' in ui_source
    assert 'self.budget_edit.setText(f"₩ {format_money(data.purchase_budget_effective)}")' in ui_source
    assert 'self.contract_edit.setText(f"₩ {format_money(data.purchase_contract_effective)}")' in ui_source
    assert "value.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)" in ui_source
    assert "card.setMaximumHeight(190)" in ui_source
    assert 'build_section_card("1. 내역서 기본정보")' in ui_source
    assert 'build_section_card("2. 구매물품 내역")' in ui_source
    assert 'self.quote_splitter.addWidget(summary_group)' not in ui_source
    assert 'self.quote_splitter = ElegantSplitter(Qt.Orientation.Vertical)' in ui_source
    assert 'self.main_splitter = ElegantSplitter(Qt.Orientation.Horizontal)' in ui_source


def test_material_claim_ui_naming_alignment_and_color_contract() -> None:
    claim_ui = (ROOT / "src/material_claim_manager/ui.py").read_text(encoding="utf-8")
    launcher = (ROOT / "src/material_document_app/launcher.py").read_text(encoding="utf-8")
    expense_ui = (ROOT / "src/expense_statement_app/ui.py").read_text(encoding="utf-8")

    assert 'APP_TITLE = "자재 입고 청구관리"' in claim_ui
    assert "자재입출고·청구관리" not in claim_ui
    assert "자재입출고·청구관리" not in launcher
    assert 'upper_splitter.setMaximumHeight(420)' in claim_ui
    assert 'self.tabs.setMinimumHeight(350)' in claim_ui
    assert '("current_intake", "당월 입고분", "#2563EB")' in claim_ui
    assert '("brought_in", "전월 이월분", "#0F8F70")' in claim_ui
    assert '("moved_out", "다음 달 이월분", "#D97706")' in claim_ui
    assert '("claim_target", "실제 청구대상", "#7C3AED")' in claim_ui
    assert 'QWidget#claimRoot, QWidget#claimWorkspace { background: #FFFFFF; }' in claim_ui
    assert 'QFrame#sectionCard, QFrame#previewCard { background: #FFFFFF;' in expense_ui
