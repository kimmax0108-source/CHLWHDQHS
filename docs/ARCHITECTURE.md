# 자재 문서 표준화 v2.0.0 구조

## 실행 흐름

```text
main.py
  ↓
material_document_app.LauncherWindow
  ├─ 구매 품의서 양식 → purchase_request_app.MainWindow
  ├─ 지출결의서       → expense_statement_app.MainWindow
  └─ 자재 청구관리    → material_claim_manager.MaterialClaimWindow
```

세 업무는 하나의 EXE와 공통 UI 셸을 사용하지만 작업 데이터와 출력은 서로 분리합니다.

## 계층

```text
UI (PySide6)
  ↓
업무 서비스/계산 정책
  ↓
모델
  ↓
Excel 입출력 / JSON 저장
```

### 통합 셸

- `material_document_app/launcher.py`: 3개 업무 카드, 좌측 직접 메뉴, 설정/정보
- `material_document_app/resource.py`: 개발 실행과 PyInstaller 실행의 리소스 경로 통합

### 구매 품의서 양식

- `purchase_request_app/models.py`: Decimal 계산, 업체 순위, 가실행, 원단위 정책
- `purchase_request_app/ui.py`: 3단계 작성 화면, 가변 표, 프리셋, 미리보기
- `purchase_request_app/xlsx_engine.py`: 3개 시트 OOXML 출력
- `purchase_request_app/preset_store.py`: 현장/품목 프리셋 및 작업 저장

### 지출결의서

- `expense_statement_app/models.py`: 안분·계좌·품목 데이터
- `expense_statement_app/calculations.py`: 한글금액, 수량 서식, 합계
- `expense_statement_app/ui.py`: 입력·미리보기·내보내기
- `expense_statement_app/xlsx_engine.py`: 동적 행·도형·인쇄 설정 보존

### 자재 입고 청구관리

- `material_claim_manager/models.py`: 원본 행, 지문, 공급가/VAT/검증 상태
- `material_claim_manager/excel_io.py`: xls/xlsx/xlsm, 표준 다중 시트, 헤더 별칭
- `material_claim_manager/services.py`: 필터, 집계, 이월, 제외, 분류, 상태 계산
- `material_claim_manager/storage.py`: 숨김 JSON, 이력, 자동백업 5개, 수동 복원
- `material_claim_manager/exporter.py`: 조회결과 Excel 출력
- `material_claim_manager/ui.py`: PySide6 통합 화면

## 원본 보존 원칙

1. 자재대장은 `data_only=True`, 읽기 전용 모드로 해석합니다.
2. 청구월·분류·제외·메모는 원본 셀에 쓰지 않습니다.
3. 각 원본 행은 `원본 시트 + 원본 행 + 주요값` 기반 지문으로 관리합니다.
4. 관리 JSON 저장 전 기존 JSON을 자동 백업합니다.
5. Excel 내보내기는 새로운 파일에만 수행합니다.

## 패키징

`material_document_standardization.spec`는 다음 항목을 포함합니다.

- 세 업무 패키지의 하위 모듈
- 구매 품의서/지출결의서 Excel 템플릿
- 기본 프리셋
- 공통 SVG/ICO 리소스

GitHub Actions는 Windows Python 3.12에서 테스트 후 단일 EXE와 배포 ZIP을 생성합니다.
