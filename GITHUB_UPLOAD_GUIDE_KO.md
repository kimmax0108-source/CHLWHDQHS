# GitHub 업로드 및 EXE 생성 안내 — v2.0.0

## 1. 저장소 최상단에 업로드

배포 ZIP을 압축 해제한 뒤 내부 파일 전체를 GitHub 저장소 최상단에 업로드합니다.

이 정리본은 **총 92개 파일**이므로 GitHub 웹 업로드의 100개 파일 제한 안에서 한 번에 업로드할 수 있습니다. 저장소 첫 화면에 `main.py`, `src`, `.github`, `templates`가 바로 보여야 합니다.

## 2. `.github` 폴더 확인

압축을 푼 폴더 안의 `.github` 폴더도 함께 선택하여 업로드합니다. Windows 파일 탐색기에서는 일반 폴더처럼 표시됩니다. 업로드 후 아래 파일이 존재하는지 확인합니다.

```text
.github/dependabot.yml
.github/workflows/build-windows.yml
.github/workflows/ci.yml
```

## 3. 자동 테스트

`main` 또는 `master` 브랜치에 올리면 CI가 다음을 실행합니다.

```text
의존성 설치 → Ruff 검사 → Pytest 전체 테스트
```

## 4. Windows EXE 생성

```text
Actions → Build Windows EXE → Run workflow
```

완료 후 아래 아티팩트를 내려받습니다.

```text
material-document-standardization-v2-windows
  └─ material_document_standardization_v2.0.0_windows.zip
```

## 5. GitHub Release

`v2.0.0` 태그를 푸시하면 테스트와 EXE 빌드 후 Release가 자동 생성됩니다.

## 6. 저장소 루트 확인

```text
.github/
assets/
docs/
examples/
presets/
src/
templates/
tests/
main.py
material_document_standardization.spec
requirements.txt
requirements-dev.txt
```

## 7. 제외한 항목

실행 또는 빌드에 필요하지 않은 다음 항목은 업로드 제한 대응을 위해 제외했습니다.

- UI 시안·출력 미리보기 PNG
- `.github` 중복 복사본이 들어 있던 업로드 보조 폴더
- 중복 예제 파일
- 미사용 SVG 아이콘
- 중복 배포 보고서
