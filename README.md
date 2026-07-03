# TechManual Table Extractor

기술교범 한글 파일(.hwp / .hwpx)에서 **조건에 맞는 표만 선별**하여 원본 구조(병합 셀, 행/열 크기, 제목)를 유지한 채 엑셀(.xlsx)로 추출하는 RPA 도구입니다.

- OCR을 사용하지 않고 HWPX 내부 XML 구조를 직접 파싱합니다.
- `.hwp` 파일은 한글(Hancom Office) COM 자동화로 `.hwpx`로 변환 후 처리합니다. (`.hwpx`만 처리한다면 한글 설치 불필요)
- 표 선별은 **점수 기반 다중 기준 매칭**(헤더 일치율 + 제목 키워드 + 열 개수 + 셀 패턴)으로 동작하며, 규칙은 YAML 설정 파일로 관리합니다.

## 요구 사항

- Windows / Python 3.13 (64bit)
- `.hwp` 변환 기능 사용 시에만 한글(Hancom Office) 설치 필요

## 설치 (인터넷 없는 PC 포함)

`wheelhouse/` 폴더에 Python 3.13 (Windows 64bit)용 wheel이 모두 포함되어 있어 **인터넷 없이 설치 가능**합니다.

1. 이 프로젝트 폴더 전체를 대상 PC로 복사
2. Python 3.13 설치 확인 (`py -3.13 --version`)
3. `install_offline.bat` 더블클릭 (가상환경 생성 + wheelhouse에서 오프라인 설치)

수동 설치 시:

```bat
py -3.13 -m venv .venv
.venv\Scripts\python -m pip install --no-index --find-links wheelhouse -r requirements.txt
```

### wheelhouse 재생성 (인터넷 되는 PC에서)

의존성을 갱신하고 싶을 때만 필요합니다.

```bat
py -3.13 -m pip download -r requirements.txt -d wheelhouse --only-binary :all:
```

## 사용법 1: GUI (비개발자 권장)

```bat
run_gui.bat
```

더블클릭하면 창이 뜹니다. 사용 순서:

1. **파일 선택** — 추출할 .hwp/.hwpx 파일(또는 폴더) 선택
2. **모든 표 추출** 선택 후 [추출 실행] — 결과 엑셀에서 원하는 표의 1행 헤더 확인
3. 규칙 표에 **규칙 추가** — 헤더 칸에 `품명, 규격, 수량` 처럼 쉼표로 입력 → [규칙 저장]
4. **규칙에 맞는 표만 추출** 선택 후 [추출 실행]
5. [결과 폴더 열기]로 엑셀 확인

규칙 저장 시 `config/target_tables.yaml`이 자동으로 갱신되므로 yaml을 직접 편집할 필요가 없습니다.

## 사용법 2: 명령어 (CLI)

```bat
REM 파일 하나 처리
run.bat data\교범.hwpx

REM 폴더 내 모든 hwp/hwpx 일괄 처리
run.bat data\

REM 규칙 무시하고 모든 표 추출 (규칙 튜닝 전 문서 구조 파악용)
run.bat data\교범.hwpx --all

REM 규칙 파일/출력 폴더 지정
run.bat data\교범.hwpx -c config\target_tables.yaml -o output\
```

또는 직접 실행:

```bat
.venv\Scripts\python -m src.main data\교범.hwpx
```

결과:

- `output\<파일명>_tables.xlsx` — 선별된 표 1개당 시트 1개 (`include_title: true`면 1행 제목·3행부터 표, `false`면 1행부터 표만)
- `output\extract.log` — 표별 매칭 점수와 근거 로그 (오탐/누락 검수용)

## 출력 옵션 (`config/target_tables.yaml`)

```yaml
output:
  include_title: false   # false: 엑셀에 표 본문만 (1행부터) / true: 1행 제목 + 3행부터 표
```

`--all` 실행 시에도 `output` 설정은 적용됩니다. `title_keywords`는 표 **선별**에만 쓰이며, `include_title: false`이면 엑셀에는 제목이 들어가지 않습니다.

## 표 선별 규칙 설정 (`config/target_tables.yaml`)

```yaml
rules:
  - name: "부품목록표"                                  # 시트명 접두어
    expected_headers: ["품명", "규격", "수량", "비고"]    # 1행 헤더 (일치율 x 50점)
    title_keywords: ["부품 목록", "수리부속"]             # 캡션/직전 문단 키워드 (+30점)
    expected_cols: 4                                    # 열 개수 일치 (+10점)
    cell_pattern: "\\d{4}-\\d{2}-\\d{3}"                # 셀 내 정규식 (+10점)
    threshold: 60                                       # 이 점수 이상이면 선정
```

- 헤더/키워드 비교 시 공백은 무시됩니다 ("품 명" = "품명").
- 한 표가 여러 규칙에 걸리면 점수가 가장 높은 규칙 하나로 선정됩니다.
- 새 문서 유형은 규칙을 추가하기만 하면 되고 코드 수정이 필요 없습니다.

### 규칙 튜닝 순서 (권장)

1. `run.bat 샘플.hwpx --all` 로 문서의 모든 표를 추출해 구조 파악
2. 대상 표의 헤더/제목을 보고 규칙 작성
3. `run.bat 샘플.hwpx` 실행 후 `output\extract.log` 의 점수를 보고 threshold 조정

## .hwp / DRM 문서 처리 방식

입력 파일 형식에 따라 자동으로 처리 경로가 나뉩니다.

| 입력 | 처리 방식 | 한글 설치 |
|------|-----------|-----------|
| 정상 `.hwpx` (ZIP) | XML 직접 파싱 | 불필요 |
| `.hwp` 또는 DRM 암호화 파일 | **한글 COM으로 문서를 열어 표를 읽음** | **전체 한글 필수** |

DRM 문서 처리 원리: 파일을 복호화하는 것이 아니라, 열람 권한이 있는 PC에서
한글이 문서를 열면(=메모리에서 복호화됨) 그 상태에서 표마다 HWPML(XML) 블록을
내보내 파싱합니다. 따라서 **해당 PC에서 그 문서가 한글로 정상 열려야** 동작합니다.

주의 사항:

- **한컴오피스 뷰어(hwpviewer)만 설치된 PC에서는 동작하지 않습니다.** 편집 가능한
  전체 한글(Hwp.exe)이 설치되어 있어야 COM 자동화 개체가 존재합니다.
- 무인 실행하려면 한글 **보안 승인 모듈(FilePathCheckerModule)** 을 레지스트리에
  등록해야 파일 열 때 승인 팝업이 뜨지 않습니다.
- DRM 정책이 "내용 내보내기"까지 차단하는 경우 표를 읽지 못할 수 있으며,
  이때는 명확한 오류 메시지가 출력됩니다.
- 실행 중 한글 프로세스가 백그라운드로 뜨며, 완료 후 자동 종료됩니다.

## 프로젝트 구조

```
├── config/target_tables.yaml   # 표 선별 규칙
├── src/
│   ├── gui.py                  # PyQt6 GUI (파일 선택 + 규칙 편집 + 실행)
│   ├── main.py                 # CLI 진입점 (ZIP이면 XML 파싱, 아니면 COM 리더)
│   ├── hwpx_parser.py          # HWPX XML → 표 모델 (병합/크기/캡션 포함)
│   ├── hwp_com_reader.py       # DRM/구형 HWP → 한글 COM으로 열어 표 읽기
│   ├── table_matcher.py        # 점수 기반 표 선별
│   ├── excel_writer.py         # openpyxl 저장 (병합/열너비/행높이 유지)
│   ├── convert.py              # .hwp → .hwpx 변환 (참고용, 기본 흐름에서는 미사용)
│   └── models.py               # Cell / Table / MatchResult
├── tools/
│   ├── make_sample_hwpx.py     # 테스트용 샘플 hwpx 생성기
│   ├── make_test_hwp.py        # COM 검증용 테스트 hwp 생성기 (한글 필요)
│   └── test_hml_parser.py      # COM 리더의 XML 파서 단위 테스트 (한글 불필요)
├── wheelhouse/                 # 오프라인 설치용 wheel (py3.13 win64)
├── install_offline.bat         # 오프라인 설치 스크립트
├── run_gui.bat                 # GUI 실행 (비개발자용)
└── run.bat                     # CLI 실행 스크립트
```
