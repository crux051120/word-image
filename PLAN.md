# PLAN.md — word_image

## 목표

docx 의 2행 1열 표 첫 행에 적힌 이름으로 `image/` 의 이미지를 찾아 두 번째 행에 삽입하는 GUI 도구(`start.exe`).

## 성공 기준

- `demo.docx`(표 10개)에 대해 `대상 표 10 · 삽입 10 · 누락 0`, `demo_filled.docx` 에 이미지 10개
- 누락·비대상 표·중복 확장자·0건 케이스가 테스트로 고정됨
- `start.exe` 더블클릭으로 창이 뜨고 같은 결과를 냄

## 확정 사항 (2026-10-01 grilling)

| 항목 | 결정 |
|---|---|
| 라이브러리 | python-docx, tkinter/ttk, PyInstaller onefile |
| 대상 표 | 정확히 2행 1열 |
| 매칭 | 공백 제거, 대소문자 무시, `.png .jpg .jpeg .gif .bmp` 순, 중복 시 경고 후 첫 번째 |
| 크기 | 셀 너비에 맞춤(비율 유지), 너비 없으면 본문 폭 |
| 기존 내용 | 두 번째 행 지우고 교체 |
| 누락 | 건너뛰고 경고, 끝에 건수 요약. 대상 0 또는 삽입 0 → 실패 |
| 출력 | 원본 보존, `<이름>_filled.docx`, 존재 시 덮어쓰기 확인 |
| 이미지 지정 | 폴더(기본 exe 옆 `image`) / 파일 다중 선택 |
| 실행 | 작업 스레드, 실행 중 버튼 비활성 |
| 데모 | `make_demo.py` 개발용. demo_01~10 (PNG 7 + JPG 3, 600×400) |

## 단계

| # | 내용 | 검증 |
|---|---|---|
| 1 | git init, 브랜치, .gitignore | git status |
| 2 | AGENTS/CLAUDE/PLAN/DESIGN.md | 파일 확인 |
| 3 | 의존성 설치 | import |
| 4 | insert_images.py, make_demo.py | 이미지 10 · 표 10 |
| 5 | pytest | 통과 |
| 6 | app.py GUI | 창 띄워 데모 실행 |
| 7 | start.exe 빌드 | exe 로 demo_filled.docx 이미지 10개 |

## 실행 기록

### 2026-10-01

- 1~3: git init, `claude/word-image-insert` 브랜치, python-docx 1.2.0 / pyinstaller 6.22.3 설치
- 4: `make_demo.py` → image 10개(PNG 7, JPG 3), demo.docx 표 10개 확인
- 5: pytest 7 passed (정상·누락·비대상 표·중복 확장자·삽입 0건·대상 0개·출력=입력 거부)
- 6: app.py 를 스크립트로 구동 — 폴더 모드 `대상 표 10 · 삽입 10 · 누락 0`, demo_filled.docx 그림 10개 /
  파일 2개 선택 모드 `대상 표 10 · 삽입 2 · 누락 8`
- 7: start.exe 16.9MB 빌드, 실행 시 창 "워드 이미지 자동 삽입" 응답 정상 확인
- 미검증: exe 화면에서 사람이 직접 [찾기]·[실행] 클릭하는 흐름(파일 대화상자)은 자동 확인하지 못함

### 2026-10-01 — 실행 결과 화면 개편 (grilling 2차)

확정: 이미지 수=허용 확장자 파일 수 · 표 수=대상 표만 · 확장자 중복=정상 · 표 이름 중복=예외(어느 표에도 미삽입) ·
예외 1건 이상=실패·미저장 + [예외 무시하고 저장](재처리, 정상 표만 채움, 예외 표 유지) · 필터 라디오 전체/예외만

- pytest 12 passed (이미지만 있음 정렬·표 이름 중복·빈 이름 다수·예외 무시 저장 시 예외 표 유지·정상 0건 저장 불가 추가)
- GUI 스크립트 구동 (demo/ 기준, 결과는 임시 폴더):
  1) 정상 `이미지 10 · 대상 표 10 · 정상 10 · 예외 0` 저장
  2) 임시 demo_11.png 추가 → `이미지 11 · 예외 1(이미지만 있음 1)` 실패·미저장, 예외만 필터 1행, [예외 무시하고 저장] 활성
  3) 예외 무시 저장 → 그림 10개 저장, 필터 선택 유지. demo_11.png 삭제 확인
- 작업 중 사용자가 데모를 `demo/`(image, demo.docx) 로 옮김 → GUI 기본 이미지 폴더와 make_demo.py 출력을 `demo/` 기준으로 변경, 빌드 산출물은 %TEMP% 로 빼고 build/·dist/·start.spec 삭제. pytest 12 passed, exe 창 정상
