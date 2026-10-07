# AGENTS.md — word_image

## 개요

docx 문서의 **2행 1열 표**를 찾아, 첫 행 텍스트(확장자 없는 이미지 이름)와 같은 이름의 이미지를
두 번째 행에 자동 삽입한다. 사용자는 `start.exe`(tkinter GUI)로 실행한다.

## 기술 스택

- Python 3.14, `python-docx`, `Pillow`(데모 이미지 생성), `tkinter`/`ttk`(GUI)
- 패키징: PyInstaller `--onefile --windowed`
- 테스트: pytest

## 명령어

```bash
python make_demo.py                 # demo/image/ 데모 이미지 10개 + demo/demo.docx 생성
python -m pytest -q                 # 테스트
python app.py                       # GUI 개발 실행
python -m PyInstaller --onefile --windowed --name start --distpath . --workpath %TEMP%/word_image_build --specpath %TEMP%/word_image_build app.py
# 빌드 산출물(build/·dist/·spec)을 프로젝트 밖에 두고 start.exe 만 루트에 만든다
```

## 규칙

- 핵심 로직은 `insert_images.py` 에만 둔다. GUI(`app.py`)는 호출과 표시만 한다.
- 대상 표 판정은 **정확히 2행 1열**. 다른 모양의 표는 건드리지 않는다.
- 이름 매칭: 앞뒤 공백 제거, 대소문자 무시. 확장자 우선순위 `.png .jpg .jpeg .gif .bmp`.
- 원본 docx 는 덮어쓰지 않는다. 기본 출력은 `<이름>_filled.docx`.
- 예외 3종: 표만 있음(이미지 없음·첫 행 빈 표) / 이미지만 있음(해당 표 없음) / 표 이름 중복(어느 표에도 넣지 않음).
  확장자 중복은 예외가 아니라 정상 + 안내 문구.
- 예외가 1건이라도 있으면 실패이며 저장하지 않는다. `allow_exceptions=True`(GUI [예외 무시하고 저장])일 때만
  정상 표만 채워 저장하고, 예외 표의 두 번째 행은 건드리지 않는다.

## 반복 오류 방어

- 대상 표 0개 또는 정상 0건은 예외 무시로도 저장하지 않는다 — `ok`·`can_save_ignoring_exceptions` 모두 False.
- 매 실행마다 `이미지 / 대상 표 / 정상 / 예외(3종)` 건수를 화면에 남긴다.
- exe 에서 기준 폴더는 `sys.executable` 위치다(`__file__` 은 onefile 임시 폴더를 가리킨다).
