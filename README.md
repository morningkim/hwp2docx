# hwp2docx

한글 ↔ 워드 문서를 서로 변환하는 순수 Python 도구입니다.
한컴오피스 · MS오피스 · LibreOffice **설치 없이** 동작합니다.

- **HWP → DOCX** : HWP 5.0 바이너리(OLE 복합 파일)를 직접 파싱해 본문·표를 추출하여 DOCX 생성
- **DOCX → HWPX** : DOCX를 읽어 한글 OWPML(HWPX, XML+ZIP) 패키지로 생성

> Pure-Python converter between Hangul Word Processor and Word formats,
> with no office suite (Hancom / MS Office / LibreOffice) required.

## 특징

- **무설치 변환** — 오피스 프로그램 불필요
- 본문 문단 + 표(행·열, 셀 병합 `colSpan`/`rowSpan`) 처리
- 간단한 GUI 제공(파이썬 기본 내장 tkinter) — 한 창에서 양방향 선택
- CLI 도 제공

## 설치

```bash
pip install olefile python-docx
```

## 사용법

### GUI (Windows) — 가장 간단

- `한글-워드 변환기.bat` 더블클릭 → 창에서 **변환 방향**을 고르고 파일/폴더 선택
- 또는: `python hwp2docx_gui.py`

### 명령줄 (CLI)

```bash
# HWP -> DOCX
python hwp2docx.py                 # 현재 폴더의 모든 *.hwp
python hwp2docx.py a.hwp           # a.hwp -> a.docx
python hwp2docx.py a.hwp out.docx

# DOCX -> HWPX
python docx2hwpx.py                # 현재 폴더의 모든 *.docx
python docx2hwpx.py a.docx         # a.docx -> a.hwpx
python docx2hwpx.py a.docx out.hwpx
```

변환 결과는 기본적으로 원본과 같은 폴더에 저장됩니다.

## 구성

| 파일 | 설명 |
|---|---|
| `hwp2docx.py` | HWP → DOCX 변환 엔진 (CLI) |
| `docx2hwpx.py` | DOCX → HWPX 변환 엔진 (CLI) |
| `hwp2docx_gui.py` | 양방향 GUI |
| `한글-워드 변환기.bat` | GUI 실행기(더블클릭) |
| `hwpx_assets/header.xml` | HWPX 생성용 서식 정의 템플릿(한글 기본 스타일, 본문·개인정보 없음) |

## 한계

- 글꼴 · 글자 크기 · 색상 · 정밀 레이아웃 · 이미지 · 도형은 재현하지 않습니다(**텍스트와 표 구조 위주**).
- HWP → DOCX 는 HWP **5.0(.hwp) 바이너리** 대상이며, 암호 보호 문서는 지원하지 않습니다.
- DOCX → HWPX 는 한글 OWPML 패키지를 생성하며, 레이아웃은 한글에서 열 때 재계산됩니다.
  **생성한 .hwpx 는 한글(또는 한글 뷰어)에서 열어 확인하길 권장합니다.**
  필요하면 한글에서 "다른 이름으로 저장"으로 .hwp(바이너리)로 변환할 수 있습니다.
- 진짜 `.hwp` **바이너리 직접 생성**은 포맷이 비공개·복잡하여 지원하지 않습니다(대신 .hwpx 생성).

## 라이선스

MIT License. 자세한 내용은 [LICENSE](LICENSE) 참고.
