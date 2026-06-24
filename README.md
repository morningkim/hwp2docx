# hwp2docx

한글(HWP 5.0) 문서를 워드(DOCX)로 변환하는 순수 Python 도구입니다.
한컴오피스 · MS오피스 · LibreOffice **설치 없이** 동작합니다.

HWP 5.0 바이너리(OLE 복합 파일)를 직접 파싱하여 본문 텍스트와 표(셀 병합 포함)를
추출하고 `.docx`로 재생성합니다.

> A pure-Python tool that converts Hangul Word Processor (HWP 5.0) files to DOCX,
> with no office suite (Hancom / MS Office / LibreOffice) required.

## 특징

- **무설치 변환** — 오피스 프로그램 불필요
- 여러 Section, 본문 문단, 표(행·열, `colSpan`/`rowSpan` 병합) 처리
- 한글 제어문자 폭 처리(char=1워드 / inline·extended=8워드)로 정확한 텍스트 추출
- 파일/버전별로 다른 셀 레벨 인코딩(LIST_HEADER의 자식 vs 형제) 모두 대응
- 간단한 GUI 제공(파이썬 기본 내장 tkinter, 추가 설치 없음)

## 설치

```bash
pip install olefile python-docx
```

## 사용법

### 명령줄 (CLI)

```bash
python hwp2docx.py                 # 현재 폴더의 모든 *.hwp 변환
python hwp2docx.py a.hwp           # a.hwp -> a.docx
python hwp2docx.py a.hwp out.docx  # 출력 경로 지정
```

### GUI (Windows)

- `HWP를 DOCX로 변환.bat` 더블클릭 → 창이 뜨면 파일/폴더 선택
- 또는: `python hwp2docx_gui.py`

변환된 `.docx`는 기본적으로 원본 HWP와 같은 폴더에 저장됩니다.

## 한계

- 글꼴 · 글자 크기 · 색상 · 정밀 레이아웃 · 이미지 · 도형은 재현하지 않습니다
  (**텍스트와 표 구조 위주**).
- 암호로 보호된 문서는 지원하지 않습니다.
- HWP 5.0(.hwp) 바이너리 형식을 대상으로 합니다. HWPX(.hwpx)는 대상이 아닙니다.

## 라이선스

MIT License. 자세한 내용은 [LICENSE](LICENSE) 참고.
