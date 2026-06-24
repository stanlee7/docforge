# DocForge — 엑셀 행 → 개인별 AI 문서 일괄 생성

엑셀 명단의 **각 행을 한 편의 문서로** 자동 생성하는 온디바이스 무료 도구입니다.
이름·직무·성과 같은 구조화 데이터를 넣으면, 로컬 LLM(Ollama + Gemma)이 **자연어 문서**(성과평가서 등)를 행마다 작성해 `.docx`로 떨어뜨립니다. 단순 `{{치환}}` 머지가 아니라 **데이터→산문 작문**이며, 인터넷 없이 완전 로컬에서 돌아갑니다(개인정보 외부 전송 0).

> 자매 도구 **DocBatch**(문서 일괄 요약)의 온디바이스 엔진을 그대로 재사용하며, 모델도 공유합니다.

## 동작 방식
1. 엑셀(첫 행=헤더, 예: `이름, 직무, 성과, 평가지표`)을 선택
2. 문서 유형(프리셋) · AI에 넣을 컬럼 · 작성 지침 · 파일명 패턴(`{이름}_성과평가서`)을 지정
3. 행마다 `Gemma`가 문서를 작성 → `출력폴더/{이름}_성과평가서.docx` 생성 + `_생성결과.xlsx` 요약

## 요구 환경 (Windows)
- Python 3.10+ / Ollama (첫 실행 시 자동 설치·모델 다운로드 안내)
- 모델: `gemma4:26b-a4b-it-qat`(기본·품질, MoE 활성 4B / **RAM 16GB+ 권장**) / `gemma3:4b-it-qat`(빠른·저사양 모드, ~4GB·8GB램 PC)
- 모델 경로는 한글 계정 버그 회피를 위해 ASCII 경로 `C:\ProgramData\LocalDocAI\models` 고정

## 설치 & 실행
```powershell
pip install -r requirements.txt

# GUI
python app.py

# CLI
python make_sample.py                       # 검증용 샘플 명단 생성
python cli.py samples\sample.xlsx --preset 성과평가서 --out-dir samples\out --fmt docx [--fast]
```

## 빌드 (단일 exe)
```powershell
powershell -ExecutionPolicy Bypass -File build.ps1   # -> dist\DocForge.exe
```

## 구조
| 파일 | 역할 |
|------|------|
| `docforge/config.py` | 모델·프리셋(`PRESETS`)·출력형식·생성옵션 |
| `docforge/setup_env.py` | Ollama 설치·ASCII 경로 고정·서버 재기동·모델 pull (DocBatch 재사용) |
| `docforge/engine.py` | 서버 보장·keep_warm·`build_prompt`/`generate` |
| `docforge/excel_io.py` | 엑셀 읽기·`{{컬럼}}` 머지·파일명 채우기 (mail-batch 발췌) |
| `docforge/writers.py` | docx/txt/md 라이터 + 마크다운 평문화 |
| `docforge/pipeline.py` | 행 순회·진행/취소 콜백·결과 xlsx |
| `docforge/gui.py` | Tkinter 메인 폼 + 온보딩 다이얼로그 |
| `app.py` / `cli.py` | 진입점(GUI / CLI) |

## 출력 형식
`docx`(기본) · `txt` · `md`. (HWP는 미포함 — pyhwpx/COM 의존·한글 설치 필요로 v1 제외.)
