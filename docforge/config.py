"""DocForge - 설정 한 곳에 모음."""
import os

# --- 모델 ---
# 품질 기본값: Gemma 4 26B A4B (MoE, 활성 4B). 추론은 4B급 속도지만
# 메모리는 26B 전체를 적재 → RAM 16GB+ 권장. 한국어 작문 품질이 대폭 향상.
MODEL_FULL = "gemma4:26b-a4b-it-qat"
# 빠른/저사양 모드: 메모리 ~4GB로 8GB램 PC 안전 폴백 겸 속도용.
MODEL_FAST = "gemma3:4b-it-qat"

OLLAMA_URL = "http://localhost:11434"

# --- 모델 저장 경로 (한글 계정명 회피: 반드시 ASCII 경로) ---
# 검증에서 확인된 버그: C:\Users\<한글>\.ollama 경로면 llama-server 로딩 실패.
# DocBatch와 동일 경로를 공유하여 모델 재다운로드를 피한다.
ASCII_MODELS_DIR = r"C:\ProgramData\LocalDocAI\models"

# --- 출력 포맷 ---
OUTPUT_FORMATS = ["docx", "txt", "md"]
DEFAULT_FORMAT = "docx"

# --- 생성 옵션 (요약/추출이 아니라 '작문'이므로 온도↑, 길이↑) ---
GEN_OPTIONS = {"temperature": 0.7, "num_ctx": 8192, "num_predict": 1536}
REQUEST_TIMEOUT = 600

# --- 문서 유형 프리셋 (v1: 성과평가서 1종, 추후 추가) ---
# 각 프리셋: 라벨 + system(역할/톤/형식) + 기본 사용자 지침 + 기본 파일명 패턴.
PRESETS = {
    "성과평가서": {
        "label": "성과평가서",
        "system": (
            "당신은 한국 기업의 인사 평가 문서를 작성하는 전문가입니다.\n"
            "주어진 평가 데이터를 근거로 한 편의 완결된 성과평가서를 한국어로 작성하세요.\n"
            "- 존댓말, 객관적이고 구체적인 서술. 데이터에 근거한 사실만 쓰고 과장하지 마세요.\n"
            "- 강점과 보완점을 균형 있게 다루고, 마지막에 종합 의견을 제시하세요.\n"
            "- 머리말(예: '물론입니다', '다음은…'), 코드블록, 마크다운 표식 없이 문서 본문만 출력하세요."
        ),
        "default_instructions": "평가지표 데이터를 근거로 대상자의 강점과 보완점을 구체적인 사례로 서술하세요.",
        "default_filename": "{이름}_성과평가서",
    },
}

DEFAULT_PRESET = "성과평가서"


def ollama_exe() -> str:
    """설치된 ollama.exe 경로를 찾는다."""
    candidates = [
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Ollama\ollama.exe"),
        os.path.expandvars(r"%ProgramFiles%\Ollama\ollama.exe"),
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return "ollama"  # PATH에 있다고 가정
