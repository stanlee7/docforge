"""Ollama 추론 엔진 래퍼: 서버 보장 + keep-warm + 행 데이터 → 문서 생성.

DocBatch engine.py에서 서버/keep_warm 패턴을 그대로 가져오고, '요약·추출'을
'문서 작문(generate)'으로 교체했다.
"""
import os
import re
import subprocess
import time

import requests

from . import config, excel_io

_FENCE_RE = re.compile(r"^```[a-zA-Z]*\n?|\n?```$")
# 약한 모델이 본문 앞에 붙이는 대화체 군말(예: "물론입니다.", "다음은 …평가서입니다.")
# — 첫 문장 하나만 보고, 군말 키워드로 시작하는 짧은 도입부일 때만 제거(본문 침범 방지).
_PREAMBLE_OPENER = re.compile(
    r"^\s*(물론입니다|네[,.]|알겠습니다|다음은|아래는|작성해\s?드리겠|기꺼이)"
    r"[^\n。.!?]*[。.!?:：]?\s*"
)
# 약한 모델이 프롬프트 골격을 그대로 되뱉는 경우 제거할 섹션 머리표(정규화 비교).
_SCAFFOLD_LABELS = {"작성지침", "대상데이터", "기준골격"}
_SCAFFOLD_TAIL = "위 데이터만 근거로"
_NORM_RE = re.compile(r"[\s\*#:：\[\]\(\)\-_.]+")


def _is_scaffold_line(line: str) -> bool:
    """마크다운/괄호/콜론을 벗겨낸 짧은 라벨 줄이 골격 머리표면 True."""
    norm = _NORM_RE.sub("", line)
    return norm in _SCAFFOLD_LABELS


# 한 줄 통째로 대화체 군말인 경우(문서 본문일 수 없음) 어디에 있든 제거.
_FILLER_LINES = {
    "물론입니다", "네", "알겠습니다", "기꺼이도와드리겠습니다",
    "작성해드리겠습니다", "다음과같습니다", "아래와같습니다",
}


def _is_filler_line(line: str) -> bool:
    norm = _NORM_RE.sub("", line)
    return 0 < len(norm) <= 14 and norm in _FILLER_LINES


def _server_up() -> bool:
    try:
        requests.get(config.OLLAMA_URL, timeout=2)
        return True
    except requests.RequestException:
        return False


def ensure_server() -> None:
    """서버가 없으면 ASCII 모델 경로로 띄운다."""
    if _server_up():
        return
    os.makedirs(config.ASCII_MODELS_DIR, exist_ok=True)
    env = dict(os.environ, OLLAMA_MODELS=config.ASCII_MODELS_DIR)
    subprocess.Popen(
        [config.ollama_exe(), "serve"],
        env=env,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    for _ in range(30):
        if _server_up():
            return
        time.sleep(1)
    raise RuntimeError("Ollama 서버를 시작하지 못했습니다.")


def keep_warm(model: str) -> None:
    """모델을 메모리에 상주시켜 콜드스타트(첫 건 수십 초)를 제거한다."""
    requests.post(
        f"{config.OLLAMA_URL}/api/generate",
        json={"model": model, "prompt": "안녕", "stream": False,
              "keep_alive": "30m", "options": {"num_predict": 1}},
        timeout=config.REQUEST_TIMEOUT,
    )


def _serialize_row(row: dict, columns) -> str:
    """선택된 컬럼만 '- 필드: 값' 줄로 직렬화. 빈 값은 생략."""
    cols = columns or list(row.keys())
    lines = []
    for c in cols:
        v = (row.get(c, "") or "").strip()
        if v:
            lines.append(f"- {c}: {v}")
    return "\n".join(lines)


def build_messages(preset_key: str, user_instructions: str, row: dict,
                   columns=None, hybrid_template: str = None) -> list:
    """Gemma instruct에 맞춘 chat 메시지(system/user) 구성.

    역할을 분리하면 약한 모델이 지시문을 본문에 되뱉는 현상이 크게 줄어든다.
    """
    preset = config.PRESETS.get(preset_key) or config.PRESETS[config.DEFAULT_PRESET]

    user_parts = []
    instr = (user_instructions or preset.get("default_instructions", "")).strip()
    if instr:
        user_parts.append(f"[작성 지침]\n{instr}")
    user_parts.append("[대상 데이터]\n" + (_serialize_row(row, columns) or "(데이터 없음)"))
    if hybrid_template and hybrid_template.strip():
        filled = excel_io.render_template(hybrid_template, row)
        user_parts.append(
            "[기준 골격] — 아래 틀과 핵심 문구를 유지하되 자연스러운 산문으로 보강하세요.\n" + filled)
    user_parts.append(
        "위 데이터만 근거로 한국어로 완성된 한 편의 문서 본문을 작성하세요. "
        "위 라벨([작성 지침] 등)이나 데이터 목록을 그대로 출력하지 말고, 문서 본문만 작성하세요.")

    return [
        {"role": "system", "content": preset["system"].strip()},
        {"role": "user", "content": "\n\n".join(user_parts)},
    ]


def _clean(text: str) -> str:
    out = (text or "").strip()
    out = _FENCE_RE.sub("", out).strip()
    # 프롬프트 골격이 새어나온 경우: 머리표 줄 제거 + 마지막 지시문단 이후만 채택
    lines = [ln for ln in out.splitlines()
             if not _is_scaffold_line(ln) and not _is_filler_line(ln)]
    out = "\n".join(lines)
    if _SCAFFOLD_TAIL in out:
        out = out.split(_SCAFFOLD_TAIL, 1)[-1]
        out = out.split("\n", 1)[-1] if "\n" in out else ""
    out = out.strip()
    # 선두 대화체 도입부를 최대 2개까지 제거(짧은 도입부에 한함)
    for _ in range(2):
        m = _PREAMBLE_OPENER.match(out)
        if not m or m.end() == 0 or m.end() > 60:
            break
        out = out[m.end():].lstrip()
    return out.strip()


def generate(messages, model: str = None) -> str:
    """chat 메시지 리스트 → 생성 문서 본문(문자열). 빈 응답이면 예외.

    하위호환: messages에 문자열을 주면 단일 user 메시지로 감싼다.
    """
    model = model or config.MODEL_FULL
    if isinstance(messages, str):
        messages = [{"role": "user", "content": messages}]
    resp = requests.post(
        f"{config.OLLAMA_URL}/api/chat",
        json={"model": model, "messages": messages, "stream": False,
              "keep_alive": "30m", "options": config.GEN_OPTIONS},
        timeout=config.REQUEST_TIMEOUT,
    )
    resp.raise_for_status()
    raw = resp.json().get("message", {}).get("content", "")
    out = _clean(raw)
    if not out:
        raise ValueError("빈 응답(모델이 본문을 생성하지 못함)")
    return out
