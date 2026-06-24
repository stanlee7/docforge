"""첫 실행 온보딩: Ollama 설치 확인 → ASCII 경로 강제 → 서버 재기동 → 모델 pull.

실전 함정: Ollama 설치 시 트레이 앱이 기본(한글 계정) 경로로 서버를 자동 기동한다.
그 서버가 떠 있으면 우리가 둔 ASCII 경로 모델을 못 찾고, pull도 한글 경로로 간다.
따라서 env를 영구 설정한 뒤 서버를 ASCII 경로로 '재기동'해야 일관성이 유지된다.

(DocBatch의 검증된 온보딩을 그대로 재사용 — 모델 경로를 공유하므로 이미 받은 모델은 재다운로드하지 않는다.)
"""
import json
import os
import shutil
import subprocess
import time

import requests

from . import config

OLLAMA_DOWNLOAD_URL = "https://ollama.com/download"
_NOWIN = getattr(subprocess, "CREATE_NO_WINDOW", 0)


# ---------- Ollama 설치 ----------
def is_installed() -> bool:
    return os.path.exists(config.ollama_exe()) or shutil.which("ollama") is not None


def winget_available() -> bool:
    return shutil.which("winget") is not None


def install_ollama_winget() -> tuple:
    """winget로 Ollama 무인 설치. (성공여부, 메시지)."""
    if not winget_available():
        return False, "winget 없음 — 수동 설치 필요"
    try:
        p = subprocess.run(
            ["winget", "install", "--id", "Ollama.Ollama", "-e", "--silent",
             "--accept-package-agreements", "--accept-source-agreements"],
            capture_output=True, text=True, timeout=600, creationflags=_NOWIN)
        if p.returncode == 0:
            return True, "Ollama 설치 완료"
        return False, f"설치 실패(코드 {p.returncode})"
    except Exception as e:  # noqa: BLE001
        return False, f"설치 오류: {e}"


# ---------- ASCII 모델 경로 ----------
def ensure_ascii_models_env() -> str:
    """모델 경로를 ASCII로 고정(현재 세션 + 영구). 경로 반환."""
    path = config.ASCII_MODELS_DIR
    os.makedirs(path, exist_ok=True)
    os.environ["OLLAMA_MODELS"] = path
    try:  # 다음 로그인부터 트레이 앱도 이 경로를 쓰도록 영구 저장
        subprocess.run(["setx", "OLLAMA_MODELS", path],
                       capture_output=True, timeout=15, creationflags=_NOWIN)
    except Exception:  # noqa: BLE001
        pass
    return path


# ---------- 서버 ----------
def server_up() -> bool:
    try:
        requests.get(config.OLLAMA_URL, timeout=2)
        return True
    except requests.RequestException:
        return False


def restart_server_ascii() -> bool:
    """기존 ollama 프로세스를 종료하고 ASCII 경로로 서버를 재기동."""
    ensure_ascii_models_env()
    subprocess.run(["taskkill", "/F", "/IM", "ollama.exe"],
                   capture_output=True, creationflags=_NOWIN)
    subprocess.run(["taskkill", "/F", "/IM", "ollama app.exe"],
                   capture_output=True, creationflags=_NOWIN)
    time.sleep(1.5)
    env = dict(os.environ, OLLAMA_MODELS=config.ASCII_MODELS_DIR)
    subprocess.Popen([config.ollama_exe(), "serve"], env=env,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                     creationflags=_NOWIN)
    for _ in range(30):
        if server_up():
            return True
        time.sleep(1)
    return False


# ---------- 모델 ----------
def list_models() -> set:
    try:
        r = requests.get(f"{config.OLLAMA_URL}/api/tags", timeout=5)
        return {m["name"] for m in r.json().get("models", [])}
    except requests.RequestException:
        return set()


def has_model(name: str) -> bool:
    return name in list_models()


def pull_model(name: str, progress=None) -> bool:
    """모델 스트리밍 다운로드. progress(percent:int, status:str) 콜백."""
    try:
        with requests.post(f"{config.OLLAMA_URL}/api/pull",
                           json={"model": name, "stream": True},
                           stream=True, timeout=None) as r:
            r.raise_for_status()
            for line in r.iter_lines():
                if not line:
                    continue
                data = json.loads(line.decode("utf-8"))
                if data.get("error"):
                    if progress:
                        progress(-1, data["error"])
                    return False
                status = data.get("status", "")
                total, done = data.get("total"), data.get("completed")
                pct = int(done / total * 100) if total and done else 0
                if progress:
                    progress(pct, status)
            return has_model(name)
    except Exception as e:  # noqa: BLE001
        if progress:
            progress(-1, str(e))
        return False


# ---------- 종합 ----------
def check_readiness() -> dict:
    installed = is_installed()
    up = server_up() if installed else False
    model = has_model(config.MODEL_FULL) if up else False
    return {"installed": installed, "server": up, "model": model,
            "ready": installed and up and model}
