"""진입점.
- 인자 없이 실행: GUI
- 엑셀 인자 주면: CLI 모드 — 예) DocForge.exe 명단.xlsx --out-dir 출력 --fmt docx
"""
import sys


def main():
    if len(sys.argv) > 1:
        # windowed(=no console) 빌드에서 stdout/stderr가 None이면 print가 깨진다.
        import os
        if sys.stdout is None:
            sys.stdout = open(os.devnull, "w", encoding="utf-8")
        if sys.stderr is None:
            sys.stderr = sys.stdout
        from cli import main as cli_main
        cli_main()
    else:
        from docforge.gui import main as gui_main
        gui_main()


if __name__ == "__main__":
    main()
