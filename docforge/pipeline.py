"""엑셀 행 일괄 처리 → 행별 문서 생성 + 결과 엑셀 출력.

DocBatch pipeline.py의 골격(progress 콜백 + 한 건 실패 격리 + 결과 xlsx)을
'폴더 순회'에서 '엑셀 행 순회'로 바꿔 재사용한다.
"""
import os
import time

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

from . import config, engine, excel_io, writers


def _name_of(row: dict, columns) -> str:
    """행을 대표하는 이름(파일명/표시용). '이름' 컬럼 우선, 없으면 첫 값."""
    if row.get("이름"):
        return row["이름"]
    for c in (columns or row.keys()):
        if row.get(c):
            return row[c]
    return ""


def generate_documents(excel_path: str, out_dir: str, preset_key: str = None,
                       user_instructions: str = "", columns=None,
                       filename_pattern: str = None, fmt: str = None,
                       model: str = None, progress=None, cancel=None,
                       hybrid_template: str = None):
    """행마다 AI 문서를 생성해 out_dir에 기록. 결과 행 리스트 반환.

    progress(done, total, name) 콜백, cancel() -> True면 중단.
    한 행 실패는 기록만 하고 계속 진행한다(배치 전체 중단 방지).
    """
    preset_key = preset_key or config.DEFAULT_PRESET
    preset = config.PRESETS.get(preset_key) or config.PRESETS[config.DEFAULT_PRESET]
    filename_pattern = filename_pattern or preset["default_filename"]
    fmt = fmt or config.DEFAULT_FORMAT
    model = model or config.MODEL_FULL

    engine.ensure_server()
    engine.keep_warm(model)

    _, rows = excel_io.read_excel(excel_path)
    os.makedirs(out_dir, exist_ok=True)

    results = []
    total = len(rows)
    for i, row in enumerate(rows):
        if cancel and cancel():
            break
        name = _name_of(row, columns)
        if progress:
            progress(i, total, name)
        out = {"행": i + 1, "이름": name, "출력파일": "", "글자수": 0, "상태": "", "소요(초)": 0}
        t0 = time.time()
        try:
            messages = engine.build_messages(preset_key, user_instructions, row,
                                             columns=columns, hybrid_template=hybrid_template)
            text = engine.generate(messages, model)
            out_path = writers.build_out_path(out_dir, filename_pattern, row, fmt)
            title = f"{name} {preset['label']}".strip()
            writers.write_document(text, out_path, fmt, title=title)
            out.update({"출력파일": os.path.basename(out_path),
                        "글자수": len(text), "상태": "완료"})
        except Exception as e:  # noqa: BLE001 - 한 건 실패가 배치 전체를 막지 않게
            out["상태"] = f"오류: {e}"
        out["소요(초)"] = round(time.time() - t0, 1)
        results.append(out)
    if progress:
        progress(total, total, "완료")
    return results


COLUMNS = ["행", "이름", "출력파일", "글자수", "상태", "소요(초)"]
WIDTHS = [6, 18, 34, 10, 30, 10]


def write_result_excel(rows, out_path: str):
    """생성 결과 요약을 xlsx로 저장(DocBatch write_excel 패턴)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "생성결과"
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")
    for r in rows:
        ws.append([r.get(c, "") for c in COLUMNS])
    for idx, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[chr(64 + idx)].width = w
    ws.freeze_panes = "A2"
    wb.save(out_path)
    return out_path
