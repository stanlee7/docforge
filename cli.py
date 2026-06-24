"""CLI 진입점 - 엑셀 행 → 문서 일괄 생성 테스트용.
사용: python cli.py <엑셀.xlsx> [--preset 성과평가서] [--out-dir 출력] [--fmt docx] [--fast]
"""
import argparse
import os
import sys

from docforge import config, pipeline


def main():
    ap = argparse.ArgumentParser(description="DocForge 엑셀 행 → AI 문서 일괄 생성")
    ap.add_argument("excel", help="명단 엑셀(.xlsx) 경로")
    ap.add_argument("--preset", default=config.DEFAULT_PRESET,
                    choices=list(config.PRESETS.keys()), help="문서 유형")
    ap.add_argument("--out-dir", default=None, help="출력 폴더(기본: 엑셀 옆 '생성문서')")
    ap.add_argument("--fmt", default=config.DEFAULT_FORMAT,
                    choices=config.OUTPUT_FORMATS, help="출력 형식")
    ap.add_argument("--instructions", default="", help="작성 지침(생략 시 프리셋 기본값)")
    ap.add_argument("--fast", action="store_true", help="1b 빠른 모드")
    args = ap.parse_args()

    if not os.path.isfile(args.excel):
        print(f"엑셀 없음: {args.excel}")
        sys.exit(1)

    out_dir = args.out_dir or os.path.join(os.path.dirname(os.path.abspath(args.excel)),
                                           "생성문서")
    model = config.MODEL_FAST if args.fast else config.MODEL_FULL

    def prog(done, total, name):
        print(f"[{done}/{total}] {name}")

    print(f"모델: {model} | 프리셋: {args.preset} | 형식: {args.fmt}")
    rows = pipeline.generate_documents(
        args.excel, out_dir, preset_key=args.preset,
        user_instructions=args.instructions, fmt=args.fmt, model=model, progress=prog)
    res = os.path.join(out_dir, "_생성결과.xlsx")
    pipeline.write_result_excel(rows, res)
    ok = sum(1 for r in rows if r["상태"] == "완료")
    print(f"\n완료: {ok}/{len(rows)}건 -> {out_dir}")
    for r in rows:
        print(f"  - 행{r['행']} {r['이름']}: {r['상태']} | {r['출력파일']} "
              f"({r['글자수']}자, {r['소요(초)']}s)")


if __name__ == "__main__":
    main()
