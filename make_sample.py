"""검증용 샘플 명단 엑셀 생성: samples/sample.xlsx"""
import os

from openpyxl import Workbook

SAMPLE = [
    ["이름", "직무", "성과", "평가지표"],
    ["홍길동", "백엔드 개발", "결제 모듈 리팩터링으로 장애율 40% 감소", "목표달성 110%, 협업 우수"],
    ["김영희", "마케팅", "신규 캠페인으로 가입 전환율 2배", "리드 1.8만, ROAS 320%"],
    ["박철수", "영업", "신규 거래처 7곳 확보, 매출 1.2억 증가", "목표달성 95%, 고객만족 4.6/5"],
    ["", "기타", "이름 누락 행(실패 격리 테스트용)", "-"],
]


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(here, "samples")
    os.makedirs(out_dir, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "명단"
    for row in SAMPLE:
        ws.append(row)
    path = os.path.join(out_dir, "sample.xlsx")
    wb.save(path)
    print(f"샘플 생성 -> {path}")


if __name__ == "__main__":
    main()
