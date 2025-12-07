# totalS, totalZ 통합모듈

import openpyxl
from openpyxl.styles import Font, PatternFill
from decimal import Decimal, ROUND_HALF_UP
import numpy as np


# =========================
# 1. S 점수 (원본 totalS.py 동일)
# =========================
def calc_s(prices, window: int):
    arr = [p for p in prices if p is not None]
    if len(arr) < window:
        return None

    arr = arr[-window:]  # 최근 window일

    min_val = min(arr)
    max_val = max(arr)
    if max_val == min_val:
        return 0

    val = 100 * ((arr[-1] - min_val) / (max_val - min_val))

    score = int(Decimal(str(val)).to_integral_value(rounding=ROUND_HALF_UP))
    return score


# =========================
# 2. Z 점수 (원본 totalZ.py 동일)
# =========================
def calc_z(prices, window: int):
    arr = [p for p in prices if p is not None]
    if len(arr) < window:
        return None

    arr = arr[-window:]  # 최근 window일

    mean = float(np.mean(arr))
    std = float(np.std(arr, ddof=1))  # 표본 표준편차

    if std == 0:
        return 0

    z = (arr[-1] - mean) / std
    val = 50 * z  

    score = int(Decimal(str(val)).to_integral_value(rounding=ROUND_HALF_UP))
    return score


# =========================
# 3. 종가 시트를 읽어서 dates, stocks 반환
# =========================
def get_close_data(filename: str):
    wb = openpyxl.load_workbook(filename, data_only=True)

    if "종가" not in wb.sheetnames:
        raise ValueError(f"'{filename}' 파일에 '종가' 시트가 없습니다.")

    sheet = wb["종가"]

    # 날짜 가져오기
    dates = []
    for col in range(3, sheet.max_column + 1):
        v = sheet.cell(row=1, column=col).value
        if v is None:
            continue
        try:
            dates.append(int(v))
        except:
            continue

    # 종목 데이터
    stocks = []
    for row in range(2, sheet.max_row + 1):
        name = sheet.cell(row=row, column=1).value
        code = sheet.cell(row=row, column=2).value
        if not name or not code:
            continue

        prices = []
        for col in range(3, sheet.max_column + 1):
            v = sheet.cell(row=row, column=col).value
            if v is None or v == "":
                prices.append(None)
            else:
                try:
                    prices.append(float(v))
                except:
                    prices.append(None)

        stocks.append({
            "name": name,
            "code": str(code),
            "prices": prices,
        })

    return dates, stocks


# =========================
# 4. 기존 시트 삭제 후 새로 생성
# =========================
def create_or_clear_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    return wb.create_sheet(sheet_name)


# =========================
# 5. S/Z 단일 시트 저장 엔진
# =========================
def save_score_sheet(filename, dates, stocks, window, sheet_name, calc_func):
    wb = openpyxl.load_workbook(filename)

    if len(dates) < window:
        print(f"⚠ {sheet_name}: 날짜가 {window}일보다 적어 계산 불가.")
        return

    # window일 이후 날짜만 계산됨
    valid_dates = dates[window - 1:]

    sheet = create_or_clear_sheet(wb, sheet_name)

    # 헤더
    sheet.cell(row=1, column=1, value="종목명").font = Font(bold=True)
    sheet.cell(row=1, column=2, value="종목코드").font = Font(bold=True)

    sheet.cell(row=1, column=1).fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    sheet.cell(row=1, column=2).fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")

    for idx, d in enumerate(valid_dates, start=3):
        c = sheet.cell(row=1, column=idx)
        c.value = d
        c.font = Font(bold=True)
        c.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")

    # 데이터 쓰기
    for row_idx, stock in enumerate(stocks, start=2):
        sheet.cell(row=row_idx, column=1, value=stock["name"])
        sheet.cell(row=row_idx, column=2, value=stock["code"])

        prices = stock["prices"]

        for j, d in enumerate(valid_dates):
            full_index = dates.index(d)
            sub_prices = prices[:full_index + 1]

            score = calc_func(sub_prices, window)
            if score is not None:
                sheet.cell(row=row_idx, column=3 + j, value=score)

    # 열 너비
    sheet.column_dimensions["A"].width = 20
    sheet.column_dimensions["B"].width = 12
    for col in range(3, 3 + len(valid_dates)):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 10

    wb.save(filename)
    print(f"✅ {sheet_name} 저장 완료 ({filename})")


# =========================
# 6. S/Z 전체 수행
# =========================
def run_total_sz(filename):
    print(f"\n=== S/Z 통합 점수 계산 시작: {filename} ===")

    dates, stocks = get_close_data(filename)

    # ---- S 점수 ----
    for window, sheet in [(20, "s20"), (60, "s60"), (120, "s120")]:
        save_score_sheet(filename, dates, stocks, window, sheet, calc_s)

    # ---- Z 점수 ----
    for window, sheet in [(20, "z20"), (60, "z60"), (120, "z120")]:
        save_score_sheet(filename, dates, stocks, window, sheet, calc_z)

    print(f"=== S/Z 통합 점수 계산 완료 ===\n")


def main():
    run_total_sz("KR_Stocks_ETF.xlsx")


if __name__ == "__main__":
    main()
