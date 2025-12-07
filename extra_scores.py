# gap, quant, std 전체 계산 모듈
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import numpy as np
from decimal import Decimal, ROUND_HALF_UP


# =========================
# 1. 데이터 로드 함수
# =========================

def get_close_data(filename):
    """
    '종가' 시트에서 날짜와 종가 시계열을 읽어온다.
    - dates: ['YYYYMMDD', ...]
    - stocks: [{'name': 종목명, 'code': 종목코드, 'prices': [가격 또는 None, ...]}, ...]
    """
    dates = []
    stocks = []
    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb['종가']

        # 1행: 날짜 (3열부터 끝까지)
        raw_dates = [sheet.cell(row=1, column=col).value for col in range(3, sheet.max_column + 1)]
        for d in raw_dates:
            if d is None:
                continue
            d_str = str(d)
            if len(d_str) == 8 and d_str.isdigit():
                if d_str not in dates:
                    dates.append(d_str)

        # 2행 이후: 종목명, 종목코드, 종가들
        for row in range(2, sheet.max_row + 1):
            name = sheet.cell(row=row, column=1).value
            code = sheet.cell(row=row, column=2).value
            prices = [sheet.cell(row=row, column=col).value for col in range(3, sheet.max_column + 1)]
            prices_numeric = []
            for p in prices:
                try:
                    prices_numeric.append(int(p) if p not in (None, '') else None)
                except (ValueError, TypeError):
                    prices_numeric.append(None)
            stocks.append({'name': name, 'code': code, 'prices': prices_numeric})
    except Exception as e:
        print(f"⚠ 종가 시트 로딩 중 오류: {e}")
    return dates, stocks


def get_volume_data(filename):
    """
    '거래량' 시트에서 날짜와 거래량 시계열을 읽어온다.
    - dates: ['YYYYMMDD', ...]
    - stocks: [{'name': 종목명, 'code': 종목코드, 'volumes': [거래량 또는 None, ...]}, ...]
    """
    dates = []
    stocks = []
    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb['거래량']

        raw_dates = [sheet.cell(row=1, column=col).value for col in range(3, sheet.max_column + 1)]
        for d in raw_dates:
            if d is None:
                continue
            d_str = str(d)
            if len(d_str) == 8 and d_str.isdigit():
                if d_str not in dates:
                    dates.append(d_str)

        for row in range(2, sheet.max_row + 1):
            name = sheet.cell(row=row, column=1).value
            code = sheet.cell(row=row, column=2).value
            volumes = [sheet.cell(row=row, column=col).value for col in range(3, sheet.max_column + 1)]
            volumes_numeric = []
            for v in volumes:
                try:
                    volumes_numeric.append(int(v) if v not in (None, '') else None)
                except (ValueError, TypeError):
                    volumes_numeric.append(None)
            stocks.append({'name': name, 'code': code, 'volumes': volumes_numeric})
    except Exception as e:
        print(f"⚠ 거래량 시트 로딩 중 오류: {e}")
    return dates, stocks


# =========================
# 2. GAP / QUANT / STD 계산식
# =========================

def calc_gap(prices_window20):
    """
    GAP 점수: (오늘 종가 / 최근 20일 평균) * 100
    (_gap.py의 calc_gap과 동일)
    """
    arr = [p for p in prices_window20 if p is not None]
    if len(arr) < 20:
        return None
    mean = np.mean(arr)
    if mean == 0:
        return 0
    val = 100 * (arr[-1] / mean)
    score = int(Decimal(str(val)).to_integral_value(rounding=ROUND_HALF_UP))
    return score


def calc_quant(volumes_window60):
    """
    QUANT 점수: ((현재 거래량 / 최근 60일 평균) * 100) / 2
    (_quant.py의 calc_quant와 동일)
    """
    arr = [v for v in volumes_window60 if v is not None]
    if len(arr) < 60:
        return None
    mean = np.mean(arr)
    if mean == 0:
        return 0
    val = ((arr[-1] / mean) * 100) / 2
    score = int(Decimal(str(val)).to_integral_value(rounding=ROUND_HALF_UP))
    return score


def calc_std_value(prices, idx, window_std=20, window_mean=20):
    """
    STD 값 계산 (_std.py의 calc_std_value 그대로):

    - idx 시점에서 20일 롤링 표준편차 σ_t 계산
    - 과거 window_mean(20)일 동안 각 시점의 20일 롤링 σ의 평균(avg_std) 계산
    - STD = (σ_t / avg_std - 1) * 100
    - 소수 둘째 자리까지 반올림
    """
    min_idx = window_std + window_mean - 2  # 예: 20 + 20 - 2 = 38
    if idx < min_idx:
        return None

    std_list = []
    for j in range(idx - window_mean + 1, idx + 1):  # j: idx-19 ~ idx
        start = j - window_std + 1
        end = j + 1
        if start < 0:
            return None

        window_prices = prices[start:end]

        if any(p is None for p in window_prices):
            return None

        arr = np.array(window_prices, dtype=float)
        sigma = float(np.std(arr, ddof=0))  # 모표준편차
        std_list.append(sigma)

    if not std_list:
        return None

    std_today = std_list[-1]
    avg_std = sum(std_list) / len(std_list)
    if avg_std == 0:
        return 0

    raw_val = (std_today / avg_std - 1) * 100
    val = float(Decimal(str(raw_val)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    return val


# =========================
# 3. 공통 시트 생성 유틸
# =========================

def create_or_clear_sheet(wb, sheet_name):
    """sheet가 있으면 삭제 후 새로 생성"""
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    return wb.create_sheet(sheet_name)


# =========================
# 4. GAP 시트 전체 재계산 저장
# =========================

def save_gap_sheet(filename, dates, stocks, window=20, sheet_name='gap'):
    try:
        wb = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

    sheet = create_or_clear_sheet(wb, sheet_name)

    # 헤더
    sheet.cell(row=1, column=1, value='종목명')
    sheet.cell(row=1, column=2, value='종목코드')

    valid_dates = dates[window - 1:]
    for col, date in enumerate(valid_dates, start=3):
        cell = sheet.cell(row=1, column=col)
        cell.value = date
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

    sheet.cell(row=1, column=1).font = Font(bold=True)
    sheet.cell(row=1, column=1).fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
    sheet.cell(row=1, column=2).font = Font(bold=True)
    sheet.cell(row=1, column=2).fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

    # 데이터
    for row_idx, stock in enumerate(stocks, start=2):
        sheet.cell(row=row_idx, column=1, value=stock['name'])
        sheet.cell(row=row_idx, column=2, value=stock['code'])
        prices = stock['prices']

        for i in range(window - 1, len(prices)):
            window_prices = prices[i - window + 1:i + 1]
            if None in window_prices:
                val = None
            else:
                val = calc_gap(window_prices)
            col_idx = 3 + (i - (window - 1))
            sheet.cell(row=row_idx, column=col_idx, value=val)

    # 열 너비
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 12
    for col_idx in range(3, len(valid_dates) + 3):
        sheet.column_dimensions[get_column_letter(col_idx)].width = 12

    wb.save(filename)
    print(f"✅ GAP 전체 저장 완료: {filename} (탭: {sheet_name}, 날짜 수: {len(valid_dates)}, 종목 수: {len(stocks)})")


# =========================
# 5. QUANT 시트 전체 재계산 저장
# =========================

def save_quant_sheet(filename, dates, stocks, window=60, sheet_name='quant'):
    try:
        wb = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

    sheet = create_or_clear_sheet(wb, sheet_name)

    # 헤더
    sheet.cell(row=1, column=1, value='종목명')
    sheet.cell(row=1, column=2, value='종목코드')

    valid_dates = dates[window - 1:]
    for col, date in enumerate(valid_dates, start=3):
        cell = sheet.cell(row=1, column=col)
        cell.value = date
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

    sheet.cell(row=1, column=1).font = Font(bold=True)
    sheet.cell(row=1, column=1).fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
    sheet.cell(row=1, column=2).font = Font(bold=True)
    sheet.cell(row=1, column=2).fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

    # 데이터
    for row_idx, stock in enumerate(stocks, start=2):
        sheet.cell(row=row_idx, column=1, value=stock['name'])
        sheet.cell(row=row_idx, column=2, value=stock['code'])
        volumes = stock['volumes']

        for i in range(window - 1, len(volumes)):
            window_volumes = volumes[i - window + 1:i + 1]
            if None in window_volumes:
                val = None
            else:
                val = calc_quant(window_volumes)
            col_idx = 3 + (i - (window - 1))
            sheet.cell(row=row_idx, column=col_idx, value=val)

    # 열 너비
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 12
    for col_idx in range(3, len(valid_dates) + 3):
        sheet.column_dimensions[get_column_letter(col_idx)].width = 12

    wb.save(filename)
    print(f"✅ QUANT 전체 저장 완료: {filename} (탭: {sheet_name}, 날짜 수: {len(valid_dates)}, 종목 수: {len(stocks)})")


# =========================
# 6. STD 시트 전체 재계산 저장
# =========================

def save_std_sheet(filename, dates, stocks, sheet_name='std', window_std=20, window_mean=20):
    min_idx = window_std + window_mean - 2  # 첫 STD가 나오는 인덱스 (예: 38)

    try:
        wb = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

    sheet = create_or_clear_sheet(wb, sheet_name)

    # 헤더
    sheet.cell(row=1, column=1, value='종목명')
    sheet.cell(row=1, column=2, value='종목코드')

    valid_dates = dates[min_idx:]
    for col, date in enumerate(valid_dates, start=3):
        cell = sheet.cell(row=1, column=col)
        cell.value = date
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

    sheet.cell(row=1, column=1).font = Font(bold=True)
    sheet.cell(row=1, column=1).fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
    sheet.cell(row=1, column=2).font = Font(bold=True)
    sheet.cell(row=1, column=2).fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

    # 데이터
    for row_idx, stock in enumerate(stocks, start=2):
        sheet.cell(row=row_idx, column=1, value=stock['name'])
        sheet.cell(row=row_idx, column=2, value=stock['code'])
        prices = stock['prices']

        for i in range(min_idx, len(prices)):
            val = calc_std_value(prices, i, window_std=window_std, window_mean=window_mean)
            col_idx = 3 + (i - min_idx)
            sheet.cell(row=row_idx, column=col_idx, value=val)

    # 열 너비
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 12
    for col_idx in range(3, len(valid_dates) + 3):
        sheet.column_dimensions[get_column_letter(col_idx)].width = 12

    wb.save(filename)
    print(f"✅ STD 전체 저장 완료: {filename} (탭: {sheet_name}, 날짜 수: {len(valid_dates)}, 종목 수: {len(stocks)})")


# =========================
# 7. 통합 실행 함수
# =========================

def run_extra_scores(filename):
    print(f"\n=== EXTRA SCORES 계산 시작: {filename} ===")

    # GAP, STD는 '종가' 시트 기반
    close_dates, close_stocks = get_close_data(filename)
    if close_dates and close_stocks:
        save_gap_sheet(filename, close_dates, close_stocks, window=20, sheet_name='gap')
        save_std_sheet(filename, close_dates, close_stocks, sheet_name='std', window_std=20, window_mean=20)
    else:
        print("⚠ 종가 데이터가 없어 GAP/STD 계산을 건너뜁니다.")

    # QUANT는 '거래량' 시트 기반
    vol_dates, vol_stocks = get_volume_data(filename)
    if vol_dates and vol_stocks:
        save_quant_sheet(filename, vol_dates, vol_stocks, window=60, sheet_name='quant')
    else:
        print("⚠ 거래량 데이터가 없어 QUANT 계산을 건너뜁니다.")

    print(f"=== EXTRA SCORES 계산 완료: {filename} ===\n")


def main():
    run_extra_scores("KR_Stocks_ETF.xlsx")


if __name__ == "__main__":
    main()
