import json
import os

from totalSZ import run_total_sz
from extra_scores import run_extra_scores


# JSON 파일 경로 (필요하면 여기 이름만 바꿔줘)
JSON_PATH = "stock_file_map.json"


def load_excel_map(json_path=JSON_PATH):
    """
    JSON에서 카테고리 -> 엑셀파일명 매핑을 읽어온다.
    예)
    {
      "KR_Stocks_Individual": "KR_Stocks_Individual.xlsx",
      ...
    }
    """
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"⚠ JSON 파일을 찾을 수 없습니다: {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise ValueError("⚠ JSON 최상위 구조는 dict(객체)여야 합니다. { ... } 형태인지 확인해주세요.")

    return data


def run_all_scores_for_file(category_name, filename):
    """
    하나의 엑셀 파일에 대해:
      - S/Z 점수 (s20/s60/s120, z20/z60/z120)
      - extra scores (gap, quant, std)
    를 모두 실행한다.
    """
    if not os.path.exists(filename):
        print(f"⚠ [{category_name}] 파일 없음: {filename}  → 건너뜀")
        return

    print(f"\n=== [{category_name}] {filename} 처리 시작 ===")

    # 1) S/Z 점수 계산
    try:
        run_total_sz(filename)
    except Exception as e:
        print(f"⚠ [{category_name}] S/Z 계산 중 오류: {e}")

    # 2) GAP / QUANT / STD 계산
    try:
        run_extra_scores(filename)
    except Exception as e:
        print(f"⚠ [{category_name}] EXTRA SCORES 계산 중 오류: {e}")

    print(f"=== [{category_name}] {filename} 처리 완료 ===")


def main():
    # 1) JSON에서 엑셀 파일 목록 로드
    excel_map = load_excel_map(JSON_PATH)

    print(f"\n📁 JSON에서 {len(excel_map)}개 항목을 읽었습니다.")
    for category, filename in excel_map.items():
        run_all_scores_for_file(category, filename)

    print("\n✅ 모든 파일 처리 완료!")


if __name__ == "__main__":
    main()
