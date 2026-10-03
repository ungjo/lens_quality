from pathlib import Path
import sqlite3

import pandas as pd
from openpyxl.styles import PatternFill


# 파일 위치: 이 Python 파일이 있는 폴더를 기준으로 설정
BASE = Path(__file__).resolve().parent

INPUT = BASE / "data" / "lens_qualityCheck.xlsx"

RESULTS = BASE.parents[1] / "resources" / "lens_quality" / "outputs" / "results"
DB_FILE = RESULTS / "lens_quality.db"
OUTPUT = RESULTS / "lens_inspection_result.xlsx"


def check_value(item, value, lower, upper):
    """항목 하나의 측정값을 하한·상한과 비교합니다."""
    if pd.isna(value):
        return "판정불가", f"{item}: 측정값 누락"

    if pd.isna(lower) or pd.isna(upper):
        return "판정불가", f"{item}: 기준 누락"

    if lower > upper:
        return "판정불가", f"{item}: 하한이 상한보다 큼"

    if value < lower:
        return "불합격", f"{item}: {value:g} < 하한 {lower:g}"

    if value > upper:
        return "불합격", f"{item}: {value:g} > 상한 {upper:g}"

    return "합격", f"{item}: 허용범위 {lower:g}~{upper:g} 이내"


def format_inspection_results(writer, result, detail):
    """판정 색상과 불합격 원인 측정값을 표시합니다. 재실행 시 갱신됩니다."""
    fills = {
        "합격": PatternFill(fill_type="solid", fgColor="D9EAD3"),
        "불합격": PatternFill(fill_type="solid", fgColor="F4CCCC"),
    }
    failed_fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
    summary_sheet = writer.sheets["검사결과"]
    detail_sheet = writer.sheets["항목별상세"]

    # 열 위치가 바뀌어도 열 이름과 제품번호로 올바른 셀을 찾습니다.
    summary_columns = {name: i for i, name in enumerate(result.columns, 1)}
    detail_columns = {name: i for i, name in enumerate(detail.columns, 1)}
    product_rows = {}
    for row_number, record in enumerate(result.to_dict("records"), 2):
        product_rows[record["제품번호"]] = row_number
        fill = fills.get(record["판정"])
        if fill is not None:
            summary_sheet.cell(row_number, summary_columns["판정"]).fill = fill

    for row_number, record in enumerate(detail.to_dict("records"), 2):
        status = record["항목판정"]
        fill = fills.get(status)
        if fill is not None:
            detail_sheet.cell(row_number, detail_columns["항목판정"]).fill = fill
        if status == "불합격":
            summary_sheet.cell(
                product_rows[record["제품번호"]],
                summary_columns[record["검사항목"]],
            ).fill = failed_fill
            detail_sheet.cell(row_number, detail_columns["측정값"]).fill = failed_fill

def main():
    RESULTS.mkdir(parents=True, exist_ok=True)
    # 1. Excel 읽기
    # A:P에는 제품 정보와 측정값이 있습니다.
    # Excel의 기존 판정·판정사유 열은 읽지 않습니다.
    measurements = pd.read_excel(
        INPUT,
        sheet_name="측정데이터",
        usecols="A:P",
    )
    measurements = measurements.dropna(subset=["제품번호"])

    # 현재 기준표의 첫 15행: 공통 12개 + 소재별 굴절률 3개
    criteria = pd.read_excel(
        INPUT,
        sheet_name="기준표",
        usecols="A:F",
        nrows=15,
    )

    if measurements["제품번호"].duplicated().any():
        raise ValueError("제품번호가 중복되어 있습니다.")

    if criteria.duplicated(["검사항목", "적용소재"]).any():
        raise ValueError("같은 항목·소재의 기준이 중복되어 있습니다.")

    # 측정값과 기준을 숫자로 변환합니다.
    # 잘못된 문자열은 자동으로 0으로 바꾸지 않고 오류를 냅니다.
    items = criteria["검사항목"].drop_duplicates().tolist()

    for item in items:
        measurements[item] = pd.to_numeric(
            measurements[item], errors="raise"
        )

    for column in ["하한", "목표치", "상한"]:
        criteria[column] = pd.to_numeric(
            criteria[column], errors="raise"
        )

    with sqlite3.connect(DB_FILE) as conn:  #DB파일에 연결. 파일이 없으면 생성.
        # 2. SQL 테이블에 저장
        # 이 예제는 실행할 때마다 최신 데이터로 테이블을 교체합니다.
        measurements.to_sql(    #테이블 생성과 데이터 삽입을 코드로 작성하는 대신 to_sql()이 테이블로 저장한다.
            "measurements", conn, if_exists="replace", index=False
        )
        criteria.to_sql(
            "criteria", conn, if_exists="replace", index=False
        )

        # 3. SQL로 저장된 데이터 조회
        data = pd.read_sql_query(
            'SELECT * FROM measurements ORDER BY "제품번호"',
            conn,
        )
        rules = pd.read_sql_query("SELECT * FROM criteria", conn)

        summaries = []
        details = []

        # 4. 제품별로 Python 조건문을 사용해 판정
        for _, product in data.iterrows():
            failures = []
            unresolved = []

            for item in items:
                # 공통 기준 또는 해당 제품 소재의 기준을 찾습니다.
                matched = rules[
                    (rules["검사항목"] == item)
                    & (
                        (rules["적용소재"] == "공통")
                        | (rules["적용소재"] == product["소재종류"])
                    )
                ]

                value = product[item]
                lower = upper = None

                if len(matched) != 1:
                    status = "판정불가"
                    reason = f"{item}: 적용 기준 누락 또는 중복"
                else:
                    rule = matched.iloc[0]
                    lower = rule["하한"]
                    upper = rule["상한"]
                    status, reason = check_value(
                        item, value, lower, upper
                    )

                if status == "불합격":
                    failures.append(reason)
                elif status == "판정불가":
                    unresolved.append(reason)

                details.append({
                    "제품번호": product["제품번호"],
                    "검사항목": item,
                    "측정값": value,
                    "하한": lower,
                    "상한": upper,
                    "항목판정": status,
                    "사유": reason,
                })

            if failures:
                final_status = "불합격"
            elif unresolved:
                final_status = "판정불가"
            else:
                final_status = "합격"

            reasons = failures + unresolved

            summaries.append({
                **product.to_dict(),
                "판정": final_status,
                "판정사유": (
                    "\n".join(reasons)
                    if reasons
                    else f"{len(items)}개 항목 모두 허용범위 이내"
                ),
            })

        result = pd.DataFrame(summaries)
        detail = pd.DataFrame(details)

        # 5. 판정 결과도 SQL에 저장
        result.to_sql(
            "inspection_results", conn,
            if_exists="replace", index=False
        )
        detail.to_sql(
            "inspection_details", conn,
            if_exists="replace", index=False
        )

    # 6. 결과 Excel 생성
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="검사결과", index=False)
        detail.to_excel(writer, sheet_name="항목별상세", index=False)
        criteria.to_excel(writer, sheet_name="적용기준", index=False)
        format_inspection_results(writer, result, detail)

    print(result["판정"].value_counts().to_string())
    print(f"\nDB 저장: {DB_FILE}")
    print(f"Excel 저장: {OUTPUT}")


if __name__ == "__main__":
    main()


