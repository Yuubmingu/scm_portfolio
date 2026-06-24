import pandas as pd
import numpy as np
import re
from pathlib import Path

# =========================================================
# 0. 기본 경로 설정
# =========================================================

BASE_DIR = Path("/kaggle/working")

production_plan_path = BASE_DIR / "production_plan.xlsx"
bom_master_path = BASE_DIR / "bom_master.xlsx"
inventory_path = BASE_DIR / "inventory.xlsx"
po_open_path = BASE_DIR / "po_open.xlsx"
material_master_path = BASE_DIR / "material_master.xlsx"

output_path = BASE_DIR / "material_supply_plan.xlsx"


# =========================================================
# 1. 공통 함수 정의
# =========================================================

def print_df_info(name, df):
    """
    데이터프레임의 행 수와 컬럼명을 출력하는 함수
    """
    print(f"\n[{name}]")
    print(f"- 행 수: {len(df):,}")
    print(f"- 컬럼: {list(df.columns)}")


def check_file_exists(file_path, file_name):
    """
    파일 존재 여부 검증 함수
    """
    if not file_path.exists():
        raise FileNotFoundError(f"[{file_name}] 파일이 존재하지 않습니다: {file_path}")


def check_required_columns(df, required_columns, file_name):
    """
    필수 컬럼 존재 여부 검증 함수
    """
    missing_cols = [col for col in required_columns if col not in df.columns]

    if missing_cols:
        raise ValueError(f"[{file_name}]에 필수 컬럼이 없습니다: {missing_cols}")


def convert_numeric(df, columns):
    """
    숫자 컬럼을 숫자형으로 변환하는 함수
    변환 불가 값과 빈값은 0으로 처리
    """
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def clean_columns(df):
    """
    컬럼명 앞뒤 공백 제거
    """
    df.columns = df.columns.astype(str).str.strip()
    return df


def clean_code_column(df, column_name):
    """
    코드 컬럼을 문자열로 변환하고 앞뒤 공백 제거
    """
    if column_name in df.columns:
        df[column_name] = df[column_name].astype(str).str.strip()
    return df


def parse_iso_week(plan_week_value):
    """
    plan_week 값을 날짜 정보로 변환하는 함수

    처리 가능 예시:
    - 2026-W27
    - 2026W27
    - 2026-07-05 같은 일반 날짜

    반환 컬럼:
    - week_start_date: 해당 주 월요일
    - required_date: 해당 주 일요일
    - plan_month: required_date 기준 YYYY-MM
    - month_end_date: required_date가 속한 월의 말일
    """

    value = str(plan_week_value).strip()

    # 2026-W27 또는 2026W27 형식 처리
    match = re.match(r"^(\d{4})-?W(\d{1,2})$", value)

    if match:
        year = int(match.group(1))
        week = int(match.group(2))

        # ISO 주차 기준
        # 월요일 = 1, 일요일 = 7
        week_start_date = pd.to_datetime(
            f"{year}-W{week:02d}-1",
            format="%G-W%V-%u",
            errors="coerce"
        )

        required_date = pd.to_datetime(
            f"{year}-W{week:02d}-7",
            format="%G-W%V-%u",
            errors="coerce"
        )

        if pd.isna(required_date):
            return pd.Series({
                "week_start_date": pd.NaT,
                "required_date": pd.NaT,
                "plan_month": value,
                "month_end_date": pd.NaT
            })

        plan_month = required_date.strftime("%Y-%m")
        month_end_date = required_date + pd.offsets.MonthEnd(0)

        return pd.Series({
            "week_start_date": week_start_date,
            "required_date": required_date,
            "plan_month": plan_month,
            "month_end_date": month_end_date
        })

    # 일반 날짜 형식 처리
    parsed_date = pd.to_datetime(value, errors="coerce")

    if not pd.isna(parsed_date):
        return pd.Series({
            "week_start_date": parsed_date,
            "required_date": parsed_date,
            "plan_month": parsed_date.strftime("%Y-%m"),
            "month_end_date": parsed_date + pd.offsets.MonthEnd(0)
        })

    # 변환 실패 시 안전 처리
    return pd.Series({
        "week_start_date": pd.NaT,
        "required_date": pd.NaT,
        "plan_month": value,
        "month_end_date": pd.NaT
    })


def calculate_recommended_order_qty(shortage_qty, moq):
    """
    MOQ 기준 추천발주수량 계산

    원칙:
    1. 부족수량이 0이면 추천발주수량도 0
    2. MOQ가 0이면 부족수량 그대로 추천
    3. MOQ가 있으면 MOQ 배수로 올림 처리
    """
    if shortage_qty <= 0:
        return 0

    if moq <= 0:
        return shortage_qty

    return np.ceil(shortage_qty / moq) * moq


def adjust_excel_column_width(writer, sheet_name, df):
    """
    openpyxl을 사용해 엑셀 컬럼 너비와 표시 형식을 조정하는 함수

    적용 내용:
    1. 모든 컬럼 너비 자동 조정
    2. 컬럼명에 amount가 들어간 컬럼은 1,000 형식 적용
    3. 컬럼명에 date가 들어간 컬럼은 yyyy-mm-dd 날짜 형식 적용
    4. risk_level 컬럼 값에 따라 배경색 적용
       - High: 빨간색
       - Medium: 노란색
       - Low: 초록색
    """

    from openpyxl.styles import PatternFill, Font, Alignment

    worksheet = writer.sheets[sheet_name]

    # risk_level 색상 정의
    high_fill = PatternFill(
        fill_type="solid",
        fgColor="FFC7CE"  # 연한 빨강
    )

    medium_fill = PatternFill(
        fill_type="solid",
        fgColor="FFEB9C"  # 연한 노랑
    )

    low_fill = PatternFill(
        fill_type="solid",
        fgColor="C6EFCE"  # 연한 초록
    )

    normal_fill = PatternFill(
        fill_type="solid",
        fgColor="FFFFFF"  # 흰색
    )

    # risk_level 글자색 정의
    high_font = Font(color="9C0006")
    medium_font = Font(color="9C6500")
    low_font = Font(color="006100")
    normal_font = Font(color="000000")

    for idx, col in enumerate(df.columns, start=1):
        # =========================
        # 1) 컬럼 너비 자동 조정
        # =========================
        if len(df) > 0:
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(str(col))
            )
        else:
            max_length = len(str(col))

        adjusted_width = min(max_length + 2, 35)
        column_letter = worksheet.cell(row=1, column=idx).column_letter
        worksheet.column_dimensions[column_letter].width = adjusted_width

        # =========================
        # 2) 표시 형식 적용
        # =========================
        col_lower = str(col).lower()

        # amount 컬럼: 1,000 형식
        if "amount" in col_lower:
            for cell in worksheet[column_letter][1:]:
                cell.number_format = '#,##0'

        # date 컬럼: yyyy-mm-dd 형식
        if "date" in col_lower:
            for cell in worksheet[column_letter][1:]:
                cell.number_format = 'yyyy-mm-dd'

        # =========================
        # 3) risk_level 컬럼 색상 적용
        # =========================
        if col_lower == "risk_level":
            for cell in worksheet[column_letter][1:]:
                risk_value = str(cell.value).strip()

                if risk_value == "High":
                    cell.fill = high_fill
                    cell.font = high_font

                elif risk_value == "Medium":
                    cell.fill = medium_fill
                    cell.font = medium_font

                elif risk_value == "Low":
                    cell.fill = low_fill
                    cell.font = low_font

                else:
                    cell.fill = normal_fill
                    cell.font = normal_font

                cell.alignment = Alignment(horizontal="center")


# =========================================================
# 2. 파일 존재 여부 검증
# =========================================================

check_file_exists(production_plan_path, "production_plan.xlsx")
check_file_exists(bom_master_path, "bom_master.xlsx")
check_file_exists(inventory_path, "inventory.xlsx")
check_file_exists(po_open_path, "po_open.xlsx")
check_file_exists(material_master_path, "material_master.xlsx")

print("파일 존재 여부 검증 완료")


# =========================================================
# 3. 파일 읽기
# =========================================================

production_plan = pd.read_excel(production_plan_path)
bom_master = pd.read_excel(bom_master_path)
inventory = pd.read_excel(inventory_path)
po_open = pd.read_excel(po_open_path)
material_master = pd.read_excel(material_master_path)

# 컬럼명 공백 제거
production_plan = clean_columns(production_plan)
bom_master = clean_columns(bom_master)
inventory = clean_columns(inventory)
po_open = clean_columns(po_open)
material_master = clean_columns(material_master)

print_df_info("production_plan 원본", production_plan)
print_df_info("bom_master 원본", bom_master)
print_df_info("inventory 원본", inventory)
print_df_info("po_open 원본", po_open)
print_df_info("material_master 원본", material_master)


# =========================================================
# 4. 필수 컬럼 검증
# =========================================================

production_plan_required_cols = [
    "plan_week",
    "customer",
    "model",
    "product_code",
    "qty"
]

bom_master_required_cols = [
    "product_code",
    "material_code",
    "material_name",
    "usage_qty"
]

inventory_required_cols = [
    "material_code",
    "current_stock"
]

po_open_required_cols = [
    "material_code",
    "po_no",
    "eta_date",
    "open_qty"
]

material_master_required_cols = [
    "material_code",
    "supplier",
    "lead_time_days",
    "moq",
    "unit_price"
]

check_required_columns(production_plan, production_plan_required_cols, "production_plan.xlsx")
check_required_columns(bom_master, bom_master_required_cols, "bom_master.xlsx")
check_required_columns(inventory, inventory_required_cols, "inventory.xlsx")
check_required_columns(po_open, po_open_required_cols, "po_open.xlsx")
check_required_columns(material_master, material_master_required_cols, "material_master.xlsx")

print("\n필수 컬럼 검증 완료")


# =========================================================
# 5. 숫자 컬럼 처리
# =========================================================

production_plan = convert_numeric(production_plan, ["qty"])
bom_master = convert_numeric(bom_master, ["usage_qty"])
inventory = convert_numeric(inventory, ["current_stock"])
po_open = convert_numeric(po_open, ["open_qty"])
material_master = convert_numeric(material_master, ["lead_time_days", "moq", "unit_price"])

print("\n숫자 컬럼 변환 완료")


# =========================================================
# 6. 코드 컬럼 타입 정리
# =========================================================

production_plan = clean_code_column(production_plan, "product_code")
bom_master = clean_code_column(bom_master, "product_code")
bom_master = clean_code_column(bom_master, "material_code")
inventory = clean_code_column(inventory, "material_code")
po_open = clean_code_column(po_open, "material_code")
material_master = clean_code_column(material_master, "material_code")

print("\n코드 컬럼 타입 정리 완료")


# =========================================================
# 7. plan_week 날짜 정보 생성
# =========================================================

date_info = production_plan["plan_week"].apply(parse_iso_week)
production_plan = pd.concat([production_plan, date_info], axis=1)

print("\nplan_week 날짜 변환 결과")
print(
    production_plan[
        ["plan_week", "week_start_date", "required_date", "plan_month", "month_end_date"]
    ].drop_duplicates()
)

# 날짜 변환 실패 경고
failed_date_count = production_plan["required_date"].isna().sum()

if failed_date_count > 0:
    print(f"\n경고: plan_week 날짜 변환 실패 행이 {failed_date_count:,}건 있습니다.")
    print("해당 행은 plan_month는 원본 문자열 기준으로 유지되며, 월별 입고 계산에서는 입고수량이 0으로 처리될 수 있습니다.")


# =========================================================
# 8. po_open eta_date 날짜 변환
# =========================================================

po_open["eta_date"] = pd.to_datetime(po_open["eta_date"], errors="coerce")

invalid_eta_count = po_open["eta_date"].isna().sum()

if invalid_eta_count > 0:
    print(f"\n경고: po_open.xlsx의 eta_date 중 날짜 변환 실패 행이 {invalid_eta_count:,}건 있습니다.")
    print("해당 행은 월별 입고예정수량 계산에서 제외됩니다.")

po_open_valid = po_open.dropna(subset=["eta_date"]).copy()

print_df_info("po_open 날짜 변환 후 유효 데이터", po_open_valid)


# =========================================================
# 9. BOM 누락 product_code 경고
# =========================================================

plan_product_codes = set(production_plan["product_code"].dropna().astype(str))
bom_product_codes = set(bom_master["product_code"].dropna().astype(str))

missing_bom_product_codes = sorted(list(plan_product_codes - bom_product_codes))

if missing_bom_product_codes:
    print("\n경고: BOM에 없는 product_code가 있습니다.")
    print("아래 product_code는 BOM이 없어 필요수량 계산에서 제외됩니다.")
    print(missing_bom_product_codes)
else:
    print("\nBOM 누락 product_code 없음")


# =========================================================
# 10. production_plan과 bom_master 병합
# =========================================================

plan_bom = production_plan.merge(
    bom_master,
    on="product_code",
    how="left",
    indicator=True
)

print_df_info("production_plan + bom_master 병합 결과", plan_bom)

bom_missing_rows = plan_bom[plan_bom["_merge"] == "left_only"]

if len(bom_missing_rows) > 0:
    print(f"\n경고: BOM 매칭 실패 행 수: {len(bom_missing_rows):,}")
    print("BOM이 없는 생산계획 행은 필요수량 계산에서 제외됩니다.")
    print(
        bom_missing_rows[
            ["plan_week", "customer", "model", "product_code", "qty"]
        ].drop_duplicates()
    )

# BOM이 있는 행만 계산에 사용
plan_bom = plan_bom[plan_bom["_merge"] == "both"].copy()
plan_bom = plan_bom.drop(columns=["_merge"])


# =========================================================
# 11. 자재 필요수량 계산
# =========================================================

plan_bom["required_qty"] = plan_bom["qty"] * plan_bom["usage_qty"]

required_by_material = (
    plan_bom
    .groupby(
        [
            "plan_week",
            "plan_month",
            "month_end_date",
            "material_code",
            "material_name"
        ],
        as_index=False,
        dropna=False
    )
    .agg(
        required_qty=("required_qty", "sum"),
        week_start_date=("week_start_date", "min"),
        required_date=("required_date", "min")
    )
)

print_df_info("자재별 주차/월별 필요수량 required_by_material", required_by_material)
print(required_by_material.head())

# 핵심 검증: 필요수량 결과가 0행이면 실패
if len(required_by_material) == 0:
    raise ValueError(
        "자재별 필요수량 required_by_material이 0행입니다. "
        "BOM merge, plan_week 변환, groupby 기준을 확인하세요."
    )


# =========================================================
# 12. 월별 입고예정수량 계산
# =========================================================
# 기준:
# eta_date <= 해당 plan_month의 month_end_date 인 open_qty만 incoming_qty에 포함

incoming_base = required_by_material[
    ["plan_week", "plan_month", "month_end_date", "material_code"]
].drop_duplicates()

incoming_merge = incoming_base.merge(
    po_open_valid[["material_code", "po_no", "eta_date", "open_qty"]],
    on="material_code",
    how="left"
)

incoming_merge["is_available_in_month"] = (
    incoming_merge["eta_date"].notna()
    & incoming_merge["month_end_date"].notna()
    & (incoming_merge["eta_date"] <= incoming_merge["month_end_date"])
)

incoming_merge["incoming_qty_for_month"] = np.where(
    incoming_merge["is_available_in_month"],
    incoming_merge["open_qty"],
    0
)

incoming_by_material_month = (
    incoming_merge
    .groupby(
        ["plan_week", "plan_month", "material_code"],
        as_index=False
    )
    .agg(incoming_qty=("incoming_qty_for_month", "sum"))
)

print_df_info("월별 입고예정수량 incoming_by_material_month", incoming_by_material_month)
print(incoming_by_material_month.head())


# =========================================================
# 13. 재고 정보 정리
# =========================================================

inventory_by_material = (
    inventory
    .groupby("material_code", as_index=False)
    .agg(current_stock=("current_stock", "sum"))
)

print_df_info("자재별 현재고 inventory_by_material", inventory_by_material)
print(inventory_by_material.head())


# =========================================================
# 14. 자재마스터 정리
# =========================================================

material_master_clean = (
    material_master
    .sort_values("material_code")
    .drop_duplicates(subset=["material_code"], keep="first")
    [
        [
            "material_code",
            "supplier",
            "lead_time_days",
            "moq",
            "unit_price"
        ]
    ]
)

print_df_info("자재마스터 정리 material_master_clean", material_master_clean)
print(material_master_clean.head())


# =========================================================
# 15. 필요수량 + 재고 + 월별 입고 + 자재마스터 연결
# =========================================================

supply_plan = required_by_material.copy()

supply_plan = supply_plan.merge(
    inventory_by_material,
    on="material_code",
    how="left"
)

supply_plan = supply_plan.merge(
    incoming_by_material_month,
    on=["plan_week", "plan_month", "material_code"],
    how="left"
)

supply_plan = supply_plan.merge(
    material_master_clean,
    on="material_code",
    how="left"
)

print_df_info("수급계획 병합 결과 supply_plan", supply_plan)

# 핵심 검증: 최종 병합 결과가 0행이면 실패
if len(supply_plan) == 0:
    raise ValueError(
        "최종 supply_plan이 0행입니다. "
        "수급계획 결과가 비어 있으므로 정상 산출로 볼 수 없습니다."
    )


# =========================================================
# 16. 빈값 처리
# =========================================================

supply_plan["current_stock"] = supply_plan["current_stock"].fillna(0)
supply_plan["incoming_qty"] = supply_plan["incoming_qty"].fillna(0)
supply_plan["supplier"] = supply_plan["supplier"].fillna("Unknown")
supply_plan["lead_time_days"] = supply_plan["lead_time_days"].fillna(0)
supply_plan["moq"] = supply_plan["moq"].fillna(0)
supply_plan["unit_price"] = supply_plan["unit_price"].fillna(0)

numeric_cols = [
    "required_qty",
    "current_stock",
    "incoming_qty",
    "lead_time_days",
    "moq",
    "unit_price"
]

supply_plan = convert_numeric(supply_plan, numeric_cols)


# =========================================================
# 17. 가용수량, 부족수량, 과잉수량 계산
# =========================================================

supply_plan["available_qty"] = (
    supply_plan["current_stock"] + supply_plan["incoming_qty"]
)

supply_plan["shortage_qty"] = (
    supply_plan["required_qty"] - supply_plan["available_qty"]
).clip(lower=0)

supply_plan["excess_qty"] = (
    supply_plan["available_qty"] - supply_plan["required_qty"]
).clip(lower=0)


# =========================================================
# 18. 부족금액, 과잉금액 계산
# =========================================================

supply_plan["shortage_amount"] = (
    supply_plan["shortage_qty"] * supply_plan["unit_price"]
)

supply_plan["excess_amount"] = (
    supply_plan["excess_qty"] * supply_plan["unit_price"]
)


# =========================================================
# 19. MOQ 기준 추천발주수량 계산
# =========================================================

supply_plan["recommended_order_qty"] = supply_plan.apply(
    lambda row: calculate_recommended_order_qty(
        row["shortage_qty"],
        row["moq"]
    ),
    axis=1
)

supply_plan["recommended_order_amount"] = (
    supply_plan["recommended_order_qty"] * supply_plan["unit_price"]
)

print("\nMOQ 기준 추천발주수량 계산 완료")
print(
    supply_plan[
        [
            "material_code",
            "shortage_qty",
            "moq",
            "recommended_order_qty",
            "recommended_order_amount"
        ]
    ].head(10)
)


# =========================================================
# 20. 리드타임 기준 발주필요일 계산
# =========================================================
# 계산식:
# order_required_date = required_date - lead_time_days

supply_plan["order_required_date"] = supply_plan["required_date"] - pd.to_timedelta(
    supply_plan["lead_time_days"],
    unit="D"
)

print("\n리드타임 기준 발주필요일 계산 완료")
print(
    supply_plan[
        [
            "material_code",
            "required_date",
            "lead_time_days",
            "order_required_date"
        ]
    ].head(10)
)


# =========================================================
# 21. 리스크 등급 계산
# =========================================================

conditions = [
    (supply_plan["shortage_qty"] > 0) & (supply_plan["lead_time_days"] >= 45),
    (supply_plan["shortage_qty"] > 0) & (supply_plan["lead_time_days"] < 45),
    (supply_plan["shortage_qty"] == 0) & (supply_plan["excess_qty"] > 0),
    (supply_plan["shortage_qty"] == 0) & (supply_plan["excess_qty"] == 0)
]

choices = ["High", "Medium", "Low", "Normal"]

supply_plan["risk_level"] = np.select(
    conditions,
    choices,
    default="Normal"
)


# =========================================================
# 22. 계산 결과 검증
# =========================================================

validation_errors = []

if not (supply_plan["shortage_qty"] >= 0).all():
    validation_errors.append("shortage_qty에 음수가 있습니다.")

if not (supply_plan["excess_qty"] >= 0).all():
    validation_errors.append("excess_qty에 음수가 있습니다.")

if not np.allclose(
    supply_plan["available_qty"],
    supply_plan["current_stock"] + supply_plan["incoming_qty"]
):
    validation_errors.append("available_qty 계산이 current_stock + incoming_qty와 일치하지 않습니다.")

if not np.allclose(
    supply_plan["shortage_amount"],
    supply_plan["shortage_qty"] * supply_plan["unit_price"]
):
    validation_errors.append("shortage_amount 계산이 shortage_qty × unit_price와 일치하지 않습니다.")

if not np.allclose(
    supply_plan["excess_amount"],
    supply_plan["excess_qty"] * supply_plan["unit_price"]
):
    validation_errors.append("excess_amount 계산이 excess_qty × unit_price와 일치하지 않습니다.")

if not np.allclose(
    supply_plan["recommended_order_amount"],
    supply_plan["recommended_order_qty"] * supply_plan["unit_price"]
):
    validation_errors.append("recommended_order_amount 계산이 recommended_order_qty × unit_price와 일치하지 않습니다.")

if (supply_plan["recommended_order_qty"] < supply_plan["shortage_qty"]).any():
    validation_errors.append("recommended_order_qty가 shortage_qty보다 작은 행이 있습니다.")

if validation_errors:
    print("\n계산 결과 검증 실패")
    for error in validation_errors:
        print(f"- {error}")

    raise ValueError("계산 결과 검증 중 오류가 발견되었습니다.")
else:
    print("\n계산 결과 검증 완료")


# =========================================================
# 23. 최종 컬럼 정리
# =========================================================

final_columns = [
    "plan_week",
    "plan_month",
    "week_start_date",
    "required_date",
    "month_end_date",
    "material_code",
    "material_name",
    "required_qty",
    "current_stock",
    "incoming_qty",
    "available_qty",
    "shortage_qty",
    "shortage_amount",
    "excess_qty",
    "excess_amount",
    "supplier",
    "lead_time_days",
    "moq",
    "unit_price",
    "recommended_order_qty",
    "recommended_order_amount",
    "order_required_date",
    "risk_level"
]

supply_plan = supply_plan[final_columns]


# =========================================================
# 24. 정렬
# =========================================================

risk_order = {
    "High": 1,
    "Medium": 2,
    "Low": 3,
    "Normal": 4
}

supply_plan["risk_sort"] = supply_plan["risk_level"].map(risk_order)

supply_plan = (
    supply_plan
    .sort_values(
        by=["risk_sort", "shortage_amount", "material_code"],
        ascending=[True, False, True]
    )
    .drop(columns=["risk_sort"])
    .reset_index(drop=True)
)

print_df_info("최종 수급계획 supply_plan", supply_plan)
print(supply_plan.head(10))


# =========================================================
# 25. 요약 시트 생성: summary_by_risk
# =========================================================

summary_by_risk = (
    supply_plan
    .groupby("risk_level", as_index=False)
    .agg(
        material_count=("material_code", "nunique"),
        total_required_qty=("required_qty", "sum"),
        total_shortage_qty=("shortage_qty", "sum"),
        total_shortage_amount=("shortage_amount", "sum"),
        total_excess_qty=("excess_qty", "sum"),
        total_excess_amount=("excess_amount", "sum"),
        total_recommended_order_qty=("recommended_order_qty", "sum"),
        total_recommended_order_amount=("recommended_order_amount", "sum")
    )
)

summary_by_risk["risk_sort"] = summary_by_risk["risk_level"].map(risk_order)

summary_by_risk = (
    summary_by_risk
    .sort_values("risk_sort")
    .drop(columns=["risk_sort"])
    .reset_index(drop=True)
)

print_df_info("리스크별 요약 summary_by_risk", summary_by_risk)
print(summary_by_risk)


# =========================================================
# 26. 요약 시트 생성: summary_by_supplier
# =========================================================

supplier_risk_count = (
    supply_plan
    .pivot_table(
        index="supplier",
        columns="risk_level",
        values="material_code",
        aggfunc="nunique",
        fill_value=0
    )
    .reset_index()
)

if "High" not in supplier_risk_count.columns:
    supplier_risk_count["High"] = 0

if "Medium" not in supplier_risk_count.columns:
    supplier_risk_count["Medium"] = 0

supplier_amount_summary = (
    supply_plan
    .groupby("supplier", as_index=False)
    .agg(
        material_count=("material_code", "nunique"),
        total_shortage_qty=("shortage_qty", "sum"),
        total_shortage_amount=("shortage_amount", "sum"),
        total_excess_qty=("excess_qty", "sum"),
        total_excess_amount=("excess_amount", "sum"),
        total_recommended_order_qty=("recommended_order_qty", "sum"),
        total_recommended_order_amount=("recommended_order_amount", "sum")
    )
)

summary_by_supplier = supplier_amount_summary.merge(
    supplier_risk_count[["supplier", "High", "Medium"]],
    on="supplier",
    how="left"
)

summary_by_supplier = summary_by_supplier.rename(
    columns={
        "High": "high_risk_count",
        "Medium": "medium_risk_count"
    }
)

summary_by_supplier["high_risk_count"] = summary_by_supplier["high_risk_count"].fillna(0).astype(int)
summary_by_supplier["medium_risk_count"] = summary_by_supplier["medium_risk_count"].fillna(0).astype(int)

summary_by_supplier = summary_by_supplier[
    [
        "supplier",
        "material_count",
        "high_risk_count",
        "medium_risk_count",
        "total_shortage_qty",
        "total_shortage_amount",
        "total_excess_qty",
        "total_excess_amount",
        "total_recommended_order_qty",
        "total_recommended_order_amount"
    ]
]

summary_by_supplier = summary_by_supplier.sort_values(
    by=["total_shortage_amount", "high_risk_count", "supplier"],
    ascending=[False, False, True]
).reset_index(drop=True)

print_df_info("협력사별 요약 summary_by_supplier", summary_by_supplier)
print(summary_by_supplier.head(10))


# =========================================================
# 27. 최종 결과 해석 가능 여부 검증
# =========================================================

interpretation_required_columns = [
    "plan_week",
    "plan_month",
    "material_code",
    "material_name",
    "required_qty",
    "current_stock",
    "incoming_qty",
    "available_qty",
    "shortage_qty",
    "shortage_amount",
    "excess_qty",
    "excess_amount",
    "supplier",
    "lead_time_days",
    "moq",
    "unit_price",
    "recommended_order_qty",
    "recommended_order_amount",
    "order_required_date",
    "risk_level"
]

check_required_columns(
    supply_plan,
    interpretation_required_columns,
    "최종 material_supply_plan"
)

if len(supply_plan) == 0:
    raise ValueError("최종 결과가 0행입니다. 정상적인 수급계획서로 볼 수 없습니다.")

print("\n최종 결과 해석 가능 여부 검증 완료")


# =========================================================
# 28. Excel 저장
# =========================================================

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    supply_plan.to_excel(
        writer,
        sheet_name="material_supply_plan",
        index=False
    )

    summary_by_risk.to_excel(
        writer,
        sheet_name="summary_by_risk",
        index=False
    )

    summary_by_supplier.to_excel(
        writer,
        sheet_name="summary_by_supplier",
        index=False
    )

    adjust_excel_column_width(writer, "material_supply_plan", supply_plan)
    adjust_excel_column_width(writer, "summary_by_risk", summary_by_risk)
    adjust_excel_column_width(writer, "summary_by_supplier", summary_by_supplier)

print("\nExcel 저장 완료")


# =========================================================
# 29. 실행 후 확인용 조회 코드
# =========================================================

print("\n결과 파일 생성 여부")
print(output_path.exists())

high_risk_count = (supply_plan["risk_level"] == "High").sum()
medium_risk_count = (supply_plan["risk_level"] == "Medium").sum()

total_shortage_qty = supply_plan["shortage_qty"].sum()
total_shortage_amount = supply_plan["shortage_amount"].sum()

total_excess_qty = supply_plan["excess_qty"].sum()
total_excess_amount = supply_plan["excess_amount"].sum()

total_recommended_order_qty = supply_plan["recommended_order_qty"].sum()
total_recommended_order_amount = supply_plan["recommended_order_amount"].sum()

if len(supply_plan) > 0 and supply_plan["shortage_amount"].max() > 0:
    top_shortage_material = supply_plan.sort_values(
        "shortage_amount",
        ascending=False
    ).head(1)
else:
    top_shortage_material = pd.DataFrame()

print("\n부족금액이 가장 큰 자재")
if len(top_shortage_material) > 0:
    print(
        top_shortage_material[
            [
                "plan_week",
                "plan_month",
                "required_date",
                "material_code",
                "material_name",
                "shortage_qty",
                "shortage_amount",
                "moq",
                "recommended_order_qty",
                "recommended_order_amount",
                "supplier",
                "lead_time_days",
                "order_required_date",
                "risk_level"
            ]
        ]
    )
else:
    print("부족금액이 있는 자재가 없습니다.")


# =========================================================
# 30. 완료 보고
# =========================================================

print("\n==============================")
print("자재 수급계획 생성 완료")
print("==============================")
print(f"생성 완료 파일 경로: {output_path}")
print(f"전체 행 수: {len(supply_plan):,}")
print(f"전체 자재 수: {supply_plan['material_code'].nunique():,}")
print(f"High 리스크 자재 수: {high_risk_count:,}")
print(f"Medium 리스크 자재 수: {medium_risk_count:,}")
print(f"총 부족수량: {total_shortage_qty:,.0f}")
print(f"총 부족금액: {total_shortage_amount:,.0f}")
print(f"총 과잉수량: {total_excess_qty:,.0f}")
print(f"총 과잉금액: {total_excess_amount:,.0f}")
print(f"총 추천발주수량: {total_recommended_order_qty:,.0f}")
print(f"총 추천발주금액: {total_recommended_order_amount:,.0f}")


# =========================================================
# 31. 특정 material_code 검색 예시
# =========================================================

# 아래 값을 원하는 자재코드로 바꿔서 검색하면 된다.
# 예: search_material_code = "MAT-001"

search_material_code = ""

if search_material_code:
    search_result = supply_plan[
        supply_plan["material_code"].astype(str) == str(search_material_code)
    ]

    print(f"\nmaterial_code 검색 결과: {search_material_code}")
    print(search_result)
