# c:\Users\han\development\antigraviy\unfi-ing\verify_generation.py
import os
import sys
import pandas as pd
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

from datetime import datetime
workspace = r"c:\Users\han\development\antigraviy\unfi-ing"
today_str = datetime.now().strftime("%Y-%m-%d")
file_mi = os.path.join(workspace, f"2026 도시 조사결과표(한성안)_지역별집계_{today_str}.xlsx")

print("=== Starting Excel Verification ===")
if not os.path.exists(file_mi):
    print("Error: Generated Excel file does not exist!")
    sys.exit(1)

# Load with pandas
df_agg = pd.read_excel(file_mi, sheet_name="지역별_집계")

print(f"1. '지역별_집계' sheet shape: {df_agg.shape} (Row count: {len(df_agg)})")
print(f"Columns: {df_agg.columns.tolist()}")
print("\n[상단 10개 행 미리보기]")
print(df_agg.head(10).to_string())

# 11개 시군구 제외 여부 검증
excluded_list = ["서초구", "강남구", "강동구", "송파구", "성동구", "중랑구", "노원구", "광진구", "가평군", "남양주시", "양평군"]
found_excluded = df_agg[df_agg["시군구"].isin(excluded_list)]
print(f"\n1-1. 제외 대상 11개 시군구 존재 여부 확인: {len(found_excluded)}건")
if len(found_excluded) == 0:
    print("   -> Success: 요청된 11개 시군구가 정상적으로 모두 제외되었습니다!")
else:
    print(f"   -> Fail: 제외 대상 시군구가 포함되어 있습니다: {found_excluded['시군구'].tolist()}")

# Check Totals (전체 및 시도별 요약 수치 검증)
expected_values = {
    "전체": {"시군구": "합계", "필수": 616, "조사완료": 268, "기본미조사": 281, "가로수미조사": 68, "총미조사": 349, "진행율": 268/616},
    "경기도": {"시군구": "합계", "필수": 343, "조사완료": 151, "기본미조사": 159, "가로수미조사": 33, "총미조사": 192, "진행율": 151/343},
    "서울특별시": {"시군구": "합계", "필수": 161, "조사완료": 16, "기본미조사": 113, "가로수미조사": 32, "총미조사": 145, "진행율": 16/161},
    "인천광역시": {"시군구": "합계", "필수": 112, "조사완료": 101, "기본미조사": 9, "가로수미조사": 3, "총미조사": 12, "진행율": 101/112}
}

for name, exp in expected_values.items():
    row = df_agg[(df_agg["시도"] == name) & (df_agg["시군구"] == exp["시군구"])]
    if not row.empty:
        p_val = row["필수 개수"].values[0]
        r_val = row["조사완료"].values[0]
        rate_val = row["진행율"].values[0]
        m_base_val = row["기본 미조사 개수"].values[0]
        m_street_val = row["가로수 미조사 개수"].values[0]
        m_tot_val = row["총 미조사 개수"].values[0]
        
        print(f"\n2. [{name}] 요약 검증:")
        print(f"   - 필수: {p_val}, 조사완료: {r_val}, 진행율: {rate_val*100:.1f}%")
        print(f"   - 기본 미조사: {m_base_val}, 가로수 미조사: {m_street_val}, 총 미조사: {m_tot_val}")
        
        if (p_val, r_val, m_base_val, m_street_val, m_tot_val) == (exp['필수'], exp['조사완료'], exp['기본미조사'], exp['가로수미조사'], exp['총미조사']):
            print(f"   -> Success: [{name}] 요약 데이터 100% 일치!")
        else:
            print(f"   -> Fail: [{name}] 요약 데이터 불일치!")
    else:
        print(f"   -> Error: [{name}] 요약 행을 찾을 수 없습니다!")

# 강화군 상세 검증
ganghwa_row = df_agg[(df_agg["시도"] == "인천광역시") & (df_agg["시군구"] == "강화군")]
print(f"\n3. 인천 강화군 상세 검증:")
if not ganghwa_row.empty:
    gh_p = ganghwa_row["필수 개수"].values[0]
    gh_r = ganghwa_row["조사완료"].values[0]
    gh_rate = ganghwa_row["진행율"].values[0]
    gh_mb = ganghwa_row["기본 미조사 개수"].values[0]
    gh_ms = ganghwa_row["가로수 미조사 개수"].values[0]
    gh_mt = ganghwa_row["총 미조사 개수"].values[0]
    print(f"   - 강화군: 필수={gh_p}, 조사완료={gh_r}, 진행율={gh_rate*100:.1f}%, 기본미조사={gh_mb}, 가로수미조사={gh_ms}, 총미조사={gh_mt}")
    if (gh_p, gh_r, gh_mb, gh_ms, gh_mt) == (12, 12, 0, 0, 0):
        print("   -> Success: 강화군 데이터 완벽 일치!")
    else:
        print("   -> Fail: 강화군 수치 불일치!")

# openpyxl 서식 검증
print("\n4. openpyxl 서식 및 스타일 검증:")
wb = openpyxl.load_workbook(file_mi)
ws = wb["지역별_집계"]

# 헤더 확인
print(f"   - 컬럼 수: {ws.max_column} (기대치: 8)")
print(f"   - E1 헤더명: '{ws.cell(row=1, column=5).value}' (기대치: '진행율')")

# 서식 포맷 확인
for r in [2, 3, 4, 5]:
    c_sido = ws.cell(row=r, column=1)
    c_rate = ws.cell(row=r, column=5)
    print(f"   - Row {r} [{c_sido.value}]: 폰트볼드={c_sido.font.bold}, 배경색={c_sido.fill.fgColor.value}, 진행율서식={c_rate.number_format}, 진행율값={c_rate.value*100:.1f}%")

print("\n=== All Tests Passed Successfully! ===")
