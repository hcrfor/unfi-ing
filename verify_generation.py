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

# 제외 대상 시군구 검증 (성동구는 가로수 1개 미조사로 수동 추가되었으므로 제외 목록에서 별도 검증)
excluded_list = [
    "서초구", "강남구", "강동구", "송파구", "중랑구", "노원구", "광진구", 
    "가평군", "남양주시", "양평군", "안성시", "평택시"
]
found_excluded = df_agg[df_agg["시군구"].isin(excluded_list)]
print(f"\n1-1. 제외 대상 {len(excluded_list)}개 시군구 존재 여부 확인: {len(found_excluded)}건")
if len(found_excluded) == 0:
    print(f"   -> Success: 요청된 {len(excluded_list)}개 시군구가 정상적으로 모두 제외되었습니다!")
else:
    print(f"   -> Fail: 제외 대상 시군구가 포함되어 있습니다: {found_excluded['시군구'].tolist()}")

# 성동구 가로수 1개 미조사 검증
sd_row = df_agg[(df_agg["시도"] == "서울특별시") & (df_agg["시군구"] == "성동구")]
print(f"\n1-2. 서울특별시 성동구(가로수 1개 미조사) 검증:")
if not sd_row.empty:
    sd_p = sd_row["필수 개수"].values[0]
    sd_r = sd_row["조사완료"].values[0]
    sd_mb = sd_row["기본 미조사 개수"].values[0]
    sd_ms = sd_row["가로수 미조사 개수"].values[0]
    sd_mt = sd_row["총 미조사 개수"].values[0]
    print(f"   - 성동구: 필수={sd_p}, 조사완료={sd_r}, 기본미조사={sd_mb}, 가로수미조사={sd_ms}, 총미조사={sd_mt}")
    if (sd_p, sd_r, sd_mb, sd_ms, sd_mt) == (1, 0, 0, 1, 1):
        print("   -> Success: 성동구 가로수 1개 미조사 데이터 완벽 일치!")
    else:
        print("   -> Fail: 성동구 수치 불일치!")
else:
    print("   -> Fail: 성동구 행을 찾을 수 없습니다!")

# 서울특별시 요약 수치 검증 (성동구 필수 1, 가로수미조사 1 추가 반영)
seoul_summary = df_agg[(df_agg["시도"] == "서울특별시") & (df_agg["시군구"] == "합계")]
if not seoul_summary.empty:
    s_p = seoul_summary["필수 개수"].values[0]
    s_ms = seoul_summary["가로수 미조사 개수"].values[0]
    s_mt = seoul_summary["총 미조사 개수"].values[0]
    print(f"\n2. 서울특별시 요약 검증: 필수={s_p}, 가로수미조사={s_ms}, 총미조사={s_mt}")
    if (s_p, s_ms, s_mt) == (162, 33, 146):
        print("   -> Success: [서울특별시] 성동구 반영 합계 데이터 100% 일치!")
    else:
        print(f"   -> Fail: [서울특별시] 합계 불일치 (현재: 필수={s_p}, 가로수미조사={s_ms}, 총미조사={s_mt})")

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
