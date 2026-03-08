import streamlit as st
import pandas as pd
import requests
import xlsxwriter
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="AtheraCLOUD Stability Planner", layout="wide")

# 1. Notion 데이터 호출 (기존 로직 활용)
@st.cache_data(ttl=60)
def fetch_notion_data(database_id, token):
    url = f"https://api.notion.com/v1/databases/{database_id}/query"
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers)
    if response.status_code != 200: return pd.DataFrame()
    results = response.json().get("results", [])
    data = []
    for page in results:
        props = page.get("properties", {})
        row = {}
        for key, val in props.items():
            p_type = val.get("type")
            if p_type == "title": row[key] = val["title"][0]["plain_text"] if val["title"] else ""
            elif p_type in ["rich_text", "select"]:
                if p_type == "select": row[key] = val["select"]["name"] if val["select"] else ""
                else: row[key] = val["rich_text"][0]["plain_text"] if val["rich_text"] else ""
            else: row[key] = str(val.get(p_type, ""))
        data.append(row)
    return pd.DataFrame(data)

# --- UI 설정 ---
st.title("📉 Tool 4: Stability Study Protocol Planner")
st.markdown("ICH Q1A(R2) 기반의 안정성 시험 매트릭스를 자동으로 생성합니다.")

# 사이드바 설정: 보관 조건 선택
st.sidebar.header("❄️ Storage Conditions")
conditions = st.sidebar.multiselect(
    "보관 조건 선택",
    ["Long-term (5°C ± 3°C)", "Accelerated (25°C / 60% RH)", "Stress (40°C / 75% RH)"],
    default=["Long-term (5°C ± 3°C)", "Accelerated (25°C / 60% RH)"]
)
start_date = st.sidebar.date_input("안정성 시험 착수일", datetime(2026, 8, 1))

# 데이터 로드
try:
    df = fetch_notion_data(st.secrets["NOTION_DB_ID"], st.secrets["NOTION_TOKEN"])
except:
    st.error("Secrets 설정을 확인해주세요.")
    st.stop()

if not df.empty:
    # 안정성 지시력이 있는 항목만 필터링
    stab_df = df[df['Stability-indicating'].str.lower().isin(['yes', 'partial'])]
    st.success(f"🟢 노션에서 {len(stab_df)}개의 안정성 시험 대상 항목을 확인했습니다.")
    st.dataframe(stab_df[['Category', 'Method', 'Stability-indicating']], use_container_width=True)

    def create_stability_excel(dataframe, conds, start_dt):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        
        # 스타일 설정
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
        cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        mark_fmt = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'bold': True})

        # 각 보관 조건별 시트 생성
        for cond in conds:
            sheet_name = cond.split(' (')[0]
            sheet = workbook.add_worksheet(sheet_name)
            
            # 헤더: 시험 항목 및 주차(Timepoints)
            headers = ['Category', 'Method', 'Attribute']
            timepoints = ['T0', '1M', '3M', '6M', '9M', '12M', '18M', '24M']
            
            for c, h in enumerate(headers + timepoints):
                sheet.write(0, c, h, header_fmt)
                sheet.set_column(c, c, 12)

            # 데이터 작성
            for r, (_, row) in enumerate(dataframe.iterrows(), start=1):
                sheet.write(r, 0, str(row.get('Category', '')), cell_fmt)
                sheet.write(r, 1, str(row.get('Method', '')), cell_fmt)
                sheet.write(r, 2, str(row.get('Attribute', '')), cell_fmt)
                
                # 시험 주기별 체크 표시 (자동 매트릭스)
                for c in range(len(timepoints)):
                    # 가속 조건(Accelerated)은 통상 6개월까지만 표시하는 로직 등 추가 가능
                    if "Accelerated" in cond and c > 3: # 6M(index 3) 이후는 제외
                        sheet.write(r, 3 + c, "-", cell_fmt)
                    else:
                        sheet.write(r, 3 + c, "X", mark_fmt)

        workbook.close()
        return output.getvalue()

    if st.button("📊 안정성 시험 매트릭스(Excel) 추출"):
        excel_file = create_stability_excel(stab_df, conditions, start_date)
        st.download_button("💾 Protocol_Draft.xlsx 다운로드", excel_file, "Stability_Protocol.xlsx")

else:
    st.warning("안정성 시험 대상 데이터를 찾을 수 없습니다.")