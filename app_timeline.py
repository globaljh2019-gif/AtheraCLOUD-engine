import streamlit as st
import pandas as pd
import requests
import xlsxwriter
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="AtheraCLOUD Milestone Planner", layout="wide")

# 1. Notion API 설정
try:
    NOTION_TOKEN = st.secrets["NOTION_TOKEN"]
    DATABASE_ID = st.secrets["NOTION_DB_ID"]
except Exception:
    st.error("⚠️ Secrets 설정에서 NOTION_TOKEN과 NOTION_DB_ID를 확인하세요.")
    st.stop()

# 2. 데이터 호출 함수
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

# --- 🎯 마일스톤 전략 설정 UI ---
st.title("🚀 Tool 2: Strategic CMC Milestone Planner")
st.sidebar.header("📋 Project Strategy")

# 개발 단계 및 물질 선택
substance_type = st.sidebar.selectbox("물질 유형 (Substance Type)", ["원료의약품 (DS)", "완제의약품 (DP)"])
dev_stage = st.sidebar.selectbox("개발 단계 (Phase)", ["Pre-IND / Phase 1", "Phase 2", "Phase 3 / NDA"])
base_date = st.sidebar.date_input("CMC 프로젝트 착수일", datetime(2026, 3, 1))

# 단계별 마일스톤 가중치/기간 설정 (예시)
if dev_stage == "Pre-IND / Phase 1":
    dev_weeks = 8
    val_weeks = 4
    stab_months = 6
elif dev_stage == "Phase 2":
    dev_weeks = 12
    val_weeks = 8
    stab_months = 12
else:
    dev_weeks = 16
    val_weeks = 12
    stab_months = 24

df = fetch_notion_data(DATABASE_ID, NOTION_TOKEN)

if not df.empty:
    st.success(f"🟢 {substance_type} - {dev_stage} 맞춤형 로드맵 생성 준비 완료")
    cat_col = "Method Category" if "Method Category" in df.columns else "Category"

    # --- 마일스톤 간트 차트 생성 ---
    def generate_strategic_gantt(dataframe, start_date, stage, substance):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        sheet = workbook.add_worksheet('CMC_Strategic_Roadmap')
        
        # 스타일
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        milestone_fmt = workbook.add_format({'bg_color': '#FFD966', 'bold': True, 'border': 1, 'align': 'center'})
        dev_fmt = workbook.add_format({'bg_color': '#DEEAF6', 'border': 1})
        val_fmt = workbook.add_format({'bg_color': '#FBE5D6', 'border': 1})
        stab_fmt = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1})
        comparability_fmt = workbook.add_format({'bg_color': '#E1CBFF', 'border': 1}) # 동등성 평가 (보라)

        # 헤더
        cols = ['Category', 'Activity / Milestone', 'Phase']
        for c, name in enumerate(cols):
            sheet.write(0, c, name, header_fmt)
        
        start_dt = datetime.combine(start_date, datetime.min.time())
        for w in range(52): # 1년치 주차 표시
            sheet.write(0, 3 + w, (start_dt + timedelta(weeks=w)).strftime('%y-%m-%d'), header_fmt)
            sheet.set_column(3 + w, 3 + w, 3)

        row_idx = 1
        # [공통 마일스톤: 동등성 평가]
        sheet.write(row_idx, 0, "Common")
        sheet.write(row_idx, 1, "동등성 평가 (Comparability Study)")
        sheet.write(row_idx, 2, "Critical")
        for w in range(4, 8): sheet.write(row_idx, 3 + w, "", comparability_fmt)
        row_idx += 2

        # [분석법별 세부 마일스톤]
        for _, row in dataframe.iterrows():
            m_name = row['Method']
            # 1. 시험법 개발 (Method Development)
            sheet.write(row_idx, 0, row[cat_col])
            sheet.write(row_idx, 1, f"{m_name} 개발")
            for w in range(0, dev_weeks): sheet.write(row_idx, 3 + w, "", dev_fmt)
            row_idx += 1

            # 2. 시험법 검증/동등성 (Qualification/Validation)
            sheet.write(row_idx, 1, f"{m_name} 검증")
            for w in range(dev_weeks, dev_weeks + val_weeks): sheet.write(row_idx, 3 + w, "", val_fmt)
            row_idx += 1

            # 3. 안정성 시험 (Stability)
            if str(row['Stability-indicating']).lower() in ['yes', 'partial']:
                sheet.write(row_idx, 1, f"{m_name} 안정성 시험")
                for w in range(dev_weeks + val_weeks, dev_weeks + val_weeks + (stab_months * 4)):
                    if w < 52: sheet.write(row_idx, 3 + w, "", stab_fmt)
                row_idx += 1
            
            row_idx += 1 # 항목간 간격

        workbook.close()
        return output.getvalue()

    if st.button("📊 마일스톤 간트 차트 (Excel) 추출"):
        excel_bin = generate_strategic_gantt(df, base_date, dev_stage, substance_type)
        st.download_button(
            label="💾 전략 로드맵 다운로드",
            data=excel_bin,
            file_name=f"CMC_Roadmap_{substance_type}_{dev_stage}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )