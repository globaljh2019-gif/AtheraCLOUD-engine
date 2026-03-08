import streamlit as st
import pandas as pd
import requests
import xlsxwriter
from io import BytesIO
from datetime import datetime, timedelta

# --- 1. Notion API 및 데이터 호출 (기존 로직 유지) ---
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

# --- 2. UI 설정 및 전략 파라미터 ---
st.title("🎯 Tool 2: Strategic CMC Master Scheduler")
st.sidebar.header("🗓️ Project Milestones")

dev_stage = st.sidebar.selectbox("임상 단계", ["Phase 1 (IND)", "Phase 2", "Phase 3 (BLA)"])
base_date = st.sidebar.date_input("CMC 공식 착수일", datetime(2026, 3, 1))
prod_date = st.sidebar.date_input("임상 시료 생산 예정일 (Clinical Batch)", datetime(2026, 8, 1))

# 노션 데이터 로드
try:
    df = fetch_notion_data(st.secrets["NOTION_DB_ID"], st.secrets["NOTION_TOKEN"])
except:
    st.error("Secrets 설정을 확인해주세요.")
    st.stop()

if not df.empty:
    st.success(f"🟢 {dev_stage} 맞춤형 마일스톤 연동 완료")
    
    def generate_master_gantt(dataframe, start_date, clinical_prod_date, stage):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        sheet = workbook.add_worksheet('CMC_Master_Roadmap')
        
        # 스타일 정의
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#203764', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_milestone = workbook.add_format({'bg_color': '#FFD966', 'bold': True, 'border': 1, 'align': 'center'}) # 마일스톤 (황금색)
        fmt_prod = workbook.add_format({'bg_color': '#C6E0B4', 'bold': True, 'border': 1, 'align': 'center'}) # 생산 (연두색)
        fmt_dev = workbook.add_format({'bg_color': '#DEEAF6', 'border': 1}) # 개발
        fmt_val = workbook.add_format({'bg_color': '#FBE5D6', 'border': 1}) # 검증
        fmt_stab = workbook.add_format({'bg_color': '#E2EFDA', 'border': 1}) # 안정성
        
        # 1. 헤더 구성
        headers = ['Category', 'Key Activities & Milestones', 'Dependency']
        for c, h in enumerate(headers):
            sheet.write(0, c, h, fmt_header)
            sheet.set_column(c, c, 20)
        
        start_dt = datetime.combine(start_date, datetime.min.time())
        for w in range(52): # 1년치 주차
            sheet.write(0, 3 + w, (start_dt + timedelta(weeks=w)).strftime('%m/%d'), fmt_header)
            sheet.set_column(3 + w, 3 + w, 4)

        row = 1
        # --- [SECTION 1: Project Management Milestones] ---
        # 1.1 S&P 설정 (모든 분석법 개발 완료 시점 연동)
        sheet.write(row, 1, "★ Milestone: 기준 및 시험방법(S&P) 설정 완료", fmt_milestone)
        for w in range(10, 12): sheet.write(row, 3 + w, "DONE", fmt_milestone)
        row += 1
        
        # 1.2 임상 시료 생산 (사용자 입력 날짜 연동)
        prod_week = int((datetime.combine(clinical_prod_date, datetime.min.time()) - start_dt).days / 7)
        sheet.write(row, 1, "🏭 Clinical Batch Production (Phase Material)", fmt_prod)
        sheet.write(row, 3 + prod_week, "PROD", fmt_prod)
        row += 2

        # --- [SECTION 2: Analytical Methods & Stability] ---
        cat_col = "Method Category" if "Method Category" in df.columns else "Category"
        for _, item in dataframe.iterrows():
            m_name = item['Method']
            # 개발 일정
            sheet.write(row, 0, item[cat_col])
            sheet.write(row, 1, f"{m_name} Method Dev & Optimization")
            for w in range(0, 8): sheet.write(row, 3 + w, "", fmt_dev)
            row += 1
            
            # 검증 일정 (생산 전 완료 목표)
            sheet.write(row, 1, f"{m_name} Qualification/Validation")
            for w in range(8, 12): sheet.write(row, 3 + w, "", fmt_val)
            row += 1
            
            # 안정성 시험 (생산 직후 착수)
            if str(item['Stability-indicating']).lower() in ['yes', 'partial']:
                sheet.write(row, 1, f"{m_name} Stability Study (Long-term/Accel)")
                for w in range(prod_week, prod_week + 24): # 최소 6개월 표시
                    if 3 + w < 55: sheet.write(row, 3 + w, "", fmt_stab)
                row += 1
            row += 1

        workbook.close()
        return output.getvalue()

    if st.button("📊 전략 마스터 로드맵(Excel) 생성"):
        excel_file = generate_master_gantt(df, base_date, prod_date, dev_stage)
        st.download_button("💾 엑셀 다운로드", excel_file, f"CMC_Master_Roadmap_{dev_stage}.xlsx")