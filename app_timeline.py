import streamlit as st
import pandas as pd
import requests
import xlsxwriter
from io import BytesIO
from datetime import datetime, timedelta

# --- 페이지 설정 ---
st.set_page_config(page_title="AtheraCLOUD Operation Planner", layout="wide")

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

# --- UI 및 일정 설정 ---
st.title("📊 Tool 2: CMC Operation & Gantt Planner")
st.sidebar.header("📅 Timeline Strategy")
base_date = st.sidebar.date_input("프로젝트 시작일", datetime(2026, 3, 1))

df = fetch_notion_data(DATABASE_ID, NOTION_TOKEN)

if not df.empty:
    st.success("🟢 실시간 노션 데이터 기반 간트 차트 생성 준비 완료")
    cat_col = "Method Category" if "Method Category" in df.columns else "Category"

    # --- 간트 차트 엑셀 생성 함수 ---
    def generate_gantt_excel(dataframe, start_date, category_col):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        sheet = workbook.add_worksheet('Gantt_Chart')
        
        # 스타일 정의
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
        date_header_fmt = workbook.add_format({'bg_color': '#D9EAD3', 'border': 1, 'align': 'center', 'font_size': 9})
        dev_bar_fmt = workbook.add_format({'bg_color': '#5B9BD5', 'border': 1}) # 개발 (파랑)
        qual_bar_fmt = workbook.add_format({'bg_color': '#ED7D31', 'border': 1}) # 적격성 (주황)
        stab_bar_fmt = workbook.add_format({'bg_color': '#70AD47', 'border': 1}) # 안정성 (초록)

        # 1. 왼쪽 데이터 영역 헤더
        cols = ['Category', 'Method', 'Phase']
        for c, name in enumerate(cols):
            sheet.write(0, c, name, header_fmt)
            sheet.set_column(c, c, 15)

        # 2. 오른쪽 날짜 영역 헤더 (주 단위 표시)
        start_dt = datetime.combine(start_date, datetime.min.time())
        for week in range(40): # 약 10개월치 타임라인 표시
            current_week_dt = start_dt + timedelta(weeks=week)
            sheet.write(0, 3 + week, current_week_dt.strftime('%m/%d'), date_header_fmt)
            sheet.set_column(3 + week, 3 + week, 5)

        # 3. 데이터 및 바(Bar) 생성
        current_row = 1
        for _, row in dataframe.iterrows():
            # 각 시험법마다 3개 행(Dev, Qual, Stab) 생성하여 간트 시각화
            m_name = str(row['Method'])
            m_cat = str(row[category_col])
            
            # --- Method Development ---
            sheet.write(current_row, 0, m_cat)
            sheet.write(current_row, 1, m_name)
            sheet.write(current_row, 2, 'Development')
            # 1~6주차 색칠
            for w in range(0, 6): sheet.write(current_row, 3 + w, "", dev_bar_fmt)
            current_row += 1

            # --- Method Qualification ---
            sheet.write(current_row, 2, 'Qualification')
            # 7~10주차 색칠
            for w in range(6, 10): sheet.write(current_row, 3 + w, "", qual_bar_fmt)
            current_row += 1

            # --- Stability (필요 시) ---
            if str(row['Stability-indicating']).lower() in ['yes', 'partial']:
                sheet.write(current_row, 2, 'Stability Study')
                # 11주차부터 끝까지 색칠
                for w in range(10, 40): sheet.write(current_row, 3 + w, "", stab_bar_fmt)
                current_row += 1
            
            # 구분선(빈 행)
            current_row += 1

        workbook.close()
        return output.getvalue()

    if st.button("🚀 간트 차트 엑셀 다운로드"):
        excel_bin = generate_gantt_excel(df, base_date, cat_col)
        st.download_button(
            label="📥 파일 저장",
            data=excel_bin,
            file_name=f"CMC_Gantt_{datetime.now().strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("노션 데이터를 불러올 수 없습니다.")