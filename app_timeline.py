import streamlit as st
import pandas as pd
import requests
import xlsxwriter
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="AtheraCLOUD Operation Planner", layout="wide")

# 1. Notion API 설정 (Secrets에서 호출)
try:
    NOTION_TOKEN = st.secrets["NOTION_TOKEN"]
    DATABASE_ID = st.secrets["NOTION_DB_ID"]
except Exception:
    st.error("⚠️ Streamlit Cloud의 Secrets 설정에 NOTION_TOKEN과 NOTION_DB_ID를 입력해주세요.")
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

# --- 메인 대시보드 UI ---
st.title("📊 Tool 2: CMC Operation & Timeline Planner")
st.markdown("노션 DB 기반 실무 운영 계획 및 엑셀 타임라인 생성기입니다.")

# 사이드바 설정 영역
st.sidebar.header("📅 Timeline Strategy")
base_date = st.sidebar.date_input("프로젝트 착수일", datetime(2026, 3, 1))
target_ind = st.sidebar.date_input("IND 신청 목표일", datetime(2026, 12, 31))

# 데이터 로드
df = fetch_notion_data(DATABASE_ID, NOTION_TOKEN)

if not df.empty:
    st.success("🟢 실시간 노션 데이터 로드 완료")
    
    # 엑셀 생성을 위한 데이터 가공 (Method Category 컬럼 활용)
    cat_col = "Method Category" if "Method Category" in df.columns else "Category"
    
    # 화면에 미리보기 출력
    st.subheader("📋 분석법별 일정 계획 미리보기")
    view_df = df[[cat_col, "Method", "Stability-indicating", "Typical Purpose"]].copy()
    st.dataframe(view_df, use_container_width=True, hide_index=True)

    # --- 엑셀 생성 함수 ---
    def generate_excel_planner(dataframe, start_date, category_col):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        sheet = workbook.add_worksheet('CMC_Master_Timeline')
        
        # 엑셀 스타일 정의
        bold_blue = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
        date_fmt = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1, 'align': 'center'})
        text_fmt = workbook.add_format({'border': 1})
        highlight_fmt = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'})

        # 헤더 섹션
        headers = ['Category', 'Method', 'Dev Start', 'Dev End (6w)', 'Qual End (+4w)', 'Stability Study']
        for col, head in enumerate(headers):
            sheet.write(0, col, head, bold_blue)
            sheet.set_column(col, col, 18)

        # 행 데이터 작성
        for i, (_, row) in enumerate(dataframe.iterrows(), start=1):
            d_start = datetime.combine(start_date, datetime.min.time())
            d_end = d_start + timedelta(weeks=6)
            q_end = d_end + timedelta(weeks=4)
            
            sheet.write(i, 0, str(row[category_col]), text_fmt)
            sheet.write(i, 1, str(row['Method']), text_fmt)
            sheet.write(i, 2, d_start, date_fmt)
            sheet.write(i, 3, d_end, date_fmt)
            sheet.write(i, 4, q_end, date_fmt)
            
            # 안정성 시험 여부 판단
            if str(row['Stability-indicating']).lower() in ['yes', 'partial']:
                sheet.write(i, 5, "Targeted (Start at Qual End)", highlight_fmt)
            else:
                sheet.write(i, 5, "N/A", text_fmt)

        workbook.close()
        return output.getvalue()

    # 다운로드 버튼
    st.markdown("---")
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("🚀 엑셀 플래너 생성"):
            excel_bin = generate_excel_planner(df, base_date, cat_col)
            st.download_button(
                label="📥 엑셀 파일 다운로드",
                data=excel_bin,
                file_name=f"CMC_Timeline_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with col2:
        st.info("버튼을 누르면 각 시험법의 표준 리드타임(개발 6주, 적격성 4주)이 적용된 마스터 플랜이 생성됩니다.")

else:
    st.warning("노션에서 데이터를 불러올 수 없습니다. API 설정을 확인하세요.")