import streamlit as st
import pandas as pd
import requests

st.set_page_config(page_title="AtheraCLOUD CMC Control Tower", layout="wide")

# 1. secrets.toml에서 노션 토큰과 DB ID 불러오기
try:
    # 절대 이 괄호 안에 ntn_... 토큰을 직접 넣지 마세요! 아래 글자 그대로 유지해야 합니다.
    NOTION_TOKEN = st.secrets["NOTION_TOKEN"]
    DATABASE_ID = st.secrets["NOTION_DB_ID"]
except KeyError:
    st.error("⚠️ `.streamlit/secrets.toml` 파일에 NOTION_TOKEN 또는 NOTION_DB_ID가 설정되지 않았습니다.")
    st.stop()

# 2. Notion API 호출 함수
@st.cache_data(ttl=60) # 60초마다 데이터 갱신 (서버 부하 방지)
def fetch_notion_data(database_id, token):
    url = f"https://api.notion.com/v1/databases/{database_id}/query"
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }
    
    response = requests.post(url, headers=headers)
    
    if response.status_code != 200:
        st.error(f"Notion API 에러: {response.status_code} - {response.text}")
        return pd.DataFrame()
    
    results = response.json().get("results", [])
    data = []
    
    # 노션 JSON 데이터를 파이썬 텍스트로 변환하는 파서
    for page in results:
        props = page.get("properties", {})
        row = {}
        
        for key, val in props.items():
            prop_type = val.get("type")
            if prop_type == "title":
                row[key] = val["title"][0]["plain_text"] if val["title"] else ""
            elif prop_type == "rich_text":
                row[key] = val["rich_text"][0]["plain_text"] if val["rich_text"] else ""
            elif prop_type == "select":
                row[key] = val["select"]["name"] if val["select"] else ""
            elif prop_type == "multi_select":
                row[key] = ", ".join([item["name"] for item in val["multi_select"]])
            elif prop_type == "relation":
                # Relation은 ID만 오기 때문에 임시 처리. 노션에서 'Rollup(텍스트)'로 빼두면 더 좋습니다.
                row[key] = "Relation Linked" if val["relation"] else ""
            else:
                row[key] = str(val.get(prop_type, ""))
                
        data.append(row)
        
    return pd.DataFrame(data)

# --- 메인 UI 영역 ---
st.title("🗺️ Tool 1: CMC Master Roadmap (API 연동 버전)")
st.markdown("노션 `5_Analytical Method DB`와 실시간으로 연동되어 작동하는 마스터 대시보드입니다.")

with st.spinner('노션에서 실시간 데이터를 불러오는 중입니다...'):
    df = fetch_notion_data(DATABASE_ID, NOTION_TOKEN)

if not df.empty:
    st.success("🟢 노션 데이터베이스 실시간 연동 성공!")
    
    # 노션 컬럼명이 정확히 일치해야 합니다.
    try:
        display_cols = ["Attribute", "Method", "Category", "Stability-indicating", "Typical Purpose"]
        # df에 존재하는 컬럼만 필터링하여 에러 방지
        existing_cols = [col for col in display_cols if col in df.columns]
        
        # 탭(Tab) UI로 카테고리별 출력
        if "Category" in df.columns:
            categories = df['Category'].dropna().unique()
            tabs = st.tabs([str(cat) for cat in categories if cat])
            
            for i, cat in enumerate([c for c in categories if c]):
                with tabs[i]:
                    cat_df = df[df['Category'] == cat][existing_cols]
                    st.dataframe(cat_df, use_container_width=True, hide_index=True)
        else:
            st.dataframe(df[existing_cols], use_container_width=True)
            
    except Exception as e:
        st.warning(f"데이터 표출 중 일부 오류 발생 (컬럼명 확인 필요): {e}")
        st.write("로드된 전체 데이터 확인:", df.head())