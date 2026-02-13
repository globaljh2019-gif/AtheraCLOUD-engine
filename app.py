import streamlit as st
import pandas as pd
import requests
import io
from datetime import datetime
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------
# 1. 설정 및 보안 (API 키)
# ---------------------------------------------------------
# 주의: 배포 시 Secrets에 PARAM_DB_ID를 꼭 추가해야 합니다!
try:
    NOTION_API_KEY = st.secrets["NOTION_API_KEY"]
    CRITERIA_DB_ID = st.secrets["CRITERIA_DB_ID"]
    STRATEGY_DB_ID = st.secrets["STRATEGY_DB_ID"]
    # 새로 만든 8번 DB ID (없으면 에러 방지 위해 빈 문자열 처리)
    PARAM_DB_ID = st.secrets.get("PARAM_DB_ID", "") 
except:
    # 로컬 테스트용 (Secrets가 없을 때)
    NOTION_API_KEY = "직접_입력_혹은_Secrets_설정_필요"
    CRITERIA_DB_ID = ""
    STRATEGY_DB_ID = ""
    PARAM_DB_ID = ""

headers = {
    "Authorization": "Bearer " + NOTION_API_KEY,
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28"
}

# ---------------------------------------------------------
# 2. 노션 데이터 로딩 함수들
# ---------------------------------------------------------
@st.cache_data
def get_criteria_map():
    url = f"https://api.notion.com/v1/databases/{CRITERIA_DB_ID}/query"
    response = requests.post(url, headers=headers)
    criteria_map = {}
    if response.status_code == 200:
        results = response.json().get("results", [])
        for page in results:
            try:
                page_id = page["id"]
                props = page["properties"]
                cat_name = props["Test_Category"]["title"][0]["text"]["content"]
                req_items = [item["name"] for item in props["Required_Items"]["multi_select"]]
                criteria_map[page_id] = {"Category": cat_name, "Required_Items": req_items}
            except: continue
    return criteria_map

def get_strategy_list(criteria_map):
    url = f"https://api.notion.com/v1/databases/{STRATEGY_DB_ID}/query"
    response = requests.post(url, headers=headers)
    strategy_data = []
    if response.status_code == 200:
        results = response.json().get("results", [])
        for page in results:
            try:
                props = page["properties"]
                modality = props["Modality"]["select"]["name"]
                phase = props["Phase"]["select"]["name"]
                method_name = props["Method Name"]["rich_text"][0]["text"]["content"]
                relation_ids = props["Test Category"]["relation"]
                
                required_items = []
                category_name = "Unknown"
                if relation_ids:
                    rel_id = relation_ids[0]["id"]
                    if rel_id in criteria_map:
                        category_name = criteria_map[rel_id]["Category"]
                        required_items = criteria_map[rel_id]["Required_Items"]
                
                strategy_data.append({
                    "Modality": modality,
                    "Phase": phase,
                    "Method": method_name,
                    "Category": category_name,
                    "Required_Items": required_items
                })
            except: continue
    return pd.DataFrame(strategy_data)

# [NEW] 상세 파라미터 가져오기 (8번 DB)
def get_method_params(method_name):
    if not PARAM_DB_ID: return None
    
    url = f"https://api.notion.com/v1/databases/{PARAM_DB_ID}/query"
    payload = {
        "filter": {
            "property": "Method_Name",
            "title": {"equals": method_name}
        }
    }
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        results = response.json().get("results", [])
        if results:
            props = results[0]["properties"]
            # 텍스트 안전하게 가져오기 헬퍼
            def get_text(prop_name):
                try: return props[prop_name]["rich_text"][0]["text"]["content"]
                except: return "N/A"
            
            return {
                "Instrument": get_text("Instrument"),
                "Column_Plate": get_text("Column_Plate"),
                "Condition_A": get_text("Condition_A"),
                "Condition_B": get_text("Condition_B"),
                "Detection": get_text("Detection"),
                "SST_Criteria": get_text("SST_Criteria")
            }
    return None

# ---------------------------------------------------------
# 3. 문서 생성 함수 (VMP & Protocol)
# ---------------------------------------------------------
def set_korean_font(doc):
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    style.font.size = Pt(11)

def generate_vmp_premium(modality, phase, df_strategy):
    doc = Document()
    set_korean_font(doc)
    
    doc.add_heading(f'Validation Master Plan ({modality} - {phase})', 0)
    doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
    doc.add_paragraph("\n")
    
    # 전략 테이블
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Method'
    hdr_cells[1].text = 'Category'
    hdr_cells[2].text = 'Required Items'
    
    for _, row in df_strategy.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Method'])
        row_cells[1].text = str(row['Category'])
        row_cells[2].text = ", ".join(row['Required_Items'])
        
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# [NEW] 상세 프로토콜 생성 함수
def generate_protocol_premium(method_name, category, params):
    doc = Document()
    set_korean_font(doc)
    
    # 1. 헤더
    doc.add_heading(f'Validation Protocol: {method_name}', 0)
    doc.add_paragraph(f"Test Category: {category}")
    doc.add_paragraph(f"Generated Date: {datetime.now().strftime('%Y-%m-%d')}")
    
    doc.add_heading('1. 목적 (Objective)', level=1)
    doc.add_paragraph(f"본 계획서는 '{method_name}' 시험법이 의약품 품질 관리에 적합함을 입증하기 위한 세부 절차를 기술한다.")
    
    doc.add_heading('2. 시험 기기 및 조건 (Instruments & Conditions)', level=1)
    
    if params:
        table = doc.add_table(rows=5, cols=2)
        table.style = 'Table Grid'
        
        data = [
            ("기기 (Instrument)", params['Instrument']),
            ("컬럼/플레이트 (Column)", params['Column_Plate']),
            ("분석 조건 A (Condition)", params['Condition_A']),
            ("분석 조건 B (Condition)", params['Condition_B']),
            ("검출 (Detection)", params['Detection'])
        ]
        
        for i, (key, val) in enumerate(data):
            table.rows[i].cells[0].text = key
            table.rows[i].cells[1].text = val
    else:
        doc.add_paragraph("⚠️ 상세 파라미터 정보가 노션(8_Method_Parameter_Library)에 없습니다.")

    doc.add_heading('3. 적합성 확인 (System Suitability)', level=1)
    doc.add_paragraph(f"판정 기준: {params['SST_Criteria'] if params else 'TBD'}")

    doc.add_heading('4. 밸리데이션 수행 항목', level=1)
    doc.add_paragraph("본 시험법의 카테고리에 따라 특이성, 직선성, 정밀성 등을 수행한다. (세부 절차 생략)")
    
    doc.add_paragraph("\n\n(End of Document)")
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 메인 UI
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Engine", layout="wide")

st.title("🧪 AtheraCLOUD: Validation Master & Protocol")
st.markdown("##### The First Step to IND Filing: Generate Strategy & Detail Plans")

col1, col2 = st.columns([1, 3])

with col1:
    st.header("📂 Project Setup")
    sel_modality = st.selectbox("Modality", ["mAb", "Cell Therapy", "Gene Therapy", "Exosome"])
    sel_phase = st.selectbox("Phase", ["Phase 1", "Phase 3"])
    
    st.divider()
    st.info("💡 **Tip:** VMP를 먼저 생성한 후, 개별 시험법을 선택하여 상세 계획서(Protocol)를 만드세요.")

with col2:
    # 데이터 로딩
    try:
        criteria_map = get_criteria_map()
        df_full = get_strategy_list(criteria_map)
    except Exception as e:
        st.error("Notion 연결 오류. API Key와 DB ID를 확인하세요.")
        df_full = pd.DataFrame()

    if sel_modality == "mAb":
        if not df_full.empty:
            my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
            
            if not my_plan.empty:
                # 탭으로 기능 분리
                tab1, tab2 = st.tabs(["📑 Step 1: VMP (종합 계획)", "🔬 Step 2: Protocol (상세 계획)"])
                
                # --- Tab 1: VMP ---
                with tab1:
                    st.success(f"✅ **{sel_modality} {sel_phase}** 전략 수립 완료")
                    st.dataframe(
                        my_plan[["Method", "Category", "Required_Items"]],
                        use_container_width=True
                    )
                    
                    doc_vmp = generate_vmp_premium(sel_modality, sel_phase, my_plan)
                    st.download_button(
                        "📄 VMP 다운로드 (Word)",
                        data=doc_vmp,
                        file_name=f"VMP_{sel_modality}_{sel_phase}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                # --- Tab 2: Protocol ---
                with tab2:
                    st.markdown("#### 개별 시험법 상세 계획서 생성")
                    
                    # 시험법 선택 박스
                    selected_method = st.selectbox(
                        "상세 계획서를 작성할 시험법을 선택하세요:",
                        my_plan["Method"].unique()
                    )
                    
                    if selected_method:
                        # 선택된 시험법의 정보 가져오기
                        row_data = my_plan[my_plan["Method"] == selected_method].iloc[0]
                        category = row_data["Category"]
                        
                        # 8번 DB에서 파라미터 조회
                        params = get_method_params(selected_method)
                        
                        if params:
                            st.info(f"🔍 **{selected_method}**의 상세 정보를 노션에서 불러왔습니다.")
                            with st.expander("미리보기 (Parameters)"):
                                st.json(params)
                                
                            # 프로토콜 생성 버튼
                            doc_proto = generate_protocol_premium(selected_method, category, params)
                            st.download_button(
                                label=f"📄 {selected_method} Protocol 다운로드",
                                data=doc_proto,
                                file_name=f"Protocol_{selected_method}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary"
                            )
                        else:
                            st.warning(f"⚠️ '{selected_method}'에 대한 상세 정보가 '8_Method_Parameter_Library'에 없습니다.")
                            st.markdown("노션에 데이터를 추가하거나, Method Name이 정확히 일치하는지 확인해주세요.")
            else:
                st.warning("해당 조건의 전략 데이터가 없습니다.")
    else:
        st.info("개발 중인 Modality입니다.")