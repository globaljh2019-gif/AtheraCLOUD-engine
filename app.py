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
try:
    NOTION_API_KEY = st.secrets["NOTION_API_KEY"]
    CRITERIA_DB_ID = st.secrets["CRITERIA_DB_ID"]
    STRATEGY_DB_ID = st.secrets["STRATEGY_DB_ID"]
    PARAM_DB_ID = st.secrets.get("PARAM_DB_ID", "") 
except:
    NOTION_API_KEY = ""
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

# [UPGRADE] 상세 파라미터 + 가이드라인 + 세부 절차 가져오기
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
            
            def get_text(prop_name):
                try: 
                    # 텍스트가 여러 덩어리일 경우 합침
                    texts = props[prop_name]["rich_text"]
                    return "".join([t["text"]["content"] for t in texts]) if texts else "정보 없음 (Notion 확인 필요)"
                except: return "정보 없음"
            
            return {
                "Instrument": get_text("Instrument"),
                "Column_Plate": get_text("Column_Plate"),
                "Condition_A": get_text("Condition_A"),
                "Condition_B": get_text("Condition_B"),
                "Detection": get_text("Detection"),
                "SST_Criteria": get_text("SST_Criteria"),
                
                # [NEW] 새로 추가된 항목들
                "Reference_Guideline": get_text("Reference_Guideline"),
                "Detail_Specificity": get_text("Detail_Specificity"),
                "Detail_Linearity": get_text("Detail_Linearity"),
                "Detail_Accuracy": get_text("Detail_Accuracy"),
                "Detail_Precision": get_text("Detail_Precision")
            }
    return None

# ---------------------------------------------------------
# 3. 문서 생성 함수 (VMP & Detail Protocol)
# ---------------------------------------------------------
def set_korean_font(doc):
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    style.font.size = Pt(10)

def generate_vmp_premium(modality, phase, df_strategy):
    doc = Document()
    set_korean_font(doc)
    
    doc.add_heading(f'Validation Master Plan ({modality} - {phase})', 0)
    doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
    
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

# [UPGRADE] 프로토콜 생성 함수 (디테일 강화)
def generate_protocol_premium(method_name, category, params):
    doc = Document()
    set_korean_font(doc)
    
    # 타이틀
    doc.add_heading(f'Validation Protocol: {method_name}', 0)
    p = doc.add_paragraph()
    p.add_run(f"Test Category: {category}").bold = True
    p.add_run(f"\nReference Guideline: {params.get('Reference_Guideline', 'Internal SOP')}")
    
    # 1. 시험 목적
    doc.add_heading('1. 목적 (Objective)', level=1)
    doc.add_paragraph(f"본 문서는 '{method_name}' 시험법이 의약품 품질 관리에 적합함을 과학적으로 입증하기 위한 절차 및 기준을 기술한다.")
    
    # 2. 시험 기기 및 조건
    doc.add_heading('2. 시험 기기 및 조건 (Instruments & Conditions)', level=1)
    
    if params:
        table = doc.add_table(rows=5, cols=2)
        table.style = 'Table Grid'
        data = [
            ("기기 (Instrument)", params['Instrument']),
            ("컬럼/플레이트 (Column)", params['Column_Plate']),
            ("조건 A (Condition)", params['Condition_A']),
            ("조건 B (Condition)", params['Condition_B']),
            ("검출 (Detection)", params['Detection'])
        ]
        for i, (key, val) in enumerate(data):
            cell0 = table.rows[i].cells[0]
            cell1 = table.rows[i].cells[1]
            cell0.text = key
            cell1.text = val
            cell0.paragraphs[0].runs[0].bold = True # 굵게
            
    # 3. 적합성 확인 (SST) - 근거 포함
    doc.add_heading('3. 시스템 적합성 확인 (System Suitability)', level=1)
    doc.add_paragraph("본 시험을 수행하기 전, 아래 기준을 만족해야 한다.")
    
    sst_table = doc.add_table(rows=2, cols=2)
    sst_table.style = 'Table Grid'
    sst_table.rows[0].cells[0].text = "판정 기준 (Criteria)"
    sst_table.rows[0].cells[1].text = "근거 (Reference)"
    sst_table.rows[1].cells[0].text = params['SST_Criteria']
    sst_table.rows[1].cells[1].text = params.get('Reference_Guideline', 'N/A')
    
    # 4. 밸리데이션 상세 수행 계획 (핵심!)
    doc.add_heading('4. 밸리데이션 수행 항목 및 절차', level=1)
    doc.add_paragraph("각 밸리데이션 항목에 대한 상세 절차와 판정 기준은 다음과 같다.")
    
    # 상세 항목 테이블 생성
    val_items = [
        ("특이성 (Specificity)", params.get('Detail_Specificity', 'N/A')),
        ("직선성 (Linearity)", params.get('Detail_Linearity', 'N/A')),
        ("정확성 (Accuracy)", params.get('Detail_Accuracy', 'N/A')),
        ("정밀성 (Precision)", params.get('Detail_Precision', 'N/A')),
    ]
    
    val_table = doc.add_table(rows=1, cols=2)
    val_table.style = 'Table Grid'
    val_table.rows[0].cells[0].text = "항목 (Parameter)"
    val_table.rows[0].cells[1].text = "세부 절차 및 판정 기준 (Procedure & Criteria)"
    
    # 굵게 처리
    val_table.rows[0].cells[0].paragraphs[0].runs[0].bold = True
    val_table.rows[0].cells[1].paragraphs[0].runs[0].bold = True

    for item_name, item_detail in val_items:
        # 내용이 '정보 없음'이 아닐 때만 표에 추가
        if "정보 없음" not in item_detail and item_detail.strip() != "":
            row = val_table.add_row()
            row.cells[0].text = item_name
            row.cells[1].text = item_detail

    doc.add_paragraph("\n\n--------------------------------------------------")
    doc.add_paragraph("Approved By: __________________________  Date: ____________")
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 메인 UI
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Engine", layout="wide")

st.title("🧪 AtheraCLOUD: Validation Protocol Generator")
st.markdown("##### Professional VMP & Detailed Protocol System")

col1, col2 = st.columns([1, 3])

with col1:
    st.header("📂 Project Setup")
    sel_modality = st.selectbox("Modality", ["mAb", "Cell Therapy", "Gene Therapy", "Exosome"])
    sel_phase = st.selectbox("Phase", ["Phase 1", "Phase 3"])
    st.info("💡 **Tip:** 노션에 '근거'와 '세부 절차'를 입력하면 계획서에 자동으로 반영됩니다.")

with col2:
    try:
        criteria_map = get_criteria_map()
        df_full = get_strategy_list(criteria_map)
    except:
        st.error("Notion 연결 오류. API Key를 확인하세요.")
        df_full = pd.DataFrame()

    if sel_modality == "mAb":
        if not df_full.empty:
            my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
            
            if not my_plan.empty:
                tab1, tab2 = st.tabs(["📑 Step 1: VMP (종합)", "🔬 Step 2: Protocol (상세)"])
                
                with tab1:
                    st.success(f"✅ **{sel_modality} {sel_phase}** 전략 수립 완료")
                    st.dataframe(my_plan[["Method", "Category", "Required_Items"]], use_container_width=True)
                    doc_vmp = generate_vmp_premium(sel_modality, sel_phase, my_plan)
                    st.download_button("📄 VMP 다운로드 (Word)", doc_vmp, f"VMP_{sel_modality}.docx")

                with tab2:
                    st.markdown("#### 개별 시험법 상세 계획서 생성")
                    selected_method = st.selectbox("시험법 선택:", my_plan["Method"].unique())
                    
                    if selected_method:
                        row_data = my_plan[my_plan["Method"] == selected_method].iloc[0]
                        params = get_method_params(selected_method)
                        
                        if params:
                            st.info(f"🔍 **{selected_method}** 상세 정보를 불러왔습니다.")
                            with st.expander("미리보기 (Data Preview)"):
                                st.json(params)
                                
                            doc_proto = generate_protocol_premium(selected_method, row_data["Category"], params)
                            st.download_button(
                                label=f"📄 {selected_method} Protocol 다운로드",
                                data=doc_proto,
                                file_name=f"Protocol_{selected_method}.docx",
                                type="primary"
                            )
                        else:
                            st.warning(f"⚠️ '{selected_method}' 데이터가 8번 DB에 없습니다.")
            else:
                st.warning("데이터 없음")
    else:
        st.info("개발 중")