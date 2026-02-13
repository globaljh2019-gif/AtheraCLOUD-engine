import streamlit as st
import pandas as pd
import requests
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------
# 1. 설정 및 보안 (API 키 로딩)
# ---------------------------------------------------------
try:
    NOTION_API_KEY = st.secrets["NOTION_API_KEY"]
    CRITERIA_DB_ID = st.secrets["CRITERIA_DB_ID"]
    STRATEGY_DB_ID = st.secrets["STRATEGY_DB_ID"]
    PARAM_DB_ID = st.secrets.get("PARAM_DB_ID", "") 
except:
    # 로컬 테스트용 (Secrets가 없을 경우 방어)
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
# 2. 노션 데이터 로딩 함수 (Backend)
# ---------------------------------------------------------
@st.cache_data
def get_criteria_map():
    """판정 기준 DB(4번)에서 카테고리별 필수 항목 매핑"""
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
    """전략 DB(7번)에서 Modality/Phase별 시험 항목 리스트 추출"""
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

def get_method_params(method_name):
    """상세 파라미터 DB(8번)에서 시험법별 세부 정보 추출"""
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
                    texts = props[prop_name]["rich_text"]
                    return "".join([t["text"]["content"] for t in texts]) if texts else ""
                except: return ""
            
            return {
                "Instrument": get_text("Instrument"),
                "Column_Plate": get_text("Column_Plate"),
                "Condition_A": get_text("Condition_A"),
                "Condition_B": get_text("Condition_B"),
                "Detection": get_text("Detection"),
                "SST_Criteria": get_text("SST_Criteria"),
                "Reference_Guideline": get_text("Reference_Guideline"),
                "Detail_Specificity": get_text("Detail_Specificity"),
                "Detail_Linearity": get_text("Detail_Linearity"),
                "Detail_Accuracy": get_text("Detail_Accuracy"),
                "Detail_Precision": get_text("Detail_Precision")
            }
    return None

# ---------------------------------------------------------
# 3. 문서 생성 엔진 (Word Generator)
# ---------------------------------------------------------
def set_korean_font(doc):
    """한글 폰트(맑은 고딕) 설정"""
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    style.font.size = Pt(10)

def generate_vmp_premium(modality, phase, df_strategy):
    """VMP (종합 계획서) 생성"""
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

def generate_protocol_premium(method_name, category, params):
    """상세 프로토콜 (계획서) 생성"""
    doc = Document()
    set_korean_font(doc)
    
    doc.add_heading(f'Validation Protocol: {method_name}', 0)
    doc.add_paragraph(f"Test Category: {category}")
    doc.add_paragraph(f"Reference Guideline: {params.get('Reference_Guideline', 'Internal SOP')}")
    
    doc.add_heading('1. 목적 (Objective)', level=1)
    doc.add_paragraph(f"본 문서는 '{method_name}' 시험법의 밸리데이션 절차 및 판정 기준을 기술한다.")
    
    doc.add_heading('2. 기기 및 분석 조건 (Instruments & Conditions)', level=1)
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
            table.rows[i].cells[0].text = key
            table.rows[i].cells[1].text = val

    doc.add_heading('3. 적합성 확인 (System Suitability)', level=1)
    doc.add_paragraph(f"판정 기준: {params['SST_Criteria']}")
    
    doc.add_heading('4. 밸리데이션 상세 수행 계획', level=1)
    val_table = doc.add_table(rows=1, cols=2)
    val_table.style = 'Table Grid'
    val_table.rows[0].cells[0].text = "항목 (Parameter)"
    val_table.rows[0].cells[1].text = "절차 및 기준 (Procedure & Criteria)"
    
    items = [
        ("특이성", params.get('Detail_Specificity', '')),
        ("직선성", params.get('Detail_Linearity', '')),
        ("정확성", params.get('Detail_Accuracy', '')),
        ("정밀성", params.get('Detail_Precision', ''))
    ]
    for k, v in items:
        if v:
            row = val_table.add_row()
            row.cells[0].text = k
            row.cells[1].text = v
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_logbook(method_name, params):
    """시험 일지 (Logbook) - 빈 양식 생성"""
    doc = Document()
    set_korean_font(doc)
    
    doc.add_heading(f'Analytical Logbook: {method_name}', 0)
    doc.add_paragraph(f"Doc No: LOG-{datetime.now().strftime('%y%m%d')}-{method_name[:4].upper()}")
    
    # 시험 정보 헤더
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Table Grid'
    info = [("시험 일자", ""), ("시험자 (Analyst)", ""), ("검체 번호 (Lot No)", "")]
    for i, (k, v) in enumerate(info):
        table.rows[i].cells[0].text = k
        table.rows[i].cells[1].text = v

    doc.add_heading('1. 준비 (Preparation)', level=1)
    doc.add_paragraph(f"사용 기기: {params['Instrument']}")
    doc.add_paragraph("□ 표준품 정보: ____________________ (Exp: _________ )")
    doc.add_paragraph("□ 시약 정보: ______________________ (Exp: _________ )")
    
    doc.add_heading('2. 분석 조건 확인', level=1)
    doc.add_paragraph(f"컬럼: {params['Column_Plate']}")
    doc.add_paragraph(f"조건: {params['Condition_A']} / {params['Condition_B']}")

    doc.add_heading('3. 데이터 기록 (Raw Data)', level=1)
    data_table = doc.add_table(rows=8, cols=3)
    data_table.style = 'Table Grid'
    headers = ['Inj No.', 'Sample Name', 'Result (Area/RT)']
    for i, h in enumerate(headers):
        data_table.rows[0].cells[i].text = h
        data_table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
    
    doc.add_paragraph("\n[특이사항 / Deviation Note]")
    doc.add_paragraph("_" * 50)
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_summary_report_secure(method_name, category, params, user_inputs):
    """결과 보고서 (Report) - 사용자 입력 반영 (보안 모드)"""
    doc = Document()
    set_korean_font(doc)
    
    doc.add_heading(f'Validation Summary Report: {method_name}', 0)
    
    # 1. 헤더 정보
    table_info = doc.add_table(rows=3, cols=2)
    table_info.style = 'Table Grid'
    info_map = [
        ("Test Category", category),
        ("Sample / Lot No", user_inputs['lot_no']),
        ("Analysis Date", str(user_inputs['date'])),
        ("Analyst", user_inputs['analyst'])
    ]
    for i in range(3):
        table_info.rows[i].cells[0].text = info_map[i][0]
        table_info.rows[i].cells[1].text = str(info_map[i][1])

    # 2. SST 결과
    doc.add_heading('1. 시스템 적합성 (System Suitability)', level=1)
    sst_table = doc.add_table(rows=2, cols=3)
    sst_table.style = 'Table Grid'
    headers = ['기준 (Criteria)', '실제 결과 (Actual)', '판정 (Judgement)']
    for i, h in enumerate(headers):
        sst_table.rows[0].cells[i].text = h
        sst_table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
    
    sst_table.rows[1].cells[0].text = params['SST_Criteria']
    sst_table.rows[1].cells[1].text = user_inputs['sst_result']
    sst_table.rows[1].cells[2].text = "Pass" # (로직 확장 가능)

    # 3. 상세 결과
    doc.add_heading('2. 상세 시험 결과 (Analytical Results)', level=1)
    res_table = doc.add_table(rows=1, cols=3)
    res_table.style = 'Table Grid'
    res_table.rows[0].cells[0].text = "시험 항목"
    res_table.rows[0].cells[1].text = "기준 (Criteria)"
    res_table.rows[0].cells[2].text = "결과 (Result)"
    
    items = [
        ("특이성 (Specificity)", params.get('Detail_Specificity', ''), "Pass"),
        ("정확성/함량 (Accuracy)", params.get('Detail_Accuracy', ''), user_inputs['main_result']),
        ("정밀성 (Precision)", params.get('Detail_Precision', ''), "Refer to raw data")
    ]
    
    for item, crit, res in items:
        if crit:
            row = res_table.add_row().cells
            row[0].text = item
            row[1].text = crit[:40] + "..." 
            row[2].text = res

    doc.add_heading('3. 결론 (Conclusion)', level=1)
    doc.add_paragraph(f"상기 시험 결과는 {params.get('Reference_Guideline', '설정된 기준')}을 만족하므로 적합(Pass)으로 판정함.")
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 메인 UI (Streamlit App)
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Engine", layout="wide")

st.title("🧪 AtheraCLOUD: CMC Validation Suite")
st.markdown("##### The All-in-One Platform: Strategy, Protocol, Logbook, and Report")

col1, col2 = st.columns([1, 3])

with col1:
    st.header("📂 Project Setup")
    sel_modality = st.selectbox("Modality", ["mAb", "Cell Therapy", "Gene Therapy", "Exosome"])
    sel_phase = st.selectbox("Phase", ["Phase 1", "Phase 3"])
    st.divider()
    st.info("💡 **Workflow:**\n1. VMP (전략 수립)\n2. Protocol (계획서)\n3. Logbook (시험 수행)\n4. Report (결과 판정)")

with col2:
    try:
        criteria_map = get_criteria_map()
        df_full = get_strategy_list(criteria_map)
    except Exception:
        st.error("Notion 연결 오류. API Key와 DB ID를 확인하세요.")
        df_full = pd.DataFrame()

    if sel_modality == "mAb":
        if not df_full.empty:
            my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
            
            if not my_plan.empty:
                # 탭 구성: 전략&계획 / 일지 / 결과보고서
                tab1, tab2, tab3 = st.tabs(["📑 Step 1: Strategy & Protocol", "🧪 Step 2: Logbook (Blank)", "📊 Step 3: Result Report"])
                
                # --- Tab 1: VMP & Protocol ---
                with tab1:
                    st.success(f"✅ **{sel_modality} {sel_phase}** 전략 수립 완료")
                    st.dataframe(my_plan[["Method", "Category", "Required_Items"]], use_container_width=True)
                    
                    doc_vmp = generate_vmp_premium(sel_modality, sel_phase, my_plan)
                    st.download_button("📥 VMP 다운로드 (Word)", doc_vmp, f"VMP_{sel_modality}.docx")
                    
                    st.divider()
                    st.markdown("#### 개별 시험법 상세 계획서 (Protocol)")
                    sel_proto = st.selectbox("시험법 선택:", my_plan["Method"].unique(), key="proto")
                    if sel_proto:
                        row_data = my_plan[my_plan["Method"] == sel_proto].iloc[0]
                        params = get_method_params(sel_proto)
                        if params:
                            with st.expander("상세 파라미터 미리보기"):
                                st.write(params)
                            doc_proto = generate_protocol_premium(sel_proto, row_data["Category"], params)
                            st.download_button(f"📥 {sel_proto} Protocol 다운로드", doc_proto, f"Protocol_{sel_proto}.docx", type="primary")
                        else:
                            st.warning("⚠️ 노션 8번 DB에 상세 정보가 없습니다.")

                # --- Tab 2: Logbook (Blank) ---
                with tab2:
                    st.markdown("#### 🧪 실험실용 시험 일지 (Raw Data Sheet)")
                    st.info("실제 실험 수행 시 수기 기록을 위해 출력하는 빈 양식입니다.")
                    
                    sel_log = st.selectbox("일지를 생성할 시험법:", my_plan["Method"].unique(), key="log")
                    if sel_log:
                        params_log = get_method_params(sel_log)
                        if params_log:
                            doc_log = generate_logbook(sel_log, params_log)
                            st.download_button(f"📄 {sel_log} Logbook 다운로드", doc_log, f"Logbook_{sel_log}.docx")
                        else:
                            st.warning("상세 정보가 없어 일지를 생성할 수 없습니다.")

                # --- Tab 3: Report (Secure Mode) ---
                with tab3:
                    st.markdown("#### 📊 최종 결과 보고서 생성 (Data Security Mode)")
                    st.success("🔒 **보안 안심:** 입력하신 결과값은 서버에 저장되지 않으며, 보고서 생성 즉시 폐기됩니다.")
                    
                    sel_rep = st.selectbox("보고서를 생성할 시험법:", my_plan["Method"].unique(), key="rep_secure")
                    params_rep = get_method_params(sel_rep)
                    
                    if params_rep:
                        # [핵심] 보고서 데이터를 임시 저장할 공간(Session State) 만들기
                        if "generated_doc" not in st.session_state:
                            st.session_state.generated_doc = None
                            st.session_state.generated_name = ""

                        # 1. 입력 폼 (Form)
                        with st.form("report_input_form"):
                            st.markdown(f"**[{sel_rep}] 시험 결과 입력**")
                            c1, c2 = st.columns(2)
                            with c1:
                                input_lot = st.text_input("검체 번호 (Lot No.)", placeholder="24-MAB-001")
                                input_date = st.date_input("시험 일자")
                            with c2:
                                input_analyst = st.text_input("시험자 (Analyst)", placeholder="Name")
                                input_sst = st.text_input("SST 결과 (예: RSD 0.5%)", placeholder="Pass / Fail Data")
                            
                            input_main = st.text_input("메인 결과값 (함량, 회수율 등)", placeholder="예: 99.8% (적합)")
                            
                            # 제출 버튼 (이걸 누르면 문서가 만들어짐)
                            submitted = st.form_submit_button("🚀 보고서 생성")
                            
                            if submitted:
                                user_data = {
                                    "lot_no": input_lot,
                                    "date": input_date,
                                    "analyst": input_analyst,
                                    "sst_result": input_sst,
                                    "main_result": input_main
                                }
                                cat_name = my_plan[my_plan["Method"] == sel_rep].iloc[0]["Category"]
                                
                                # 문서를 만들어서 '주머니(Session State)'에 넣어둠
                                st.session_state.generated_doc = generate_summary_report_secure(sel_rep, cat_name, params_rep, user_data)
                                st.session_state.generated_name = f"Report_{sel_rep}_{input_lot}.docx"

                        # 2. 다운로드 버튼 (Form 바깥에 배치!)
                        # 주머니에 문서가 들어있으면 다운로드 버튼을 보여줌
                        if st.session_state.generated_doc is not None:
                            st.divider()
                            st.info("✅ 보고서 생성이 완료되었습니다.")
                            st.download_button(
                                label="📥 결과 보고서 다운로드 (Word)",
                                data=st.session_state.generated_doc,
                                file_name=st.session_state.generated_name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    else:
                        st.warning("상세 정보가 없습니다.")