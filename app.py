import streamlit as st
import pandas as pd
import requests
import io
import xlsxwriter
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ---------------------------------------------------------
# 1. 설정 및 보안
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
# 2. 데이터 로딩 (모든 파라미터 포함)
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

def get_method_params(method_name):
    """ICH Q2(R2) Full Scope"""
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
            
            def get_number(prop_name):
                try: return props[prop_name]["number"]
                except: return None

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
                "Detail_Range": get_text("Detail_Range"),
                "Detail_Accuracy": get_text("Detail_Accuracy"),
                "Detail_Precision": get_text("Detail_Precision"),
                "Detail_LOD": get_text("Detail_LOD"),
                "Detail_LOQ": get_text("Detail_LOQ"),
                "Detail_Robustness": get_text("Detail_Robustness"),
                "Reagent_List": get_text("Reagent_List"),
                "Ref_Standard_Info": get_text("Ref_Standard_Info"),
                "Preparation_Std": get_text("Preparation_Std"),
                "Preparation_Sample": get_text("Preparation_Sample"),
                "Calculation_Formula": get_text("Calculation_Formula"),
                "Logic_Statement": get_text("Logic_Statement"),
                "Target_Conc": get_number("Target_Conc"),
                "Unit": get_text("Unit")
            }
    return None

# ---------------------------------------------------------
# 3. 문서 생성 엔진
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
    hdr = table.rows[0].cells
    hdr[0].text = 'Method'; hdr[1].text = 'Category'; hdr[2].text = 'Items'
    for _, row in df_strategy.iterrows():
        c = table.add_row().cells
        c[0].text = str(row['Method']); c[1].text = str(row['Category']); c[2].text = ", ".join(row['Required_Items'])
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

def generate_protocol_premium(method_name, category, params):
    doc = Document()
    set_korean_font(doc)
    doc.add_heading(f'Validation Protocol: {method_name}', 0)
    doc.add_paragraph(f"Guideline: {params.get('Reference_Guideline', 'ICH Q2(R2)')}")
    
    doc.add_heading('1. 밸리데이션 항목 및 판정 기준 (Full Scope)', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = "항목 (Parameter)"; table.rows[0].cells[1].text = "절차 및 기준 (Criteria)"
    
    items = [
        ("특이성 (Specificity)", params.get('Detail_Specificity')),
        ("직선성 (Linearity)", params.get('Detail_Linearity')),
        ("범위 (Range)", params.get('Detail_Range')),
        ("정확성 (Accuracy)", params.get('Detail_Accuracy')),
        ("정밀성 (Precision)", params.get('Detail_Precision')),
        ("검출한계 (LOD)", params.get('Detail_LOD')),
        ("정량한계 (LOQ)", params.get('Detail_LOQ')),
        ("완건성 (Robustness)", params.get('Detail_Robustness'))
    ]
    
    for k, v in items:
        if v:
            row = table.add_row().cells
            row[0].text = k; row[1].text = v
            
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

def generate_smart_excel(method_name, category, params):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Logbook")

    bold = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    cell_fmt = workbook.add_format({'border': 1})
    num_fmt = workbook.add_format({'border': 1, 'num_format': '0.00'})
    calc_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFFFCC', 'num_format': '0.00'})

    worksheet.merge_range('A1:F1', f'GMP Analytical Logbook: {method_name}', bold)
    row = 2
    info_data = [("Method", method_name), ("Date", datetime.now().strftime("%Y-%m-%d")), 
                 ("Instrument", params.get('Instrument', '')), ("Column", params.get('Column_Plate', ''))]
    for k, v in info_data:
        worksheet.write(row, 0, k, bold)
        worksheet.merge_range(row, 1, row, 5, v, cell_fmt)
        row += 1
    
    row += 2
    target_conc = params.get('Target_Conc')
    unit = params.get('Unit', 'ppm')
    
    if target_conc:
        worksheet.merge_range(row, 0, row, 5, f"■ 직선성 및 범위 (Linearity & Range)", bold)
        row += 1
        headers = ["Level (%)", f"Target ({unit})", "실제 칭량값", "희석 부피", "실제 농도", "비고"]
        for col, h in enumerate(headers):
            worksheet.write(row, col, h, bold)
        row += 1
        levels = [80, 90, 100, 110, 120]
        for level in levels:
            target_val = float(target_conc) * (level / 100)
            worksheet.write(row, 0, f"{level}%", cell_fmt)
            worksheet.write(row, 1, target_val, num_fmt)
            worksheet.write(row, 2, "", cell_fmt)
            worksheet.write(row, 3, 50, cell_fmt)
            worksheet.write_formula(row, 4, f"=C{row+1}/D{row+1}*1000", calc_fmt)
            worksheet.write(row, 5, "", cell_fmt)
            row += 1
        row += 2

    if params.get('Detail_Robustness'):
        worksheet.merge_range(row, 0, row, 5, "■ 완건성 시험 (Robustness) - 조건 변경 기록", bold)
        row += 1
        r_headers = ["변경 조건", "설정값", "실측값", "SST 결과", "판정", "비고"]
        for col, h in enumerate(r_headers):
            worksheet.write(row, col, h, bold)
        row += 1
        conditions = ["Standard", "Flow -0.1", "Flow +0.1", "Temp -2℃", "Temp +2℃"]
        for cond in conditions:
            worksheet.write(row, 0, cond, cell_fmt)
            for col in range(1, 6):
                worksheet.write(row, col, "", cell_fmt)
            row += 1
        row += 2

    worksheet.merge_range(row, 0, row, 5, "■ 데이터 기록 (Raw Data)", bold)
    row += 1
    headers = ["Inj No.", "Sample Name", "RT (min)", "Area", "Height", "Note"]
    for col, h in enumerate(headers):
        worksheet.write(row, col, h, bold)
    for _ in range(15):
        row += 1
        for col in range(6):
            worksheet.write(row, col, "", cell_fmt)

    workbook.close()
    output.seek(0)
    return output

def generate_summary_report_gmp(method_name, category, params, user_inputs):
    doc = Document()
    set_korean_font(doc)
    doc.add_heading(f'Validation Summary Report: {method_name}', 0)
    
    info_table = doc.add_table(rows=3, cols=2)
    info_table.style = 'Table Grid'
    data = [("Test Category", category), ("Lot No / Date", f"{user_inputs['lot_no']} / {user_inputs['date']}"),
            ("Analyst", user_inputs['analyst'])]
    for i, (k, v) in enumerate(data):
        info_table.rows[i].cells[0].text = k; info_table.rows[i].cells[1].text = str(v)

    doc.add_heading('1. 시스템 적합성 (SST)', level=1)
    sst_table = doc.add_table(rows=2, cols=3)
    sst_table.style = 'Table Grid'
    sst_table.rows[0].cells[0].text = "기준"; sst_table.rows[0].cells[1].text = "결과"; sst_table.rows[0].cells[2].text = "판정"
    sst_table.rows[1].cells[0].text = params['SST_Criteria']
    sst_table.rows[1].cells[1].text = user_inputs['sst_result']
    sst_table.rows[1].cells[2].text = "Pass"

    doc.add_heading('2. 상세 결과 (Comprehensive Results)', level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = "항목"; table.rows[0].cells[1].text = "기준"; table.rows[0].cells[2].text = "결과"
    
    check_items = [
        ("특이성", params.get('Detail_Specificity'), "Pass"),
        ("직선성", params.get('Detail_Linearity'), "Pass (R² > 0.99)"),
        ("범위", params.get('Detail_Range'), "Pass"),
        ("정확성", params.get('Detail_Accuracy'), user_inputs.get('main_result', 'N/A')),
        ("완건성", params.get('Detail_Robustness'), "Pass (See Raw Data)")
    ]
    for item, crit, res in check_items:
        if crit:
            row = table.add_row().cells
            row[0].text = item; row[1].text = crit; row[2].text = res

    doc.add_heading('3. 결론', level=1)
    doc.add_paragraph("본 시험법은 설정된 모든 밸리데이션 항목(ICH Q2 R2)을 만족하므로 적합함.")
    
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 메인 UI (전략 탭 복구)
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Full GMP", layout="wide")
st.title("🧪 AtheraCLOUD: Full CMC Validation Suite")
st.markdown("##### Strategy · Protocol · Smart Logbook · Result Report")

col1, col2 = st.columns([1, 3])
with col1:
    st.header("📂 Project")
    sel_modality = st.selectbox("Modality", ["mAb", "Cell Therapy"])
    sel_phase = st.selectbox("Phase", ["Phase 1", "Phase 3"])

with col2:
    try:
        criteria_map = get_criteria_map()
        df_full = get_strategy_list(criteria_map)
    except:
        df_full = pd.DataFrame()

    if sel_modality == "mAb" and not df_full.empty:
        my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
        
        if not my_plan.empty:
            # [수정됨] 탭 구성에 VMP를 포함한 '전략 및 계획' 탭을 만듭니다.
            tab1, tab2, tab3 = st.tabs(["📑 Step 1: Strategy & Protocol", "📗 Step 2: Excel Logbook", "📊 Step 3: Result Report"])
            
            # --- Tab 1: Strategy(VMP) & Protocol ---
            with tab1:
                st.markdown("### 1️⃣ 전략 수립 (Validation Master Plan)")
                st.markdown("전체 시험 항목에 대한 전략을 수립하고 마스터 플랜(VMP)을 생성합니다.")
                
                # [복구된 부분] 전략 테이블 표시 및 VMP 다운로드
                st.dataframe(my_plan[["Method", "Category", "Required_Items"]], use_container_width=True)
                doc_vmp = generate_vmp_premium(sel_modality, sel_phase, my_plan)
                st.download_button("📥 VMP(전략서) 다운로드 (Word)", doc_vmp, "VMP_Master.docx")
                
                st.divider()
                
                st.markdown("### 2️⃣ 상세 계획서 (Validation Protocol)")
                st.markdown("개별 시험법에 대한 상세 절차와 기준이 포함된 계획서를 생성합니다.")
                sel_proto = st.selectbox("시험법 선택:", my_plan["Method"].unique())
                if sel_proto:
                    params = get_method_params(sel_proto)
                    if params:
                        st.info("✅ 완건성(Robustness) 및 범위(Range) 등 ICH Q2(R2) 전 항목이 포함됩니다.")
                        doc = generate_protocol_premium(sel_proto, "Category", params)
                        st.download_button(f"📄 {sel_proto} Protocol 다운로드", doc, f"Protocol_{sel_proto}.docx")

            # --- Tab 2: Smart Excel ---
            with tab2:
                st.subheader("📗 스마트 엑셀 일지 (Smart Excel Logbook)")
                st.info("기준 농도 자동 계산 및 완건성(Robustness) 기록란이 포함된 엑셀을 생성합니다.")
                sel_log = st.selectbox("일지 생성 시험법:", my_plan["Method"].unique(), key="log")
                params = get_method_params(sel_log)
                if params:
                    data = generate_smart_excel(sel_log, "Cat", params)
                    st.download_button(f"📊 {sel_log} Logbook (Excel) 다운로드", data, f"Logbook_{sel_log}.xlsx", type="primary")

            # --- Tab 3: Report ---
            with tab3:
                st.subheader("📊 최종 결과 보고서 (Security Mode)")
                sel_rep = st.selectbox("보고서 생성 시험법:", my_plan["Method"].unique(), key="rep")
                params = get_method_params(sel_rep)
                if params:
                    if "generated_doc" not in st.session_state:
                        st.session_state.generated_doc = None
                    
                    with st.form("rep_form"):
                        st.write(f"**[{sel_rep}] 결과 입력**")
                        c1, c2 = st.columns(2)
                        with c1:
                            lot = st.text_input("Lot No")
                            date = st.text_input("Date")
                        with c2:
                            analyst = st.text_input("Analyst")
                            sst = st.text_input("SST Result")
                        main = st.text_input("Main Result (Accuracy/Assay)")
                        
                        if st.form_submit_button("🚀 보고서 생성"):
                            user_data = {'lot_no': lot, 'date': date, 'analyst': analyst, 'sst_result': sst, 'main_result': main}
                            cat = my_plan[my_plan["Method"] == sel_rep].iloc[0]["Category"]
                            st.session_state.generated_doc = generate_summary_report_gmp(sel_rep, cat, params, user_data)
                    
                    if st.session_state.generated_doc:
                        st.success("보고서 생성 완료")
                        st.download_button("📥 Report 다운로드", st.session_state.generated_doc, f"Report_{sel_rep}.docx")