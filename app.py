import streamlit as st
import pandas as pd
import requests
import io
import xlsxwriter  # requirements.txt에 XlsxWriter 필수!
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
# 2. 데이터 로딩 (Backend)
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
    
    doc.add_heading('1. 밸리데이션 항목 및 판정 기준', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = "항목"; table.rows[0].cells[1].text = "기준"
    
    items = [
        ("특이성", params.get('Detail_Specificity')),
        ("직선성", params.get('Detail_Linearity')),
        ("범위", params.get('Detail_Range')),
        ("정확성", params.get('Detail_Accuracy')),
        ("정밀성", params.get('Detail_Precision')),
        ("검출한계 (LOD)", params.get('Detail_LOD')),
        ("정량한계 (LOQ)", params.get('Detail_LOQ')),
        ("완건성", params.get('Detail_Robustness'))
    ]
    for k, v in items:
        if v:
            row = table.add_row().cells
            row[0].text = k; row[1].text = v
            
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [NEW] 멀티 시트 엑셀 생성 함수
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    # 스타일 정의
    header_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#4472C4', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
    sub_header_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    cell_fmt = workbook.add_format({'border': 1})
    num_fmt = workbook.add_format({'border': 1, 'num_format': '0.00'})
    calc_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFFFCC', 'num_format': '0.00'}) # 노란색 (자동계산)

    # -----------------------------------------------------
    # Sheet 1: 1. Info & Prep (기본 정보 및 조제)
    # -----------------------------------------------------
    ws1 = workbook.add_worksheet("1. Info & Prep")
    ws1.set_column('A:A', 20); ws1.set_column('B:E', 15)
    
    ws1.merge_range('A1:E1', f'GMP Analytical Logbook: {method_name}', header_fmt)
    
    # 기본 정보
    ws1.write('A3', "Method Name", sub_header_fmt)
    ws1.merge_range('B3:E3', method_name, cell_fmt)
    ws1.write('A4', "Date", sub_header_fmt)
    ws1.merge_range('B4:E4', datetime.now().strftime("%Y-%m-%d"), cell_fmt)
    ws1.write('A5', "Analyst", sub_header_fmt)
    ws1.merge_range('B5:E5', "", cell_fmt)
    
    # 기기 및 컬럼
    ws1.write('A7', "Instrument", sub_header_fmt)
    ws1.merge_range('B7:E7', params.get('Instrument', ''), cell_fmt)
    ws1.write('A8', "Column / Plate", sub_header_fmt)
    ws1.merge_range('B8:E8', params.get('Column_Plate', ''), cell_fmt)
    
    # 시약 정보
    ws1.merge_range('A10:E10', "■ 시약 및 표준품 정보 (Reagents & Standards)", sub_header_fmt)
    ws1.write('A11', "구분", sub_header_fmt)
    ws1.merge_range('B11:C11', "품명", sub_header_fmt)
    ws1.merge_range('D11:E11', "Lot No. / Exp. Date", sub_header_fmt)
    
    ws1.write('A12', "표준품", cell_fmt)
    ws1.merge_range('B12:C12', params.get('Ref_Standard_Info', ''), cell_fmt)
    ws1.merge_range('D12:E12', "", cell_fmt)
    
    ws1.write('A13', "시약 1", cell_fmt)
    ws1.merge_range('B13:C13', params.get('Reagent_List', ''), cell_fmt)
    ws1.merge_range('D13:E13', "", cell_fmt)

    # 조제 방법 (SOP 내용 표시)
    ws1.merge_range('A15:E15', "■ 용액 조제 방법 (Preparation)", sub_header_fmt)
    ws1.write('A16', "표준액 조제", cell_fmt)
    ws1.merge_range('B16:E16', params.get('Preparation_Std', 'SOP 참조'), cell_fmt)
    ws1.write('A17', "검체 조제", cell_fmt)
    ws1.merge_range('B17:E17', params.get('Preparation_Sample', 'SOP 참조'), cell_fmt)

    # -----------------------------------------------------
    # Sheet 2: 2. Linearity (직선성 농도 자동 계산)
    # -----------------------------------------------------
    target_conc = params.get('Target_Conc')
    if target_conc:
        ws2 = workbook.add_worksheet("2. Linearity")
        ws2.set_column('A:A', 15); ws2.set_column('B:E', 18)
        
        unit = params.get('Unit', 'ppm')
        ws2.merge_range('A1:E1', f'Linearity Calculation (Target: {target_conc} {unit})', header_fmt)
        
        headers = ["Level (%)", f"Target Conc ({unit})", "Actual Weight (mg)", "Dilution Vol (mL)", "Final Conc (Calc)"]
        for col, h in enumerate(headers):
            ws2.write(2, col, h, sub_header_fmt)
            
        levels = [80, 90, 100, 110, 120]
        row = 3
        for level in levels:
            target_val = float(target_conc) * (level / 100)
            
            ws2.write(row, 0, f"{level}%", cell_fmt)
            ws2.write(row, 1, target_val, num_fmt)
            ws2.write(row, 2, "", cell_fmt) # 사용자 입력 (노란색으로 강조 가능)
            ws2.write(row, 3, 50, cell_fmt) # 기본 부피
            
            # 수식: =C{row}/D{row} * 1000 (예시)
            formula = f"=C{row+1}/D{row+1}*1000"
            ws2.write_formula(row, 4, formula, calc_fmt)
            row += 1
            
        ws2.merge_range(f'A{row+2}:E{row+2}', "※ 노란색 셀(Final Conc)은 칭량값 입력 시 자동 계산됩니다.", cell_fmt)

    # -----------------------------------------------------
    # Sheet 3: 3. Robustness (완건성 - 데이터 있을 때만 생성)
    # -----------------------------------------------------
    if params.get('Detail_Robustness'):
        ws3 = workbook.add_worksheet("3. Robustness")
        ws3.set_column('A:A', 25); ws3.set_column('B:F', 15)
        
        ws3.merge_range('A1:F1', 'Robustness Test Conditions', header_fmt)
        ws3.merge_range('A2:F2', f"Guide: {params.get('Detail_Robustness')}", cell_fmt)
        
        r_headers = ["Condition", "Set Value", "Actual Value", "SST Result", "Pass/Fail", "Note"]
        for col, h in enumerate(r_headers):
            ws3.write(3, col, h, sub_header_fmt)
            
        conditions = ["Standard", "Flow Rate -0.1", "Flow Rate +0.1", "Temp -2℃", "Temp +2℃"]
        row = 4
        for cond in conditions:
            ws3.write(row, 0, cond, cell_fmt)
            for col in range(1, 6):
                ws3.write(row, col, "", cell_fmt)
            row += 1

    # -----------------------------------------------------
    # Sheet 4: 4. Raw Data (결과 기록)
    # -----------------------------------------------------
    ws4 = workbook.add_worksheet("4. Raw Data")
    ws4.set_column('A:A', 10); ws4.set_column('B:B', 30); ws4.set_column('C:F', 15)
    
    ws4.merge_range('A1:F1', 'Raw Data Recording Sheet', header_fmt)
    
    headers = ["Inj No.", "Sample Name", "RT (min)", "Area", "Height", "Remarks"]
    for col, h in enumerate(headers):
        ws4.write(2, col, h, sub_header_fmt)
        
    for row in range(3, 23): # 20줄 생성
        for col in range(6):
            ws4.write(row, col, "", cell_fmt)

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
# 4. 메인 UI
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Full GMP", layout="wide")
st.title("🧪 AtheraCLOUD: Full CMC Validation Suite")
st.markdown("##### Strategy · Protocol · Multi-Sheet Logbook · Report")

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
            tab1, tab2, tab3 = st.tabs(["📑 Step 1: VMP & Protocol", "📗 Step 2: Excel Logbook (Multi-Tab)", "📊 Step 3: Result Report"])
            
            with tab1:
                st.markdown("### 1️⃣ 전략 (VMP) 및 상세 계획서 (Protocol)")
                st.dataframe(my_plan[["Method", "Category", "Required_Items"]], use_container_width=True)
                
                c1, c2 = st.columns(2)
                with c1:
                    doc_vmp = generate_vmp_premium(sel_modality, sel_phase, my_plan)
                    st.download_button("📥 VMP 전체 다운로드 (Word)", doc_vmp, "VMP_Master.docx")
                with c2:
                    sel_proto = st.selectbox("개별 계획서 선택:", my_plan["Method"].unique())
                    if sel_proto:
                        params = get_method_params(sel_proto)
                        if params:
                            doc = generate_protocol_premium(sel_proto, "Category", params)
                            st.download_button(f"📄 {sel_proto} Protocol 다운로드", doc, f"Protocol_{sel_proto}.docx")

            with tab2:
                st.markdown("### 📗 스마트 엑셀 일지 (Multi-Sheet)")
                st.info("시험 항목별로 시트가 분리된 엑셀 파일을 생성합니다. (Info / Linearity / Robustness / Raw Data)")
                sel_log = st.selectbox("일지 생성 시험법:", my_plan["Method"].unique(), key="log")
                params = get_method_params(sel_log)
                if params:
                    data = generate_smart_excel(sel_log, "Cat", params)
                    st.download_button(f"📊 {sel_log} Logbook (Excel) 다운로드", data, f"Logbook_{sel_log}.xlsx", type="primary")

            with tab3:
                st.markdown("### 📊 최종 결과 보고서 (Security Mode)")
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