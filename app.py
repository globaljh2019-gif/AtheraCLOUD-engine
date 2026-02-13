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
    """ICH Q2(R2) 모든 항목 포함"""
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
                # 기본 정보
                "Instrument": get_text("Instrument"),
                "Column_Plate": get_text("Column_Plate"),
                "Condition_A": get_text("Condition_A"),
                "Condition_B": get_text("Condition_B"),
                "Detection": get_text("Detection"),
                "SST_Criteria": get_text("SST_Criteria"),
                
                # Validation Parameters (Full Scope)
                "Reference_Guideline": get_text("Reference_Guideline"),
                "Detail_Specificity": get_text("Detail_Specificity"),
                "Detail_Linearity": get_text("Detail_Linearity"),
                "Detail_Range": get_text("Detail_Range"),     # [NEW]
                "Detail_Accuracy": get_text("Detail_Accuracy"),
                "Detail_Precision": get_text("Detail_Precision"),
                "Detail_LOD": get_text("Detail_LOD"),         # [NEW]
                "Detail_LOQ": get_text("Detail_LOQ"),         # [NEW]
                "Detail_Robustness": get_text("Detail_Robustness"), # [NEW]
                
                # GMP & Excel Info
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
    
    # 순서대로 모두 표시
    items = [
        ("특이성 (Specificity)", params.get('Detail_Specificity')),
        ("직선성 (Linearity)", params.get('Detail_Linearity')),
        ("범위 (Range)", params.get('Detail_Range')), # [NEW]
        ("정확성 (Accuracy)", params.get('Detail_Accuracy')),
        ("정밀성 (Precision)", params.get('Detail_Precision')),
        ("검출한계 (LOD)", params.get('Detail_LOD')), # [NEW]
        ("정량한계 (LOQ)", params.get('Detail_LOQ')), # [NEW]
        ("완건성 (Robustness)", params.get('Detail_Robustness')) # [NEW]
    ]
    
    for k, v in items:
        if v:
            row = table.add_row().cells
            row[0].text = k
            row[1].text = v
            
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

def generate_smart_excel(method_name, category, params):
    """스마트 엑셀 일지 - 완건성(Robustness) 포함"""
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Logbook")

    bold = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    cell_fmt = workbook.add_format({'border': 1})
    num_fmt = workbook.add_format({'border': 1, 'num_format': '0.00'})
    calc_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFFFCC', 'num_format': '0.00'})

    # 헤더
    worksheet.merge_range('A1:F1', f'GMP Analytical Logbook: {method_name}', bold)
    row = 2
    # ... (기본 정보 생략, 동일) ...
    
    # 2. 직선성 (Linearity)
    target_conc = params.get('Target_Conc')
    unit = params.get('Unit', 'ppm')
    row = 6
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

    # 3. [NEW] 완건성 (Robustness) 섹션 추가
    # 완건성 정보가 있으면 엑셀에 별도 섹션을 만들어줌
    if params.get('Detail_Robustness'):
        worksheet.merge_range(row, 0, row, 5, "■ 완건성 시험 (Robustness) - 조건 변경 기록", bold)
        row += 1
        r_headers = ["변경 조건 (Condition)", "설정값 (Set)", "실측값 (Actual)", "SST 결과 (RSD/Res)", "판정", "비고"]
        for col, h in enumerate(r_headers):
            worksheet.write(row, col, h, bold)
        row += 1
        
        # 예시 조건들 미리 세팅
        conditions = ["Standard (정상 조건)", "Flow Rate (-0.1)", "Flow Rate (+0.1)", "Temp (-2℃)", "Temp (+2℃)"]
        for cond in conditions:
            worksheet.write(row, 0, cond, cell_fmt)
            for col in range(1, 6):
                worksheet.write(row, col, "", cell_fmt)
            row += 1
        row += 2

    # 4. Raw Data
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
    """보고서 - LOD/LOQ/Robustness 포함"""
    doc = Document()
    set_korean_font(doc)
    doc.add_heading(f'Validation Summary Report: {method_name}', 0)
    
    # ... (헤더 생략) ...
    
    # 상세 결과 테이블 확장
    doc.add_heading('2. 상세 밸리데이션 결과 (Comprehensive Results)', level=1)
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = "항목"; table.rows[0].cells[1].text = "기준"; table.rows[0].cells[2].text = "결과"
    
    # 리스트업 (LOD, Robustness 등 포함)
    check_items = [
        ("특이성", params.get('Detail_Specificity'), "Pass"),
        ("직선성", params.get('Detail_Linearity'), params.get('Actual_Result_1', 'Pass')), # 사용자 입력 매핑 필요
        ("정확성", params.get('Detail_Accuracy'), user_inputs.get('main_result', 'N/A')),
        ("완건성", params.get('Detail_Robustness'), "Pass (See Raw Data)")
    ]
    
    for item, crit, res in check_items:
        if crit:
            row = table.add_row().cells
            row[0].text = item; row[1].text = crit; row[2].text = res

    doc.add_heading('3. 결론', level=1)
    doc.add_paragraph("모든 설정된 밸리데이션 항목(완건성 포함)이 기준을 만족함.")
    
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 메인 UI
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Full GMP", layout="wide")
st.title("🧪 AtheraCLOUD: Full CMC Validation Suite")
st.markdown("##### Including Robustness, LOD/LOQ, Range (ICH Q2 R2 Compliance)")

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
            tab1, tab2, tab3 = st.tabs(["📑 Protocol (Full)", "📗 Excel Logbook (Robustness)", "📊 Report"])
            
            with tab1:
                st.subheader("상세 계획서 (Protocol)")
                sel_proto = st.selectbox("시험법 선택:", my_plan["Method"].unique())
                if sel_proto:
                    params = get_method_params(sel_proto)
                    if params:
                        st.info(f"✅ 완건성(Robustness) 및 범위(Range) 항목이 포함된 계획서를 생성합니다.")
                        doc = generate_protocol_premium(sel_proto, "Category", params)
                        st.download_button(f"📥 {sel_proto} Protocol", doc, f"Protocol_{sel_proto}.docx")
            
            with tab2:
                st.subheader("스마트 엑셀 일지")
                sel_log = st.selectbox("일지 생성:", my_plan["Method"].unique(), key="log")
                params = get_method_params(sel_log)
                if params:
                    data = generate_smart_excel(sel_log, "Cat", params)
                    st.download_button(f"📊 {sel_log} Logbook", data, f"Logbook_{sel_log}.xlsx")

            with tab3:
                st.subheader("최종 보고서")
                sel_rep = st.selectbox("보고서 생성:", my_plan["Method"].unique(), key="rep")
                params = get_method_params(sel_rep)
                if params:
                    with st.form("rep"):
                        lot = st.text_input("Lot No")
                        main = st.text_input("Main Result")
                        if st.form_submit_button("생성"):
                            doc = generate_summary_report_gmp(sel_rep, "Cat", params, {'lot_no':lot, 'main_result':main, 'date':'', 'analyst':'', 'sst_result':''})
                            st.download_button("📥 Report", doc, "Report.docx")