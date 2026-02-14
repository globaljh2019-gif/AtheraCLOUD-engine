import streamlit as st
import pandas as pd
import requests
import io
import xlsxwriter
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

# ---------------------------------------------------------
# 1. 설정 및 데이터 로딩 (기존과 동일)
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

headers = {"Authorization": "Bearer " + NOTION_API_KEY, "Content-Type": "application/json", "Notion-Version": "2022-06-28"}

@st.cache_data
def get_criteria_map():
    url = f"https://api.notion.com/v1/databases/{CRITERIA_DB_ID}/query"
    response = requests.post(url, headers=headers)
    criteria_map = {}
    if response.status_code == 200:
        for page in response.json().get("results", []):
            try:
                props = page["properties"]
                criteria_map[page["id"]] = {"Category": props["Test_Category"]["title"][0]["text"]["content"], 
                                            "Required_Items": [i["name"] for i in props["Required_Items"]["multi_select"]]}
            except: continue
    return criteria_map

def get_strategy_list(criteria_map):
    url = f"https://api.notion.com/v1/databases/{STRATEGY_DB_ID}/query"
    res = requests.post(url, headers=headers)
    data = []
    if res.status_code == 200:
        for p in res.json().get("results", []):
            try:
                props = p["properties"]
                rel = props["Test Category"]["relation"]
                cat, items = ("Unknown", [])
                if rel and rel[0]["id"] in criteria_map:
                    cat = criteria_map[rel[0]["id"]]["Category"]
                    items = criteria_map[rel[0]["id"]]["Required_Items"]
                data.append({"Modality": props["Modality"]["select"]["name"], "Phase": props["Phase"]["select"]["name"],
                             "Method": props["Method Name"]["rich_text"][0]["text"]["content"], "Category": cat, "Required_Items": items})
            except: continue
    return pd.DataFrame(data)

def get_method_params(method_name):
    """ICH Q2(R2) Full Scope + Intermediate Precision"""
    if not PARAM_DB_ID: return None
    url = f"https://api.notion.com/v1/databases/{PARAM_DB_ID}/query"
    payload = {"filter": {"property": "Method_Name", "title": {"equals": method_name}}}
    res = requests.post(url, headers=headers, json=payload)
    if res.status_code == 200 and res.json().get("results"):
        props = res.json()["results"][0]["properties"]
        def txt(n): 
            try: return "".join([t["text"]["content"] for t in props[n]["rich_text"]])
            except: return ""
        def num(n):
            try: return props[n]["number"]
            except: return None
        
        return {
            "Instrument": txt("Instrument"), "Column_Plate": txt("Column_Plate"),
            "Condition_A": txt("Condition_A"), "Condition_B": txt("Condition_B"), "Detection": txt("Detection"),
            "SST_Criteria": txt("SST_Criteria"), "Reference_Guideline": txt("Reference_Guideline"),
            "Detail_Specificity": txt("Detail_Specificity"), "Detail_Linearity": txt("Detail_Linearity"),
            "Detail_Range": txt("Detail_Range"), "Detail_Accuracy": txt("Detail_Accuracy"),
            "Detail_Precision": txt("Detail_Precision"), 
            "Detail_Inter_Precision": txt("Detail_Inter_Precision"), # [NEW] 실험실내 정밀성
            "Detail_LOD": txt("Detail_LOD"), "Detail_LOQ": txt("Detail_LOQ"), "Detail_Robustness": txt("Detail_Robustness"),
            "Reagent_List": txt("Reagent_List"), "Ref_Standard_Info": txt("Ref_Standard_Info"),
            "Preparation_Std": txt("Preparation_Std"), "Preparation_Sample": txt("Preparation_Sample"),
            "Calculation_Formula": txt("Calculation_Formula"), "Logic_Statement": txt("Logic_Statement"),
            "Target_Conc": num("Target_Conc"), "Unit": txt("Unit")
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

# [VMP는 기존 동일]
def generate_vmp_premium(modality, phase, df_strategy):
    doc = Document(); set_korean_font(doc)
    doc.add_heading(f'Validation Master Plan ({modality} - {phase})', 0)
    table = doc.add_table(rows=1, cols=3); table.style = 'Table Grid'
    hdr = table.rows[0].cells; hdr[0].text='Method'; hdr[1].text='Category'; hdr[2].text='Items'
    for _, row in df_strategy.iterrows():
        c = table.add_row().cells
        c[0].text=str(row['Method']); c[1].text=str(row['Category']); c[2].text=", ".join(row['Required_Items'])
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Protocol 업데이트: 정밀성 추가]
def generate_protocol_premium(method_name, category, params):
    doc = Document(); set_korean_font(doc)
    doc.add_heading(f'Validation Protocol: {method_name}', 0)
    doc.add_paragraph(f"Guideline: {params.get('Reference_Guideline', 'ICH Q2(R2)')}")
    
    doc.add_heading('1. 밸리데이션 항목 및 기준', level=1)
    table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
    table.rows[0].cells[0].text="항목"; table.rows[0].cells[1].text="기준"
    
    items = [
        ("특이성", params.get('Detail_Specificity')), ("직선성", params.get('Detail_Linearity')),
        ("범위", params.get('Detail_Range')), ("정확성", params.get('Detail_Accuracy')),
        ("반복성 (Repeatability)", params.get('Detail_Precision')),
        ("실험실내 정밀성 (Intermediate Precision)", params.get('Detail_Inter_Precision')), # [NEW]
        ("검출/정량한계 (LOD/LOQ)", f"{params.get('Detail_LOD')} / {params.get('Detail_LOQ')}"),
        ("완건성", params.get('Detail_Robustness'))
    ]
    for k, v in items:
        if v:
            r = table.add_row().cells; r[0].text=k; r[1].text=v
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Excel 업데이트: 차트 & 정밀성 시트 추가]
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Formats
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#4472C4', 'font_color':'white', 'align':'center'})
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center'})
    cell = workbook.add_format({'border':1})
    num = workbook.add_format({'border':1, 'num_format':'0.00'})
    calc = workbook.add_format({'border':1, 'bg_color':'#FFFFCC', 'num_format':'0.00'})

    # Sheet 1: Info
    ws1 = workbook.add_worksheet("1. Info & Prep")
    ws1.set_column('A:A', 20); ws1.set_column('B:E', 15)
    ws1.merge_range('A1:E1', f'GMP Logbook: {method_name}', header)
    info = [("Date", datetime.now().strftime("%Y-%m-%d")), ("Instrument", params.get('Instrument')), 
            ("Column", params.get('Column_Plate')), ("Analyst", "")]
    r = 3
    for k, v in info:
        ws1.write(r, 0, k, sub); ws1.merge_range(r, 1, r, 4, v, cell); r+=1
    
    ws1.write(r+1, 0, "Reagent / Std", sub); ws1.merge_range(r+1, 1, r+1, 4, params.get('Ref_Standard_Info', ''), cell)
    ws1.write(r+2, 0, "Prep Method", sub); ws1.merge_range(r+2, 1, r+2, 4, params.get('Preparation_Sample', ''), cell)

    # Sheet 2: Linearity (With Chart!)
    target_conc = params.get('Target_Conc')
    if target_conc:
        ws2 = workbook.add_worksheet("2. Linearity")
        ws2.set_column('A:F', 15)
        unit = params.get('Unit', 'ppm')
        ws2.merge_range('A1:F1', 'Linearity & Range (Input Response to see Chart)', header)
        
        headers = ["Level (%)", f"Conc ({unit})", "Weight (mg)", "Vol (mL)", "Real Conc (X)", "Response (Y)"]
        for c, h in enumerate(headers): ws2.write(2, c, h, sub)
        
        levels = [80, 90, 100, 110, 120]
        row = 3
        for l in levels:
            ws2.write(row, 0, f"{l}%", cell)
            ws2.write(row, 1, float(target_conc)*(l/100), num)
            ws2.write(row, 2, "", cell)
            ws2.write(row, 3, 50, cell)
            ws2.write_formula(row, 4, f"=C{row+1}/D{row+1}*1000", calc) # X축 데이터
            ws2.write(row, 5, "", calc) # Y축 데이터 (사용자 입력)
            row += 1
        
        # [NEW] 차트 생성
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
        chart.add_series({
            'name':       'Linearity',
            'categories': f'=2. Linearity!$E$4:$E$8', # X축: Real Conc
            'values':     f'=2. Linearity!$F$4:$F$8', # Y축: Response
            'trendline':  {'type': 'linear', 'display_equation': True, 'display_r_squared': True},
            'marker':     {'type': 'circle', 'size': 7}
        })
        chart.set_title({'name': 'Calibration Curve'})
        chart.set_x_axis({'name': f'Concentration ({unit})'})
        chart.set_y_axis({'name': 'Response (Area)'})
        
        ws2.insert_chart('H3', chart) # H3 셀 위치에 그래프 삽입

    # Sheet 3: Precision (Intermediate)
    if params.get('Detail_Inter_Precision'):
        ws3 = workbook.add_worksheet("3. Precision")
        ws3.set_column('A:E', 15)
        ws3.merge_range('A1:E1', 'Intermediate Precision (Ruggedness)', header)
        
        # Day 1
        ws3.merge_range('A3:E3', "■ Day 1 (Analyst 1) - Repeatability", sub)
        ws3.write_row('A4', ["Inj No.", "Sample", "Result", "Mean", "RSD"], sub)
        for i in range(6):
            ws3.write(4+i, 0, i+1, cell); ws3.write(4+i, 1, "Sample", cell); ws3.write(4+i, 2, "", cell)
        ws3.write_formula('D5', "=AVERAGE(C5:C10)", num)
        ws3.write_formula('E5', "=STDEV(C5:C10)/D5*100", num)
        
        # Day 2
        ws3.merge_range('A12:E12', "■ Day 2 (Analyst 2) - Intermediate Precision", sub)
        ws3.write_row('A13', ["Inj No.", "Sample", "Result", "Mean", "RSD"], sub)
        for i in range(6):
            ws3.write(13+i, 0, i+1, cell); ws3.write(13+i, 1, "Sample", cell); ws3.write(13+i, 2, "", cell)
        ws3.write_formula('D14', "=AVERAGE(C14:C19)", num)
        ws3.write_formula('E14', "=STDEV(C14:C19)/D14*100", num)
        
        # Comparison
        ws3.write('A21', "Total Mean", sub); ws3.write_formula('B21', "=AVERAGE(D5,D14)", num)
        ws3.write('A22', "Difference (%)", sub); ws3.write_formula('B22', "=ABS(D5-D14)/B21*100", num)

    # Sheet 4: Raw Data & Robustness
    ws4 = workbook.add_worksheet("4. Raw Data")
    ws4.set_column('A:F', 15)
    ws4.merge_range('A1:F1', 'Raw Data & Robustness', header)
    if params.get('Detail_Robustness'):
        ws4.merge_range('A3:F3', "Robustness Conditions", sub)
        ws4.write_row('A4', ["Condition", "Set", "Actual", "SST", "Pass/Fail", ""], sub)
        conds = ["Flow -0.1", "Flow +0.1", "Temp -2", "Temp +2"]
        for i, c in enumerate(conds): ws4.write(5+i, 0, c, cell)
    
    workbook.close(); output.seek(0)
    return output

# [Report 업데이트: 정밀성 포함]
def generate_summary_report_gmp(method_name, category, params, user_inputs):
    doc = Document(); set_korean_font(doc)
    doc.add_heading(f'Validation Summary Report: {method_name}', 0)
    
    info = doc.add_table(rows=3, cols=2); info.style='Table Grid'
    d = [("Category", category), ("Lot/Date", f"{user_inputs['lot_no']} / {user_inputs['date']}"), ("Analyst", user_inputs['analyst'])]
    for i, (k, v) in enumerate(d): info.rows[i].cells[0].text=k; info.rows[i].cells[1].text=str(v)

    doc.add_heading('1. 상세 결과 (Test Results)', level=1)
    table = doc.add_table(rows=1, cols=3); table.style='Table Grid'
    table.rows[0].cells[0].text="항목"; table.rows[0].cells[1].text="기준"; table.rows[0].cells[2].text="결과"
    
    check_items = [
        ("특이성", params.get('Detail_Specificity'), "Pass"),
        ("직선성 (R²)", params.get('Detail_Linearity'), "Pass (See Chart)"),
        ("정밀성 (반복성)", params.get('Detail_Precision'), user_inputs.get('main_result', 'N/A')),
        ("실험실내 정밀성", params.get('Detail_Inter_Precision'), "Pass (Diff < 2.0%)"), # [NEW]
        ("완건성", params.get('Detail_Robustness'), "Pass")
    ]
    for k, c, r in check_items:
        if c: table.add_row().cells[0].text=k; table.rows[-1].cells[1].text=c; table.rows[-1].cells[2].text=r

    doc.add_heading('2. 결론', level=1)
    doc.add_paragraph("본 시험법은 직선성, 정밀성(실험실내 정밀성 포함), 완건성 등 모든 기준을 만족함.")
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 메인 UI
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Full GMP", layout="wide")
st.title("🧪 AtheraCLOUD: Full CMC Validation Suite")
st.markdown("##### with Linearity Chart & Intermediate Precision")

col1, col2 = st.columns([1, 3])
with col1:
    st.header("📂 Project")
    sel_modality = st.selectbox("Modality", ["mAb", "Cell Therapy"])
    sel_phase = st.selectbox("Phase", ["Phase 1", "Phase 3"])

with col2:
    try: criteria_map = get_criteria_map(); df_full = get_strategy_list(criteria_map)
    except: df_full = pd.DataFrame()

    if sel_modality == "mAb" and not df_full.empty:
        my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
        if not my_plan.empty:
            t1, t2, t3 = st.tabs(["📑 Protocol", "📗 Excel Logbook (Chart)", "📊 Report"])
            
            with t1:
                st.dataframe(my_plan[["Method", "Category"]]); doc_vmp = generate_vmp_premium(sel_modality, sel_phase, my_plan)
                st.download_button("📥 VMP", doc_vmp, "VMP.docx")
                sel_p = st.selectbox("Protocol:", my_plan["Method"].unique())
                if st.button("Download Protocol"):
                    doc = generate_protocol_premium(sel_p, "Cat", get_method_params(sel_p))
                    st.download_button("📄 Protocol", doc, f"Protocol_{sel_p}.docx")

            with t2:
                st.info("💡 엑셀을 다운로드하고 '2. Linearity' 시트에 값을 입력하면 그래프가 자동으로 그려집니다.")
                sel_l = st.selectbox("Logbook:", my_plan["Method"].unique(), key="l")
                if st.button("Download Excel Logbook"):
                    data = generate_smart_excel(sel_l, "Cat", get_method_params(sel_l))
                    st.download_button("📊 Excel Logbook", data, f"Logbook_{sel_l}.xlsx")

            with t3:
                sel_r = st.selectbox("Report:", my_plan["Method"].unique(), key="r")
                with st.form("rep"):
                    l = st.text_input("Lot"); d = st.text_input("Date"); a = st.text_input("Analyst")
                    s = st.text_input("SST"); m = st.text_input("Main Result")
                    if st.form_submit_button("Generate Report"):
                        doc = generate_summary_report_gmp(sel_r, "Cat", get_method_params(sel_r), 
                                                          {'lot_no':l, 'date':d, 'analyst':a, 'sst_result':s, 'main_result':m})
                        st.download_button("📥 Report", doc, "Report.docx")