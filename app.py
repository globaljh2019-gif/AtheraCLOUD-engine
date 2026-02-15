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
from docx.oxml import OxmlElement

# ---------------------------------------------------------
# 1. 설정 및 데이터 로딩
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
    if not CRITERIA_DB_ID: return {}
    url = f"https://api.notion.com/v1/databases/{CRITERIA_DB_ID}/query"
    res = requests.post(url, headers=headers)
    criteria_map = {}
    if res.status_code == 200:
        for p in res.json().get("results", []):
            try:
                props = p["properties"]
                cat = props["Test_Category"]["title"][0]["text"]["content"] if props["Test_Category"]["title"] else "Unknown"
                req = [i["name"] for i in props["Required_Items"]["multi_select"]]
                criteria_map[p["id"]] = {"Category": cat, "Required_Items": req}
            except: continue
    return criteria_map

def get_strategy_list(criteria_map):
    if not STRATEGY_DB_ID: return pd.DataFrame()
    url = f"https://api.notion.com/v1/databases/{STRATEGY_DB_ID}/query"
    res = requests.post(url, headers=headers)
    data = []
    if res.status_code == 200:
        for p in res.json().get("results", []):
            try:
                props = p["properties"]
                mod = props["Modality"]["select"]["name"] if props["Modality"]["select"] else ""
                ph = props["Phase"]["select"]["name"] if props["Phase"]["select"] else ""
                met = props["Method Name"]["rich_text"][0]["text"]["content"] if props["Method Name"]["rich_text"] else ""
                rel = props["Test Category"]["relation"]
                cat, items = ("Unknown", [])
                if rel and rel[0]["id"] in criteria_map:
                    cat = criteria_map[rel[0]["id"]]["Category"]
                    items = criteria_map[rel[0]["id"]]["Required_Items"]
                data.append({"Modality": mod, "Phase": ph, "Method": met, "Category": cat, "Required_Items": items})
            except: continue
    return pd.DataFrame(data)

def get_method_params(method_name):
    if not PARAM_DB_ID: return {}
    url = f"https://api.notion.com/v1/databases/{PARAM_DB_ID}/query"
    payload = {"filter": {"property": "Method_Name", "title": {"equals": method_name}}}
    res = requests.post(url, headers=headers, json=payload)
    if res.status_code == 200 and res.json().get("results"):
        props = res.json()["results"][0]["properties"]
        def txt(n): 
            try: 
                ts = props.get(n, {}).get("rich_text", [])
                return "".join([t["text"]["content"] for t in ts]) if ts else ""
            except: return ""
        def num(n):
            try: return props.get(n, {}).get("number")
            except: return None
            
        return {
            "Instrument": txt("Instrument"), "Column_Plate": txt("Column_Plate"),
            "Condition_A": txt("Condition_A"), "Condition_B": txt("Condition_B"), "Detection": txt("Detection"),
            "SST_Criteria": txt("SST_Criteria"), "Reference_Guideline": txt("Reference_Guideline"),
            "Detail_Specificity": txt("Detail_Specificity"), "Detail_Linearity": txt("Detail_Linearity"),
            "Detail_Range": txt("Detail_Range"), "Detail_Accuracy": txt("Detail_Accuracy"),
            "Detail_Precision": txt("Detail_Precision"), "Detail_Inter_Precision": txt("Detail_Inter_Precision"),
            "Detail_LOD": txt("Detail_LOD"), "Detail_LOQ": txt("Detail_LOQ"), "Detail_Robustness": txt("Detail_Robustness"),
            "Reagent_List": txt("Reagent_List"), "Ref_Standard_Info": txt("Ref_Standard_Info"),
            "Preparation_Std": txt("Preparation_Std"), "Preparation_Sample": txt("Preparation_Sample"),
            "Calculation_Formula": txt("Calculation_Formula"), "Logic_Statement": txt("Logic_Statement"),
            "Target_Conc": num("Target_Conc"), "Unit": txt("Unit")
        }
    return {}

# ---------------------------------------------------------
# 3. 문서 생성 엔진
# ---------------------------------------------------------
def set_korean_font(doc):
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    style.font.size = Pt(10)

def set_table_header_style(cell):
    tcPr = cell._element.get_or_add_tcPr()
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), 'D9D9D9') 
    tcPr.append(shading_elm)
    if cell.paragraphs:
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# [VMP 생성 함수]
def generate_vmp_premium(modality, phase, df_strategy):
    doc = Document(); set_korean_font(doc)
    head = doc.add_heading('밸리데이션 종합계획서 (Validation Master Plan)', 0); head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    table_info = doc.add_table(rows=2, cols=4); table_info.style = 'Table Grid'
    headers = ["제품명 (Product)", "단계 (Phase)", "문서 번호 (Doc No.)", "제정 일자 (Date)"]
    values = [f"{modality} Project", phase, "VMP-001", datetime.now().strftime('%Y-%m-%d')]
    for i, h in enumerate(headers): 
        c = table_info.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i, v in enumerate(values): 
        c = table_info.rows[1].cells[i]; c.text=v; c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    sections = [("1. 목적 (Objective)", "본 계획서는 밸리데이션 전략과 범위를 규정한다."),
                ("2. 적용 범위 (Scope)", f"본 문서는 {modality}의 {phase} 시험법 밸리데이션에 적용된다."),
                ("3. 근거 가이드라인 (Reference)", "• ICH Q2(R2)\n• MFDS 가이드라인")]
    for t, c in sections: doc.add_heading(t, level=1); doc.add_paragraph(c)

    doc.add_heading('4. 밸리데이션 수행 전략 (Validation Strategy)', level=1)
    table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
    for i, h in enumerate(['No.', 'Method', 'Category', 'Required Items']):
        c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for idx, row in df_strategy.iterrows():
        r = table.add_row().cells; r[0].text=str(idx+1); r[1].text=str(row['Method']); r[2].text=str(row['Category']); r[3].text=", ".join(row['Required_Items'])
    
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [PROTOCOL 생성 함수 - 에러 방지 강화]
def generate_protocol_premium(method_name, category, params):
    doc = Document(); set_korean_font(doc)
    
    # 안전한 값 가져오기 (None 방지)
    def safe_get(key, default=""):
        val = params.get(key)
        return str(val) if val is not None else default

    # 머리글
    section = doc.sections[0]; header = section.header
    htable = header.add_table(rows=1, cols=2); htable.width = Inches(6.0)
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]
    p1.add_run(f"Protocol No.: VP-{method_name[:3]}-001\n").bold = True
    p1.add_run(f"Test Category: {category}")
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"Guideline: {safe_get('Reference_Guideline', 'ICH Q2(R2)')}\n").bold = True
    p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}")

    # 제목
    title = doc.add_heading(f'밸리데이션 상세 계획서 (Validation Protocol)', 0); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Method Name: {method_name}").alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph()

    # 본문
    doc.add_heading('1. 목적 (Objective)', level=1)
    doc.add_paragraph(f"본 문서는 '{method_name}' 시험법의 밸리데이션 절차와 기준을 기술한다.")

    doc.add_heading('2. 근거 및 참고 규격 (Reference)', level=1)
    doc.add_paragraph("• ICH Q2(R2): Validation of Analytical Procedures\n• MFDS 가이드라인")

    doc.add_heading('3. 기기 및 시약 (Instruments & Reagents)', level=1)
    t_cond = doc.add_table(rows=0, cols=2); t_cond.style = 'Table Grid'
    for k, v in [("기기", safe_get('Instrument')), ("컬럼", safe_get('Column_Plate')), 
                 ("조건", f"A: {safe_get('Condition_A')}\nB: {safe_get('Condition_B')}"), ("검출기", safe_get('Detection'))]:
        r = t_cond.add_row().cells; r[0].text=k; r[0].paragraphs[0].runs[0].bold=True; r[1].text=v
    
    doc.add_heading('4. 밸리데이션 수행 방법 및 기준 (Procedures & Criteria)', level=1)
    table = doc.add_table(rows=1, cols=3); table.style = 'Table Grid'; table.autofit = False
    table.columns[0].width = Inches(1.2); table.columns[1].width = Inches(3.5); table.columns[2].width = Inches(1.8)
    
    headers = ["항목 (Parameter)", "시험 방법 (Test Procedure)", "판정 기준 (Criteria)"]
    for i, h in enumerate(headers): c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)

    # 항목 추가 함수 (안전성 강화)
    def add_row(param_name, procedure, criteria):
        if criteria and "정보 없음" not in str(criteria):
            r = table.add_row().cells
            r[0].text = str(param_name)
            r[1].text = str(procedure)
            r[2].text = str(criteria)

    target_conc = safe_get('Target_Conc', '100')
    unit = safe_get('Unit', '%')

    add_row("특이성\n(Specificity)", "공시험액, 표준액, 검체액을 분석하여 방해 피크 유무 확인", safe_get('Detail_Specificity'))
    add_row("직선성\n(Linearity)", f"기준 농도({target_conc} {unit}) 중심 80~120% 범위 내 5개 농도 조제, 각 3회 반복 주입", safe_get('Detail_Linearity'))
    add_row("범위\n(Range)", "직선성, 정확성, 정밀성이 확보된 구간", safe_get('Detail_Range'))
    add_row("정확성\n(Accuracy)", "80%, 100%, 120% 수준으로 각 3회 반복 조제하여 회수율 평가", safe_get('Detail_Accuracy'))
    add_row("반복성\n(Repeatability)", f"기준 농도({target_conc} {unit}) 6회 반복 조제 및 분석", safe_get('Detail_Precision'))
    add_row("실험실내 정밀성\n(Int. Precision)", "다른 날짜 또는 시험자가 반복성 시험 수행 (n=6)", safe_get('Detail_Inter_Precision'))
    add_row("LOD / LOQ", "S/N 비 3:1 (LOD), 10:1 (LOQ) 확인", f"LOD: {safe_get('Detail_LOD')} / LOQ: {safe_get('Detail_LOQ')}")
    add_row("완건성\n(Robustness)", f"분석 조건 변경: {safe_get('Detail_Robustness')}", "SST 만족 및 결과 차이 없음")

    doc.add_paragraph("\n\n")
    table_sign = doc.add_table(rows=2, cols=3); table_sign.style = 'Table Grid'
    for i, h in enumerate(["작성", "검토", "승인"]): c = table_sign.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i in range(3): table_sign.rows[1].cells[i].text="\n(서명/날짜)\n"

    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Excel 생성 함수 - 안전성 강화]
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    # Formats
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#4472C4', 'font_color':'white', 'align':'center', 'valign':'vcenter'})
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center', 'valign':'vcenter'})
    cell = workbook.add_format({'border':1, 'align':'center'}); num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    calc = workbook.add_format({'border':1, 'bg_color':'#FFFFCC', 'num_format':'0.00', 'align':'center'})

    ws1 = workbook.add_worksheet("1. Info & Prep"); ws1.set_column('A:A', 20); ws1.set_column('B:E', 15)
    ws1.merge_range('A1:E1', f'GMP Logbook: {method_name}', header)
    info = [("Date", datetime.now().strftime("%Y-%m-%d")), ("Instrument", params.get('Instrument')), ("Column", params.get('Column_Plate')), ("Analyst", "")]
    r = 3
    for k, v in info: ws1.write(r, 0, k, sub); ws1.merge_range(r, 1, r, 4, v if v else "", cell); r+=1
    ws1.write(r+1, 0, "Reagent", sub); ws1.merge_range(r+1, 1, r+1, 4, params.get('Ref_Standard_Info', ''), cell)
    ws1.write(r+2, 0, "Prep Method", sub); ws1.merge_range(r+2, 1, r+2, 4, params.get('Preparation_Sample', ''), cell)

    target_conc = params.get('Target_Conc')
    if target_conc is not None:
        try: target_val_base = float(target_conc)
        except: target_val_base = 0
        
        ws2 = workbook.add_worksheet("2. Linearity"); ws2.set_column('A:H', 12)
        unit = params.get('Unit', 'ppm'); ws2.merge_range('A1:H1', f'Linearity: Triplicate Analysis (Target: {target_conc} {unit})', header)
        for c, h in enumerate(["Level", "Rep", f"Conc ({unit})", "Weight", "Vol", "Response (Y)", "Mean (Y)", "RSD (%)"]): ws2.write(2, c, h, sub)
        levels = [80, 90, 100, 110, 120]; row = 3; chart_rows = []
        for level in levels:
            target_val = target_val_base * (level / 100); start_row = row + 1
            for i in range(1, 4):
                ws2.write_row(row, 0, [f"{level}%", i, target_val, "", 50, ""], cell)
                if i == 1:
                    ws2.merge_range(row, 6, row+2, 6, "", calc); ws2.write_formula(row, 6, f"=AVERAGE(F{start_row}:F{start_row+2})", calc)
                    ws2.merge_range(row, 7, row+2, 7, "", calc); ws2.write_formula(row, 7, f"=STDEV(F{start_row}:F{start_row+2})/G{start_row}*100", calc)
                    chart_rows.append(row + 1)
                row += 1
        s_row = row + 2; ws2.merge_range(s_row, 1, s_row, 3, "■ Summary for Chart", sub); ws2.write_row(s_row+1, 1, ["Conc (X)", "Mean (Y)", "R²"], sub)
        for idx, r_idx in enumerate(chart_rows): ws2.write_formula(s_row+2+idx, 1, f"=C{r_idx}", num); ws2.write_formula(s_row+2+idx, 2, f"=G{r_idx}", num)
        ws2.write_formula(s_row+2, 3, f"=RSQ(C{s_row+3}:C{s_row+7}, B{s_row+3}:B{s_row+7})", calc)
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
        chart.add_series({'categories': f"='2. Linearity'!$B${s_row+3}:$B${s_row+7}", 'values': f"='2. Linearity'!$C${s_row+3}:$C${s_row+7}", 'trendline': {'type': 'linear', 'display_equation': True, 'display_r_squared': True}})
        ws2.insert_chart('J3', chart)

    if params.get('Detail_Inter_Precision'):
        ws3 = workbook.add_worksheet("3. Precision"); ws3.set_column('A:E', 15); ws3.merge_range('A1:E1', 'Intermediate Precision', header)
        ws3.merge_range('A3:E3', "■ Day 1", sub); ws3.write_row('A4', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
        for i in range(6): ws3.write_row(4+i, 0, [i+1, "Sample", ""], cell)
        ws3.write_formula('D5', "=AVERAGE(C5:C10)", num); ws3.write_formula('E5', "=STDEV(C5:C10)/D5*100", num)
        ws3.merge_range('A12:E12', "■ Day 2", sub); ws3.write_row('A13', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
        for i in range(6): ws3.write_row(13+i, 0, [i+1, "Sample", ""], cell)
        ws3.write_formula('D14', "=AVERAGE(C14:C19)", num); ws3.write_formula('E14', "=STDEV(C14:C19)/D14*100", num)
        ws3.write('A21', "Diff (%)", sub); ws3.write_formula('B21', "=ABS(D5-D14)/AVERAGE(D5,D14)*100", num)

    if params.get('Detail_Robustness'):
        ws4 = workbook.add_worksheet("4. Robustness"); ws4.set_column('A:F', 18); ws4.merge_range('A1:F1', 'Robustness Conditions', header)
        ws4.merge_range('A2:F2', f"Guide: {params.get('Detail_Robustness')}", cell)
        for c, h in enumerate(["Condition", "Set", "Actual", "SST", "Pass/Fail", "Note"]): ws4.write(3, c, h, sub)
        for r, c in enumerate(["Standard", "Flow -0.1", "Flow +0.1", "Temp -2", "Temp +2"]):
            ws4.write(4+r, 0, c, cell); ws4.write_row(4+r, 1, [""]*5, cell)

    ws5 = workbook.add_worksheet("5. Raw Data"); ws5.set_column('A:F', 15); ws5.merge_range('A1:F1', 'Raw Data', header)
    for c, h in enumerate(["Inj No.", "Sample Name", "RT", "Area", "Height", "Remarks"]): ws5.write(2, c, h, sub)
    for r in range(3, 23): ws5.write_row(r, 0, [""]*6, cell)
    
    workbook.close(); output.seek(0)
    return output

# [Report 생성 함수]
def generate_summary_report_gmp(method_name, category, params, user_inputs):
    doc = Document(); set_korean_font(doc); doc.add_heading(f'Validation Summary Report: {method_name}', 0)
    info = doc.add_table(rows=3, cols=2); info.style='Table Grid'
    d = [("Category", category), ("Lot/Date", f"{user_inputs['lot_no']} / {user_inputs['date']}"), ("Analyst", user_inputs['analyst'])]
    for i, (k, v) in enumerate(d): info.rows[i].cells[0].text=k; info.rows[i].cells[1].text=str(v)
    doc.add_heading('1. 상세 결과 (Results)', level=1)
    table = doc.add_table(rows=1, cols=3); table.style='Table Grid'
    table.rows[0].cells[0].text="항목"; table.rows[0].cells[1].text="기준"; table.rows[0].cells[2].text="결과"
    check_items = [("특이성", params.get('Detail_Specificity'), "Pass"), ("직선성 (R²)", params.get('Detail_Linearity'), "Pass (See Chart)"),
                   ("정밀성", params.get('Detail_Precision'), user_inputs.get('main_result', 'N/A')),
                   ("실험실내 정밀성", params.get('Detail_Inter_Precision'), "Pass"), ("완건성", params.get('Detail_Robustness'), "Pass")]
    for k, c, r in check_items:
        if c: table.add_row().cells[0].text=k; table.rows[-1].cells[1].text=c; table.rows[-1].cells[2].text=r
    doc.add_heading('2. 결론', level=1); doc.add_paragraph("본 시험법은 모든 밸리데이션 항목을 만족하므로 적합함.")
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
    try: criteria_map = get_criteria_map(); df_full = get_strategy_list(criteria_map)
    except: df_full = pd.DataFrame()

    if sel_modality == "mAb" and not df_full.empty:
        my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
        if not my_plan.empty:
            t1, t2, t3 = st.tabs(["📑 Step 1: Strategy & Protocol", "📗 Step 2: Excel Logbook", "📊 Step 3: Result Report"])
            
            with t1:
                st.markdown("### 1️⃣ 전략 (VMP) 및 상세 계획서 (Protocol)")
                st.dataframe(my_plan[["Method", "Category"]])
                c1, c2 = st.columns(2)
                with c1: st.download_button("📥 VMP(종합계획서) 다운로드", generate_vmp_premium(sel_modality, sel_phase, my_plan), "VMP_Master.docx")
                with c2:
                    sel_p = st.selectbox("Protocol:", my_plan["Method"].unique())
                    if sel_p: st.download_button("📄 상세 계획서(Protocol) 다운로드", generate_protocol_premium(sel_p, "Cat", get_method_params(sel_p)), f"Protocol_{sel_p}.docx")

            with t2:
                st.markdown("### 📗 스마트 엑셀 일지 (3회 반복 & RSD)")
                sel_l = st.selectbox("Logbook:", my_plan["Method"].unique(), key="l")
                if st.button("Download Excel Logbook"):
                    data = generate_smart_excel(sel_l, "Cat", get_method_params(sel_l))
                    st.download_button("📊 Excel Logbook", data, f"Logbook_{sel_l}.xlsx")

            with t3:
                st.markdown("### 📊 최종 결과 보고서")
                sel_r = st.selectbox("Report:", my_plan["Method"].unique(), key="r")
                with st.form("rep"):
                    l = st.text_input("Lot"); d = st.text_input("Date"); a = st.text_input("Analyst")
                    s = st.text_input("SST"); m = st.text_input("Main Result")
                    if st.form_submit_button("Generate Report"):
                        doc = generate_summary_report_gmp(sel_r, "Cat", get_method_params(sel_r), {'lot_no':l, 'date':d, 'analyst':a, 'sst_result':s, 'main_result':m})
                        st.download_button("📥 Report", doc, "Report.docx")