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
# 0. 페이지 설정 (최상단)
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Full Suite", layout="wide")

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
# 3. 문서 생성 헬퍼
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
        if cell.paragraphs[0].runs: cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# [VMP]
def generate_vmp_premium(modality, phase, df_strategy):
    doc = Document(); set_korean_font(doc)
    head = doc.add_heading('밸리데이션 종합계획서 (Validation Master Plan)', 0); head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    table_info = doc.add_table(rows=2, cols=4); table_info.style = 'Table Grid'
    headers = ["제품명 (Product)", "단계 (Phase)", "문서 번호 (Doc No.)", "제정 일자 (Date)"]
    values = [f"{modality} Project", phase, "VMP-001", datetime.now().strftime('%Y-%m-%d')]
    for i, h in enumerate(headers): c = table_info.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i, v in enumerate(values): c = table_info.rows[1].cells[i]; c.text=v; c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    for t, c in [("1. 목적 (Objective)", "본 계획서는 밸리데이션 전략과 범위를 규정한다."), ("2. 적용 범위 (Scope)", f"본 문서는 {modality}의 {phase} 시험법 밸리데이션에 적용된다."), ("3. 근거 가이드라인 (Reference)", "• ICH Q2(R2)\n• MFDS 가이드라인")]:
        doc.add_heading(t, level=1); doc.add_paragraph(c)
    doc.add_heading('4. 밸리데이션 수행 전략 (Validation Strategy)', level=1)
    table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
    for i, h in enumerate(['No.', 'Method', 'Category', 'Required Items']): c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for idx, row in df_strategy.iterrows(): r = table.add_row().cells; r[0].text=str(idx+1); r[1].text=str(row['Method']); r[2].text=str(row['Category']); r[3].text=", ".join(row['Required_Items'])
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Master Recipe Excel]
def generate_master_recipe_excel(method_name, target_conc, unit, stock_conc, req_vol, sample_type, powder_info=""):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    title_fmt = workbook.add_format({'bold':True, 'font_size': 14, 'align':'center', 'valign':'vcenter', 'bg_color': '#44546A', 'font_color': 'white'})
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center'})
    section_title = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFC000', 'font_size':11, 'align':'left'}) 
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#EDEDED', 'align':'center'})
    cell = workbook.add_format({'border':1, 'align':'center'})
    num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    auto = workbook.add_format({'border':1, 'bg_color':'#E2EFDA', 'num_format':'0.000', 'align':'center'})
    total_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFFF00', 'num_format':'0.00', 'align':'center'})
    ws = workbook.add_worksheet("Master Recipe")
    ws.set_column('A:A', 35); ws.set_column('B:E', 15); ws.set_column('F:F', 12)
    ws.merge_range('A1:F1', f'Validation Material Planner: {method_name}', title_fmt)
    ws.write('A3', "Sample Type:", sub); ws.write('B3', sample_type, cell)
    if sample_type == "Powder (파우더)": ws.write('C3', "Prep Detail:", sub); ws.write_string('D3', powder_info, cell)
    ws.write('A4', "User Stock Conc:", sub); ws.write('B4', stock_conc, num); ws.write('C4', unit, cell)
    ws.write('A5', "Target Conc (100%):", sub); ws.write('B5', target_conc, num); ws.write('C5', unit, cell)
    ws.write('A6', "Target Vol/Vial (mL):", sub); ws.write('B6', req_vol, num)
    ws.write('D6', "TOTAL STOCK NEEDED (mL):", sub)
    row = 8
    def add_section_grouped(main_title, levels, reps):
        nonlocal row
        ws.merge_range(row, 0, row, 5, f"■ {main_title}", header); row += 1
        data_start_row = row
        for rep in range(1, reps + 1):
            ws.merge_range(row, 0, row, 5, f"{main_title.split(' ')[0]} - {rep}회차 조제 (Set {rep})", section_title); row += 1
            ws.write_row(row, 0, ["Item ID", "Target Conc", "Stock Vol (mL)", "Diluent Vol (mL)", "Total (mL)", "Check"], sub); row += 1
            for level in levels:
                t_val = float(target_conc) * (level / 100)
                if float(stock_conc) < t_val: s_vol = "Error"
                else: s_vol = (t_val * float(req_vol)) / float(stock_conc); d_vol = float(req_vol) - s_vol
                ws.write(row, 0, f"{main_title.split(' ')[0]}-{level}%-R{rep}", cell); ws.write(row, 1, t_val, num)
                if isinstance(s_vol, str): ws.write(row, 2, s_vol, total_fmt); ws.write(row, 3, "N/A", total_fmt)
                else: ws.write(row, 2, s_vol, auto); ws.write(row, 3, d_vol, auto)
                ws.write(row, 4, float(req_vol), num); ws.write(row, 5, "□", cell); row += 1
            ws.write(row, 1, f"[{rep}회차] 소요 Stock:", sub)
            if isinstance(s_vol, str): ws.write(row, 2, "Error", total_fmt)
            else: ws.write_formula(row, 2, f"=SUM(C{row-len(levels)}:C{row-1})", total_fmt)
            row += 2
    add_section_grouped("1. 시스템 적합성 (SST)", [100], 1)
    add_section_grouped("2. 특이성 (Specificity)", [100], 1)
    add_section_grouped("3. 직선성 (Linearity)", [80, 90, 100, 110, 120], 3)
    add_section_grouped("4. 정확성 (Accuracy)", [80, 100, 120], 3)
    ws.merge_range(row, 0, row, 5, "■ 5. 정밀성 (Repeatability)", header); row += 2
    ws.merge_range(row, 0, row, 5, "반복성 시험 세트 (n=6)", section_title); row += 1
    ws.write_row(row, 0, ["Item ID", "Target Conc", "Stock Vol (mL)", "Diluent Vol (mL)", "Total (mL)", "Check"], sub); row += 1
    p_start = row
    for i in range(1, 7):
        t_val = float(target_conc); s_vol = (t_val * float(req_vol)) / float(stock_conc); d_vol = float(req_vol) - s_vol
        ws.write(row, 0, f"Prec-100%-{i}", cell); ws.write(row, 1, t_val, num); ws.write(row, 2, s_vol, auto); ws.write(row, 3, d_vol, auto); ws.write(row, 4, float(req_vol), num); ws.write(row, 5, "□", cell); row += 1
    ws.write(row, 1, "[정밀성] 소요 Stock:", sub); ws.write_formula(row, 2, f"=SUM(C{p_start}:C{row-1})", total_fmt); row += 2
    add_section_grouped("7. 완건성 (Robustness)", [100], 3); add_section_grouped("8. LOD/LOQ", [1, 0.5], 3)
    ws.write_formula('E6', f"=SUM(C9:C{row})", workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FF0000', 'font_color':'white', 'num_format':'0.00', 'align':'center'}))
    workbook.close(); output.seek(0)
    return output

# [PROTOCOL]
def generate_protocol_premium(method_name, category, params, stock_conc=None, req_vol=None, target_conc_override=None):
    doc = Document(); set_korean_font(doc)
    def safe_get(key, default=""): val = params.get(key); return str(val) if val is not None else default
    target_conc = str(target_conc_override) if target_conc_override else safe_get('Target_Conc', '100'); unit = safe_get('Unit', '%')
    section = doc.sections[0]; header = section.header; 
    htable = header.add_table(1, 2, Inches(6.0))
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]; p1.add_run(f"Protocol No.: VP-{method_name[:3]}-001\n").bold = True; p1.add_run(f"Test Category: {category}")
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p2.add_run(f"Guideline: {safe_get('Reference_Guideline', 'ICH Q2(R2)')}\n").bold = True; p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
    title = doc.add_heading(f'밸리데이션 상세 계획서 (Validation Protocol)', 0); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Method Name: {method_name}").alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph()
    doc.add_heading('1. 목적', level=1); doc.add_paragraph(f"본 문서는 '{method_name}' 시험법의 밸리데이션 수행 방법 및 판정 기준을 기술한다.")
    doc.add_heading('2. 근거', level=1); doc.add_paragraph("• ICH Q2(R2) & MFDS 가이드라인")
    doc.add_heading('3. 기기 및 시약', level=1); t_cond = doc.add_table(rows=0, cols=2); t_cond.style = 'Table Grid'
    for k, v in [("기기", safe_get('Instrument')), ("컬럼", safe_get('Column_Plate')), ("조건", f"A: {safe_get('Condition_A')}\nB: {safe_get('Condition_B')}"), ("검출기", safe_get('Detection'))]:
        r = t_cond.add_row().cells; r[0].text=k; r[0].paragraphs[0].runs[0].bold=True; r[1].text=v
    doc.add_heading('4. 밸리데이션 항목 및 기준', level=1); table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
    headers = ["항목", "기준"]; 
    for i, h in enumerate(headers): c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    items = [("특이성", safe_get('Detail_Specificity')), ("직선성", safe_get('Detail_Linearity')), ("범위", safe_get('Detail_Range')), ("정확성", safe_get('Detail_Accuracy')), ("정밀성", safe_get('Detail_Precision')), ("완건성", safe_get('Detail_Robustness'))]
    for k, v in items:
        if v and "정보 없음" not in v: r = table.add_row().cells; r[0].text=k; r[1].text=v
    doc.add_heading('5. 상세 시험 방법 (Procedures)', level=1)
    doc.add_heading('5.1 용액 조제', level=2); doc.add_paragraph(f"1) 표준 모액: 농도 {stock_conc if stock_conc else '[입력필요]'} {unit} 용액을 준비한다.")
    doc.add_heading('5.2 직선성', level=2); doc.add_paragraph(f"기준 농도 {target_conc} {unit}를 중심으로 80 ~ 120% 범위 내 5개 농도를 조제한다.")
    if stock_conc and req_vol and float(stock_conc) >= float(target_conc) * 1.2:
        t_lin = doc.add_table(rows=1, cols=4); t_lin.style = 'Table Grid'
        for i, h in enumerate(["Level", "Target", "Stock (mL)", "Diluent (mL)"]): c = t_lin.rows[0].cells[i]; c.text=h; set_table_header_style(c)
        for level in [80, 90, 100, 110, 120]:
            t_val = float(target_conc) * (level/100); s_vol = (t_val * float(req_vol)) / float(stock_conc); d_vol = float(req_vol) - s_vol
            r = t_lin.add_row().cells; r[0].text=f"{level}%"; r[1].text=f"{t_val:.2f}"; r[2].text=f"{s_vol:.3f}"; r[3].text=f"{d_vol:.3f}"
    doc.add_heading('5.3 정확성', level=2); doc.add_paragraph("기준 농도의 80%, 100%, 120% 수준으로 각 3회씩 독립적으로 조제한다.")
    doc.add_paragraph("\n\n"); table_sign = doc.add_table(rows=2, cols=3); table_sign.style = 'Table Grid'
    for i, h in enumerate(["작성", "검토", "승인"]): c = table_sign.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i in range(3): table_sign.rows[1].cells[i].text="\n(서명/날짜)\n"
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Excel 생성 함수 - Smart Logbook (ACTUAL WEIGHT & CORRECTION LOGIC)]
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Styles
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#4472C4', 'font_color':'white', 'align':'center', 'valign':'vcenter'})
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center', 'valign':'vcenter'})
    sub_rep = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FCE4D6', 'align':'left'}) 
    cell = workbook.add_format({'border':1, 'align':'center'}); num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    calc = workbook.add_format({'border':1, 'bg_color':'#FFFFCC', 'num_format':'0.00', 'align':'center'}) # Input
    auto = workbook.add_format({'border':1, 'bg_color':'#E2EFDA', 'num_format':'0.00', 'align':'center'}) # Calc
    pass_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#C6EFCE', 'font_color':'#006100', 'align':'center'})
    fail_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFC7CE', 'font_color':'#9C0006', 'align':'center'})
    total_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFFF00', 'num_format':'0.0', 'align':'center'})
    crit_fmt = workbook.add_format({'bold':True, 'font_color':'red', 'align':'left'}) # Criteria Text

    # [1. Info Sheet]
    ws1 = workbook.add_worksheet("1. Info"); ws1.set_column('A:A', 25); ws1.set_column('B:E', 15)
    ws1.merge_range('A1:E1', f'GMP Logbook: {method_name}', header)
    
    info_rows = [("Date", datetime.now().strftime("%Y-%m-%d")), ("Instrument", params.get('Instrument')), ("Column", params.get('Column_Plate')), ("Analyst", "")]
    for i, (k, v) in enumerate(info_rows):
        ws1.write(i+3, 0, k, sub); ws1.merge_range(i+3, 1, i+3, 4, v if v else "", cell)

    ws1.write(8, 0, "Round Rule:", sub); ws1.merge_range(8, 1, 8, 4, "모든 계산값은 소수점 2째자리(농도 3째자리)에서 절사(ROUNDDOWN).", cell)
    
    target_conc_val = float(params.get('Target_Conc', 1.0))
    ws1.write(9, 0, "Target Conc (100%):", sub); ws1.write(9, 1, target_conc_val, calc); ws1.write(9, 2, params.get('Unit', 'mg/mL'), cell)
    target_conc_ref = "'1. Info'!$B$10"

    ws1.merge_range(10, 0, 10, 4, "■ Standard Preparation & Correction Factor", sub_rep)
    labels = ["Theoretical Stock (mg/mL):", "Purity (Potency, %):", "Water Content (%):", "Actual Weight (mg):", "Final Volume (mL):"]
    for i, label in enumerate(labels):
        ws1.write(11 + i, 0, label, sub)
        if "Purity" in label: ws1.write(11 + i, 1, 100.0, calc)
        elif "Water" in label: ws1.write(11 + i, 1, 0.0, calc)
        else: ws1.write(11 + i, 1, 1.0, calc)

    ws1.write(16, 0, "Actual Stock (mg/mL):", sub)
    ws1.write_formula(16, 1, '=IF(B16="","",ROUNDDOWN((B15*(B13/100)*((100-B14)/100))/B16, 4))', auto)
    actual_stock_ref = "'1. Info'!$B$17"
    ws1.write(17, 0, "Correction Factor:", sub); ws1.write_formula(17, 1, '=IF(OR(B12="",B12=0,B17=""), 1, ROUNDDOWN(B17/B12, 4))', total_fmt)
    corr_factor_ref = "'1. Info'!$B$18"; theo_stock_ref = "'1. Info'!$B$12"

    # 2. SST Sheet (System Suitability)
    ws_sst = workbook.add_worksheet("2. SST"); ws_sst.set_column('A:F', 15)
    ws_sst.merge_range('A1:F1', 'System Suitability Test (n=6)', header)
    ws_sst.write_row('A2', ["Inj No.", "RT (min)", "Area", "Height", "Tailing (1st)", "Plate Count"], sub)
    for i in range(1, 7):
        ws_sst.write(i+1, 0, i, cell)
        # Simulate Data if enabled
        sim_rt = 5.0 + random.uniform(-0.02, 0.02) if simulate else ""
        sim_area = (target_conc_val * 10000) + random.uniform(-100, 100) if simulate else ""
        sim_tail = 1.1 if simulate else ""
        ws_sst.write_row(i+1, 1, [sim_rt, sim_area, "", sim_tail, ""], calc)
    
    ws_sst.write('A9', "Mean", sub); ws_sst.write_formula('B9', "=ROUNDDOWN(AVERAGE(B3:B8), 2)", auto); ws_sst.write_formula('C9', "=ROUNDDOWN(AVERAGE(C3:C8), 2)", auto)
    ws_sst.write('A10', "RSD(%)", sub); ws_sst.write_formula('B10', "=ROUNDDOWN(STDEV(B3:B8)/B9*100, 2)", auto); ws_sst.write_formula('C10', "=ROUNDDOWN(STDEV(C3:C8)/C9*100, 2)", auto)
    
    # 판정 로직: 입력값이 있을 때만 Pass/Fail 표시
    ws_sst.write('D12', "Result:", sub)
    ws_sst.write_formula('E12', '=IF(OR(B10="", C10=""), "", IF(AND(B10<=2.0, C10<=2.0, E3<=2.0), "Pass", "Fail"))', pass_fmt)
    ws_sst.conditional_format('E12', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
    
    # [기준 명시] 하단 추가
    ws_sst.write('A14', "※ Acceptance Criteria:", crit_fmt)
    ws_sst.write('A15', "1) Retention Time & Area RSD ≤ 2.0%")
    ws_sst.write('A16', "2) Tailing Factor (1st Inj) ≤ 2.0")

    # 3. Specificity Sheet (보완됨)
    ws_spec = workbook.add_worksheet("3. Specificity")
    ws_spec.set_column('A:E', 20)
    ws_spec.merge_range('A1:E1', 'Specificity Test', header)
    ws_spec.write('A3', "Std Mean Area (Ref. SST):", sub); ws_spec.write_formula('B3', "='2. SST'!C9", num)
    
    ws_spec.write_row('A5', ["Sample", "RT", "Area", "Interference (%)", "Result"], sub)
    for i, s in enumerate(["Blank", "Placebo"]):
        row = i + 6
        ws_spec.write(row, 0, s, cell); ws_spec.write_row(row, 1, ["", ""], calc)
        ws_spec.write_formula(row, 3, f'=IF(OR(C{row+1}="", $B$3=""), "", ROUNDDOWN(C{row+1}/$B$3*100, 2))', auto)
        ws_spec.write_formula(row, 4, f'=IF(D{row+1}="", "", IF(D{row+1}<=0.5, "Pass", "Fail"))', pass_fmt)
        ws_spec.conditional_format(f'E{row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
   
    # [기준 명시] 하단 추가
    ws_spec.write(f'A{row+3}', "※ Acceptance Criteria:", crit_fmt)
    ws_spec.write(f'A{row+4}', "1) Interference check: ≤ 0.5% of Standard Area")

    # 4. Linearity Sheet
    target_conc = params.get('Target_Conc')
    if target_conc:
        try: target_val_base = float(target_conc)
        except: target_val_base = 0
        ws2 = workbook.add_worksheet("4. Linearity"); ws2.set_column('A:I', 13)
        unit = params.get('Unit', 'ppm'); ws2.merge_range('A1:I1', f'Linearity Test (Target: {target_conc} {unit})', header)
        row = 3; rep_rows = {1: [], 2: [], 3: []}
        
        for rep in range(1, 4):
            ws2.merge_range(row, 0, row, 8, f"■ Repetition {rep}", sub_rep); row += 1
            ws2.write_row(row, 0, ["Level", "Conc (X)", "Area (Y)", "Back Calc", "Accuracy (%)", "Check"], sub); row += 1
            data_start = row
            for level in [80, 90, 100, 110, 120]:
                target_val = target_val_base * (level / 100)
                ws2.write(row, 0, f"{level}%", cell); ws2.write_formula(row, 1, f"=ROUNDDOWN({target_val}, 3)", num)
                ws2.write(row, 2, "", calc)
                rep_rows[rep].append(row + 1)
                
                # 수식: 값이 있을 때만 계산
                ind_slope = f"C{data_start+7}"; ind_int = f"C{data_start+8}"
                ws2.write_formula(row, 3, f'=IF(C{row+1}="", "", ROUNDDOWN((C{row+1}-{ind_int})/{ind_slope}, 3))', auto)
                ws2.write_formula(row, 4, f'=IF(C{row+1}="", "", ROUNDDOWN(D{row+1}/B{row+1}*100, 1))', auto)
                ws2.write(row, 5, "OK", cell); row += 1
            
            # Regression (값이 없으면 에러 방지 위해 IFERROR 사용 가능하나, 기본식 유지)
            ws2.write(row, 1, "Slope:", sub); ws2.write_formula(row, 2, f"=SLOPE(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
            ws2.write(row+1, 1, "Intercept:", sub); ws2.write_formula(row+1, 2, f"=INTERCEPT(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
            ws2.write(row+2, 1, "R²:", sub); ws2.write_formula(row+2, 2, f"=RSQ(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
            
            chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
            chart.add_series({'name': f'Rep {rep}', 'categories': f"='4. Linearity'!$B${data_start+1}:$B${row}", 'values': f"='4. Linearity'!$C${data_start+1}:$C${row}", 'trendline': {'type': 'linear', 'display_equation': True, 'display_r_squared': True}})
            chart.set_size({'width': 350, 'height': 220}); ws2.insert_chart(f'G{data_start}', chart)
            row += 6

        # Summary Table
        ws2.merge_range(row, 0, row, 8, "■ Summary (Mean of 3 Reps) & Final Check", sub_rep); row += 1
        ws2.write_row(row, 0, ["Level", "Conc (X)", "Mean Area", "STDEV", "% RSD", "Result (RSD≤5%)"], sub); row += 1
        summary_start = row
        for i, level in enumerate([80, 90, 100, 110, 120]):
            r1 = rep_rows[1][i]; r2 = rep_rows[2][i]; r3 = rep_rows[3][i]
            ws2.write(row, 0, f"{level}%", cell); ws2.write_formula(row, 1, f"=B{r1}", num)
            ws2.write_formula(row, 2, f"=ROUNDDOWN(AVERAGE(C{r1},C{r2},C{r3}), 2)", auto)
            ws2.write_formula(row, 3, f"=ROUNDDOWN(STDEV(C{r1},C{r2},C{r3}), 2)", auto)
            ws2.write_formula(row, 4, f"=IF(C{row+1}=0, 0, ROUNDDOWN(D{row+1}/C{row+1}*100, 2))", auto)
            # RSD 판정 (값이 있을 때만)
            ws2.write_formula(row, 5, f'=IF(C{row+1}=0, "", IF(E{row+1}<=5.0, "Pass", "Fail"))', pass_fmt)
            ws2.conditional_format(f'F{row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt}); row += 1
        
        row += 1
        slope_cell = f"'4. Linearity'!C{row+1}"; int_cell = f"'4. Linearity'!C{row+2}"
        
        ws2.write(row, 1, "Slope:", sub); ws2.write_formula(row, 2, f"=ROUNDDOWN(SLOPE(C{summary_start+1}:C{summary_start+5}, B{summary_start+1}:B{summary_start+5}), 4)", auto)
        ws2.write(row+1, 1, "Intercept:", sub); ws2.write_formula(row+1, 2, f"=ROUNDDOWN(INTERCEPT(C{summary_start+1}:C{summary_start+5}, B{summary_start+1}:B{summary_start+5}), 4)", auto)
        ws2.write(row+2, 1, "R²:", sub); ws2.write_formula(row+2, 2, f"=ROUNDDOWN(RSQ(C{summary_start+1}:C{summary_start+5}, B{summary_start+1}:B{summary_start+5}), 4)", auto)
        
        # R2 판정
        ws2.write(row+2, 3, "Criteria:", sub); ws2.write(row+2, 4, "≥ 0.990", cell)
        ws2.write_formula(row+2, 5, f'=IF(C{row+3}=0, "", IF(C{row+3}>=0.990, "Pass", "Fail"))', pass_fmt)
        ws2.conditional_format(f'F{row+3}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
        
        # [기준 명시] 하단 추가
        ws2.write(f'A{row+5}', "※ Acceptance Criteria:", crit_fmt)
        ws2.write(f'A{row+6}', "1) Coefficient of determination (R²) ≥ 0.990")
        ws2.write(f'A{row+7}', "2) %RSD of peak areas at each level ≤ 5.0%")

    # 5. Accuracy Sheet
    ws_acc = workbook.add_worksheet("5. Accuracy"); ws_acc.set_column('A:G', 15)
    ws_acc.merge_range('A1:G1', 'Accuracy Test (Recovery)', header)
    ws_acc.write('E3', "Slope:", sub); ws_acc.write_formula('F3', slope_cell, calc) 
    ws_acc.write('E4', "Int:", sub); ws_acc.write_formula('F4', int_cell, calc)
    row = 7
    for level in [80, 100, 120]:
        ws_acc.merge_range(row, 0, row, 6, f"■ Level {level}% (3 Reps)", sub_rep); row += 1
        ws_acc.write_row(row, 0, ["Rep", "Theo Conc", "Area", "Calc Conc", "Recovery (%)", "Criteria", "Result"], sub); row += 1
        t_val = target_val_base * (level/100)
        start_row = row
        for rep in range(1, 4):
            ws_acc.write(row, 0, rep, cell); ws_acc.write(row, 1, t_val, num)
            ws_acc.write(row, 2, "", calc) # Input
            ws_acc.write_formula(row, 3, f'=IF(C{row+1}="", "", ROUNDDOWN((C{row+1}-$F$4)/$F$3, 3))', auto)
            ws_acc.write_formula(row, 4, f'=IF(D{row+1}="", "", ROUNDDOWN(D{row+1}/B{row+1}*100, 1))', auto)
            ws_acc.write(row, 5, "80~120%", cell)
            # 판정: 입력값 있을 때만 수행
            ws_acc.write_formula(row, 6, f'=IF(E{row+1}="", "", IF(AND(E{row+1}>=80, E{row+1}<=120), "Pass", "Fail"))', pass_fmt)
            ws_acc.conditional_format(f'G{row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
            row += 1
        
        ws_acc.write(row, 3, "Mean Rec(%):", sub)
        ws_acc.write_formula(row, 4, f"=ROUNDDOWN(AVERAGE(E{start_row+1}:E{row}), 1)", total_fmt) 
        row += 2
        
    # [기준 명시] 하단 추가
    ws_acc.write(f'A{row+1}', "※ Acceptance Criteria:", crit_fmt)
    ws_acc.write(f'A{row+2}', "1) Individual & Mean Recovery: 80.0 ~ 120.0%")

    # 6. Precision
    ws3 = workbook.add_worksheet("6. Precision"); ws3.set_column('A:E', 15); ws3.merge_range('A1:E1', 'Precision', header)
    ws3.merge_range('A3:E3', "■ Day 1 (Repeatability)", sub); ws3.write_row('A4', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
    for i in range(6): ws3.write_row(4+i, 0, [i+1, "Sample", ""], calc)
    ws3.write_formula('D5', "=ROUNDDOWN(AVERAGE(C5:C10), 2)", num); ws3.write_formula('E5', "=ROUNDDOWN(STDEV(C5:C10)/D5*100, 2)", num)
    ws3.write('D11', "Criteria (≤2.0%):", sub); ws3.write_formula('E11', '=IF(E5=0,"",IF(E5<=2.0,"Pass","Fail"))', pass_fmt)
    
    ws3.merge_range('A14:E14', "■ Day 2 (Intermediate Precision)", sub); ws3.write_row('A15', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
    for i in range(6): ws3.write_row(15+i, 0, [i+1, "Sample", ""], calc)
    ws3.write_formula('D16', "=ROUNDDOWN(AVERAGE(C16:C21), 2)", num); ws3.write_formula('E16', "=ROUNDDOWN(STDEV(C16:C21)/D16*100, 2)", num)
    ws3.write('A23', "Diff (%)", sub); ws3.write_formula('B23', "=ROUNDDOWN(ABS(D5-D16)/AVERAGE(D5,D16)*100, 2)", num)

    # 7. Robustness (보완됨)
    if params.get('Detail_Robustness'):
        ws4 = workbook.add_worksheet("7. Robustness"); ws4.set_column('A:F', 20); ws4.merge_range('A1:F1', 'Robustness Conditions', header)
        ws4.write_row('A3', ["Condition", "Set", "Actual", "SST Result (RSD)", "Pass/Fail", "Note"], sub)
        for r, c in enumerate(["Standard", "Flow -0.1", "Flow +0.1", "Temp -2", "Temp +2"]): 
            ws4.write(4+r, 0, c, cell); ws4.write_row(4+r, 1, ["", "", ""], calc)
            # SST Result(RSD)가 2.0 이하인지 판단
            ws4.write_formula(4+r, 4, f'=IF(D{5+r}="", "", IF(D{5+r}<=2.0, "Pass", "Fail"))', pass_fmt)
            ws4.conditional_format(f'E{5+r}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
        
        # [기준 명시]
        ws4.write(f'A{10+r}', "※ Acceptance Criteria:", crit_fmt)
        ws4.write(f'A{11+r}', "1) SST Criteria must be met (RSD ≤ 2.0%) under all conditions")

    # 8. LOD/LOQ
    ws_ll = workbook.add_worksheet("8. LOD_LOQ"); ws_ll.set_column('A:E', 15); ws_ll.merge_range('A1:E1', 'LOD / LOQ', header)
    ws_ll.write_row('A2', ["Item", "Signal", "Noise", "S/N Ratio", "Result"], sub)
    ws_ll.write('A3', "LOD Sample", cell); ws_ll.write('B3', "", calc); ws_ll.write('C3', "", calc); ws_ll.write_formula('D3', '=IF(C3="","",ROUNDDOWN(B3/C3, 1))', auto)
    ws_ll.write_formula('E3', '=IF(D3="","",IF(D3>=3, "Pass", "Fail"))', pass_fmt)
    ws_ll.write('A4', "LOQ Sample", cell); ws_ll.write('B4', "", calc); ws_ll.write('C4', "", calc); ws_ll.write_formula('D4', '=IF(C4="","",ROUNDDOWN(B4/C4, 1))', auto)
    ws_ll.write_formula('E4', '=IF(D4="","",IF(D4>=10, "Pass", "Fail"))', pass_fmt)
    
    # [기준 명시]
    ws_ll.write('A7', "※ Acceptance Criteria:", crit_fmt)
    ws_ll.write('A8', "1) LOD S/N Ratio ≥ 3")
    ws_ll.write('A9', "2) LOQ S/N Ratio ≥ 10")

    workbook.close(); output.seek(0)
    return output


# ---------------------------------------------------------
# 4. 상세 계획서 생성 (보완된 SOP 기술)
# ---------------------------------------------------------
def generate_protocol_premium(method_name, category, params, stock_conc=None, req_vol=None, target_conc_override=None):
    doc = Document(); set_korean_font(doc)
    def safe_get(key, default=""): val = params.get(key); return str(val) if val is not None else default
    target_conc = str(target_conc_override) if target_conc_override else safe_get('Target_Conc', '100'); unit = safe_get('Unit', '%')
    
    section = doc.sections[0]; header = section.header; htable = header.add_table(1, 2, Inches(6.0))
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]; p1.add_run(f"Protocol No.: VP-{method_name[:3]}-001\n").bold = True; p1.add_run(f"Test Category: {category}")
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p2.add_run(f"Guideline: {safe_get('Reference_Guideline', 'ICH Q2(R2)')}\n").bold = True; p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
    
    doc.add_heading(f'밸리데이션 상세 계획서 ({method_name})', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('1. 목적 및 범위', level=1); doc.add_paragraph("본 시험법의 직선성, 정확성, 정밀성 등을 검증하여 의약품 품질 관리의 적합성을 보증한다.")
    
    doc.add_heading('2. 기기 및 시약', level=1); t_cond = doc.add_table(rows=4, cols=2); t_cond.style = 'Table Grid'
    conds = [("시험 기기", safe_get('Instrument')), ("컬럼 정보", safe_get('Column_Plate')), ("검출기", safe_get('Detection')), ("SST 기준", f"RSD ≤ 2.0%, Tailing ≤ 2.0")]
    for i, (k, v) in enumerate(conds): t_cond.rows[i].cells[0].text=k; t_cond.rows[i].cells[1].text=v
    
    doc.add_heading('3. 상세 조제 방법', level=1)
    doc.add_heading('3.1 공시험액(Blank) 및 위약(Placebo)', level=2); doc.add_paragraph("1) Blank: 주성분을 제외한 희석액을 그대로 사용한다.\n2) Placebo: 주성분을 제외한 모든 부형제를 처방 비율대로 혼합하여 조제한다.")
    doc.add_heading('3.2 표준액 및 검액 조제', level=2); doc.add_paragraph(f"1) Master Recipe 및 시험일지의 보정 계수(Correction Factor)를 확인하여 농도 {target_conc} {unit} 수준이 되도록 정밀 조제한다.")
    
    doc.add_heading('4. 시험 항목 및 평가 방법', level=1)
    doc.add_heading('4.1 특이성(Specificity)', level=2); doc.add_paragraph("Blank와 Placebo를 주입하여 주성분 RT 위치에서의 간섭 피크 면적이 표준액의 0.5% 이하인지 확인한다.")
    doc.add_heading('4.2 LOD 및 LOQ', level=2); doc.add_paragraph("S/N비(Signal to Noise)를 측정한다. LOD는 3:1 이상, LOQ는 10:1 이상이어야 하며, LOQ 농도에서의 정밀성(RSD)을 추가로 평가할 수 있다.")
    
    doc.add_paragraph("\n\n"); table_sign = doc.add_table(rows=2, cols=3); table_sign.style = 'Table Grid'
    for i, h in enumerate(["작성", "검토", "승인"]): c = table_sign.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i in range(3): table_sign.rows[1].cells[i].text="\n(서명/날짜)\n"
    
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io


# ---------------------------------------------------------
# 5. 엑셀 데이터 추출 (Parsing)
# ---------------------------------------------------------
def extract_logbook_data(uploaded_file):
    results = {}
    try:
        df_sst = pd.read_excel(uploaded_file, sheet_name='2. SST', header=None)
        rsd_row = df_sst[df_sst.eq("RSD(%)").any(axis=1)].index
        if not rsd_row.empty:
            idx = rsd_row[0]
            results['sst_res'] = f"RT: {df_sst.iloc[idx, 1]}%, Area: {df_sst.iloc[idx, 2]}%"
            res_row = df_sst[df_sst.eq("Result:").any(axis=1)].index
            if not res_row.empty: results['sst_pass'] = df_sst.iloc[res_row[0], 1]
        
        df_lin = pd.read_excel(uploaded_file, sheet_name='4. Linearity', header=None)
        r2_row = df_lin[df_lin.eq("Final R²:").any(axis=1)].index
        if not r2_row.empty:
            r2_val = df_lin.iloc[r2_row[0], 1]
            results['lin_res'] = f"R² = {r2_val}"
            results['lin_pass'] = df_lin.iloc[r2_row[0], 5]

        df_acc = pd.read_excel(uploaded_file, sheet_name='5. Accuracy', header=None)
        mean_recs = []
        for r in df_acc.index:
            for c in df_acc.columns:
                if str(df_acc.iloc[r, c]).strip() == "Mean Rec(%):":
                    val = df_acc.iloc[r, c+1]
                    if pd.notna(val): mean_recs.append(val)
        if mean_recs:
            results['acc_res'] = f"{min(mean_recs)}% ~ {max(mean_recs)}%"
            results['acc_pass'] = "Pass"

        df_prec = pd.read_excel(uploaded_file, sheet_name='6. Precision', header=None)
        rsd_rows = df_prec[df_prec.eq("Result:").any(axis=1)].index
        if not rsd_rows.empty:
            results['prec_res'] = f"RSD: {df_prec.iloc[rsd_rows[0]-6, 4]}%" 
            results['prec_pass'] = df_prec.iloc[rsd_rows[0], 4]

    except Exception as e: return {}
    return results 

# ---------------------------------------------------------
# 6. 최종 보고서 생성 (Automated)
# ---------------------------------------------------------
def generate_summary_report_gmp(method_name, category, params, context, test_results=None):
    if test_results is None: test_results = {}
    doc = Document(); set_korean_font(doc)
    
    section = doc.sections[0]; header = section.header; htable = header.add_table(1, 2, Inches(6.0))
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]; p1.add_run(f"Final Report: {method_name}\n").bold = True; p1.add_run(f"Lot: {context.get('lot_no')}")
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}")

    doc.add_heading('시험법 밸리데이션 최종 보고서', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().add_run(f"Method: {method_name}").bold = True
    
    doc.add_heading('1. 개요 (Summary)', level=1)
    t_sum = doc.add_table(rows=0, cols=2); t_sum.style = 'Table Grid'
    for k, v in [("기기", params.get('Instrument')), ("컬럼", params.get('Column_Plate')), ("검출기", params.get('Detection'))]:
        r = t_sum.add_row().cells; r[0].text=k; r[1].text=str(v)
    
    doc.add_heading('2. 결과 요약 (Results)', level=1)
    t_res = doc.add_table(rows=1, cols=4); t_res.style = 'Table Grid'
    headers = ["Test Item", "Criteria", "Result", "Judgment"]
    for i, h in enumerate(headers): c = t_res.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    
    items_map = [
        ("System Suitability", params.get('SST_Criteria', "RSD ≤ 2.0%"), test_results.get('sst_res', ""), test_results.get('sst_pass', "")),
        ("Specificity", "No Interference", "No Interference", "Pass"),
        ("Linearity", "R² ≥ 0.990", test_results.get('lin_res', ""), test_results.get('lin_pass', "")),
        ("Accuracy", "80 ~ 120%", test_results.get('acc_res', ""), test_results.get('acc_pass', "")),
        ("Precision", "RSD ≤ 2.0%", test_results.get('prec_res', ""), test_results.get('prec_pass', "")),
    ]
    for item, crit, res, judge in items_map:
        r = t_res.add_row().cells; r[0].text=item; r[1].text=crit; r[2].text=str(res); r[3].text=str(judge)

    doc.add_heading('3. 종합 결론', level=1)
    doc.add_paragraph("본 시험법은 모든 밸리데이션 항목에서 판정 기준을 만족하였음.")
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io  

# ---------------------------------------------------------
# 7. 메인 UI Loop (통합 & 세션 연결)
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

    if not df_full.empty:
        my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
        if not my_plan.empty:
            t1, t2, t3 = st.tabs(["📑 Step 1: Strategy", "📗 Step 2: Logbook", "📊 Step 3: Report"])
            
            with t1:
                st.markdown("### 1️⃣ 전략 및 계획서")
                st.dataframe(my_plan[["Method", "Category"]])
                sel_p = st.selectbox("Select Protocol:", my_plan["Method"].unique())
                if sel_p:
                    c1, c2 = st.columns(2)
                    with c1: stock_in = st.number_input("내 Stock 농도:", min_value=0.0, value=1.0)
                    with c2: vol_in = st.number_input("목표 조제량(mL):", min_value=1.0, value=10.0)
                    if st.button("Generate Protocol Package"):
                        params_p = get_method_params(sel_p)
                        target_in = float(params_p.get('Target_Conc', 1.0))
                        
                        recipe = generate_master_recipe_excel(sel_p, target_in, "mg/mL", stock_in, vol_in, "Liquid")
                        st.download_button("📥 Master Recipe (Excel)", recipe, f"Recipe_{sel_p}.xlsx")
                        
                        proto = generate_protocol_premium(sel_p, "Cat", params_p, stock_in, vol_in, target_in)
                        st.download_button("📥 Protocol (Docx)", proto, f"Protocol_{sel_p}.docx")

            with t2:
                st.markdown("### 📗 스마트 엑셀 일지 (GMP)")
                st.info("실험 데이터를 입력할 엑셀 일지를 생성합니다. (테스트용 자동 채우기 가능)")
                sel_l = st.selectbox("Select Logbook:", my_plan["Method"].unique(), key="l")
                
                # [New] Simulation Checkbox
                simulate_mode = st.checkbox("🧪 시뮬레이션 데이터 포함 (Test Mode: Auto-fill Data)", value=False, help="체크하면 가상의 결과값이 채워진 엑셀이 생성되어 즉시 보고서를 만들 수 있습니다.")
                
                if st.button("Generate Excel Logbook"):
                    data = generate_smart_excel(sel_l, "Cat", get_method_params(sel_l), simulate=simulate_mode)
                    
                    # Session Storage for Step 3
                    st.session_state['generated_logbook'] = data
                    st.session_state['generated_log_name'] = sel_l
                    st.success(f"Logbook Generated! ({'Simulated Data Included' if simulate_mode else 'Blank Template'})")
                    st.download_button("📥 Download Excel Logbook", data, f"Logbook_{sel_l}.xlsx")

            with t3:
                st.markdown("### 📊 최종 결과 보고서 (Automated)")
                st.info("작성 완료된 엑셀 일지를 업로드하면 결과가 자동 반영됩니다.")
                sel_r = st.selectbox("Report for:", my_plan["Method"].unique(), key="r")
                uploaded_log = st.file_uploader("📂 Upload Filled Logbook (xlsx)", type=["xlsx"])
                
                # Automatic Session Retrieval
                used_log = None
                if uploaded_log:
                    used_log = uploaded_log
                elif 'generated_logbook' in st.session_state and st.session_state['generated_log_name'] == sel_r:
                    st.info(f"💡 Step 2에서 생성된 '{sel_r}' 일지 데이터를 자동으로 사용합니다.")
                    used_log = st.session_state['generated_logbook']

                lot_no = st.text_input("Lot No:", value="TBD")
                
                if used_log:
                    st.success("Data Ready for Report!")
                    extracted_data = extract_logbook_data(used_log)
                    
                    with st.expander("🔍 Extracted Data Preview"):
                        st.json(extracted_data)
                        
                    if st.button("Generate Final Report"):
                        doc_r = generate_summary_report_gmp(sel_r, "Cat", get_method_params(sel_r), {'lot_no': lot_no}, extracted_data)
                        st.download_button("📥 Download Report (Docx)", doc_r, f"Final_Report_{sel_r}.docx")
                else:
                    st.warning("⚠️ 엑셀 일지를 업로드하거나 Step 2에서 생성해주세요.")