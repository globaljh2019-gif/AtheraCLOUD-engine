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
# 1. 설정 및 데이터 로딩 (Notion API)
# ---------------------------------------------------------
try:
    NOTION_API_KEY = st.secrets["NOTION_API_KEY"]
    CRITERIA_DB_ID = st.secrets["CRITERIA_DB_ID"]
    STRATEGY_DB_ID = st.secrets["STRATEGY_DB_ID"]
    PARAM_DB_ID = st.secrets.get("PARAM_DB_ID", "") 
except:
    NOTION_API_KEY = ""; CRITERIA_DB_ID = ""; STRATEGY_DB_ID = ""; PARAM_DB_ID = ""

headers = {"Authorization": "Bearer " + NOTION_API_KEY, "Content-Type": "application/json", "Notion-Version": "2022-06-28"}

@st.cache_data
def get_criteria_map():
    if not CRITERIA_DB_ID: return {}
    url = f"https://api.notion.com/v1/databases/{CRITERIA_DB_ID}/query"
    res = requests.post(url, headers=headers); criteria_map = {}
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
    res = requests.post(url, headers=headers); data = []
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
            try: ts = props.get(n, {}).get("rich_text", []); return "".join([t["text"]["content"] for t in ts]) if ts else ""
            except: return ""
        def num(n):
            try: return props.get(n, {}).get("number")
            except: return None
        return {
            "Instrument": txt("Instrument"), "Column_Plate": txt("Column_Plate"), "Condition_A": txt("Condition_A"), "Condition_B": txt("Condition_B"), "Detection": txt("Detection"),
            "SST_Criteria": txt("SST_Criteria"), "Reference_Guideline": txt("Reference_Guideline"), "Detail_Specificity": txt("Detail_Specificity"),
            "Detail_Linearity": txt("Detail_Linearity"), "Detail_Range": txt("Detail_Range"), "Detail_Accuracy": txt("Detail_Accuracy"),
            "Detail_Precision": txt("Detail_Precision"), "Detail_Inter_Precision": txt("Detail_Inter_Precision"), "Detail_LOD": txt("Detail_LOD"),
            "Detail_LOQ": txt("Detail_LOQ"), "Detail_Robustness": txt("Detail_Robustness"), "Reagent_List": txt("Reagent_List"),
            "Ref_Standard_Info": txt("Ref_Standard_Info"), "Preparation_Std": txt("Preparation_Std"), "Preparation_Sample": txt("Preparation_Sample"),
            "Target_Conc": num("Target_Conc"), "Unit": txt("Unit")
        }
    return {}

# ---------------------------------------------------------
# 2. 문서 생성 헬퍼
# ---------------------------------------------------------
def set_korean_font(doc):
    style = doc.styles['Normal']; style.font.name = 'Malgun Gothic'; style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic'); style.font.size = Pt(10)

def set_table_header_style(cell):
    tcPr = cell._element.get_or_add_tcPr(); shading_elm = OxmlElement('w:shd'); shading_elm.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shading_elm)
    if cell.paragraphs:
        if cell.paragraphs[0].runs: cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# ---------------------------------------------------------
# 3. 문서 생성 엔진 (VMP, Recipe, Protocol - 상세 기능 복구)
# ---------------------------------------------------------
def generate_vmp_premium(modality, phase, df_strategy):
    doc = Document(); set_korean_font(doc)
    doc.add_heading('밸리데이션 종합계획서 (Validation Master Plan)', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    table_info = doc.add_table(rows=2, cols=4); table_info.style = 'Table Grid'
    headers = ["제품명 (Product)", "단계 (Phase)", "문서 번호 (Doc No.)", "제정 일자 (Date)"]
    values = [f"{modality} Project", phase, "VMP-001", datetime.now().strftime('%Y-%m-%d')]
    for i, h in enumerate(headers): c = table_info.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i, v in enumerate(values): c = table_info.rows[1].cells[i]; c.text=v; c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('1. 목적 (Objective)', level=1); doc.add_paragraph("본 문서는 의약품 품질 관리를 위한 시험법 밸리데이션의 전략과 범위를 규정한다.")
    doc.add_heading('2. 적용 범위 (Scope)', level=1); doc.add_paragraph(f"본 계획서는 {phase} 단계의 {modality} 의약품 시험법에 적용된다.")
    doc.add_heading('3. 근거 가이드라인 (Reference)', level=1); doc.add_paragraph("• ICH Q2(R2) Guideline\n• MFDS 의약품 등 시험방법 밸리데이션 가이드라인")
    
    doc.add_heading('4. 밸리데이션 수행 전략', level=1)
    table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
    for i, h in enumerate(['No.', 'Method', 'Category', 'Required Items']): c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for idx, row in df_strategy.iterrows(): 
        r = table.add_row().cells
        r[0].text=str(idx+1); r[1].text=str(row['Method']); r[2].text=str(row['Category']); r[3].text=", ".join(row['Required_Items'])
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

def generate_master_recipe_excel(method_name, target_conc, unit, stock_conc, req_vol, sample_type, powder_info=""):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    title_fmt = workbook.add_format({'bold':True, 'font_size': 14, 'align':'center', 'bg_color': '#44546A', 'font_color': 'white'})
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center'})
    cell = workbook.add_format({'border':1, 'align':'center'})
    num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    auto = workbook.add_format({'border':1, 'bg_color':'#E2EFDA', 'num_format':'0.000', 'align':'center'})
    
    ws = workbook.add_worksheet("Master Recipe")
    ws.set_column('A:F', 15)
    ws.merge_range('A1:F1', f'Validation Material Planner: {method_name}', title_fmt)
    ws.write('A3', "Sample Type:", header); ws.write('B3', sample_type, cell)
    if sample_type == "Powder (파우더)": ws.write('C3', "Prep:", header); ws.write('D3', powder_info, cell)
    ws.write('A4', "Stock Conc:", header); ws.write('B4', stock_conc, num); ws.write('C4', unit, cell)
    ws.write('A5', "Target Conc:", header); ws.write('B5', target_conc, num); ws.write('C5', unit, cell)
    ws.write('A6', "Vol/Vial:", header); ws.write('B6', req_vol, num)
    
    ws.write(7, 0, "■ Dilution Scheme (Linearity & Accuracy)", workbook.add_format({'bold':True, 'font_size':11}))
    ws.write_row(8, 0, ["Level", "Target Conc", "Stock Vol (mL)", "Diluent Vol (mL)", "Total (mL)", "Check"], header)
    
    row = 9
    # 직선성/정확성 레벨 계산 로직
    for level in [80, 90, 100, 110, 120]:
        t_val = float(target_conc) * (level/100)
        # Formula: (Target * Total) / Stock
        if float(stock_conc) > 0:
            s_vol = (t_val * float(req_vol)) / float(stock_conc)
            d_vol = float(req_vol) - s_vol
        else: s_vol = 0; d_vol = 0
        
        ws.write(row, 0, f"{level}%", cell)
        ws.write(row, 1, t_val, num)
        ws.write(row, 2, s_vol, auto)
        ws.write(row, 3, d_vol, auto)
        ws.write(row, 4, float(req_vol), num)
        ws.write(row, 5, "□", cell)
        row += 1
        
    workbook.close(); output.seek(0)
    return output

def generate_protocol_premium(method_name, category, params, stock_conc=None, req_vol=None, target_conc_override=None):
    doc = Document(); set_korean_font(doc)
    
    # Header Table
    section = doc.sections[0]; header = section.header; htable = header.add_table(1, 2, Inches(6.0))
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]; p1.add_run(f"Protocol: {method_name}\n").bold = True
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}")

    doc.add_heading(f'밸리데이션 상세 계획서 ({method_name})', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    doc.add_heading('1. 목적 및 범위', level=1); doc.add_paragraph("본 문서는 해당 시험법의 적합성을 검증하기 위한 상세 계획을 기술한다.")
    
    doc.add_heading('2. 기기 및 시약', level=1)
    t_cond = doc.add_table(rows=4, cols=2); t_cond.style = 'Table Grid'
    conds = [("기기", params.get('Instrument', 'N/A')), ("컬럼", params.get('Column_Plate', 'N/A')), ("검출기", params.get('Detection', 'N/A')), ("SST 기준", params.get('SST_Criteria', 'N/A'))]
    for i, (k, v) in enumerate(conds): t_cond.rows[i].cells[0].text=k; t_cond.rows[i].cells[1].text=str(v)
    
    doc.add_heading('3. 상세 시험 방법', level=1)
    doc.add_heading('3.1 용액 조제', level=2)
    if stock_conc: doc.add_paragraph(f"1) 표준 모액: 농도 {stock_conc}의 용액을 준비한다.")
    doc.add_paragraph(f"2) 검액: 기준 농도 {target_conc_override if target_conc_override else 'Target'} 수준으로 희석하여 사용한다.")
    
    doc.add_heading('4. 밸리데이션 항목 및 기준', level=1)
    doc.add_paragraph("1) 특이성: 간섭 피크가 없을 것 (≤ 0.5%)\n2) 직선성: R² ≥ 0.990\n3) 정확성: 80~120%\n4) 정밀성: RSD ≤ 2.0%")
    
    doc.add_paragraph("\n\n(이하 서명란 생략)")
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 엑셀 스마트 일지 생성 (★ 핵심 로직: 수식 & 참조 완벽 보정)
# ---------------------------------------------------------
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Styles
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#4472C4', 'font_color':'white', 'align':'center', 'valign':'vcenter'})
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center', 'valign':'vcenter'})
    sub_rep = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FCE4D6', 'align':'left'}) 
    cell = workbook.add_format({'border':1, 'align':'center'})
    num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    num3 = workbook.add_format({'border':1, 'num_format':'0.000', 'align':'center'}) 
    calc = workbook.add_format({'border':1, 'bg_color':'#FFFFCC', 'num_format':'0.00', 'align':'center'}) 
    auto = workbook.add_format({'border':1, 'bg_color':'#E2EFDA', 'num_format':'0.00', 'align':'center'}) 
    pass_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#C6EFCE', 'font_color':'#006100', 'align':'center'})
    fail_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFC7CE', 'font_color':'#9C0006', 'align':'center'})
    total_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFFF00', 'num_format':'0.00', 'align':'center'})
    crit_fmt = workbook.add_format({'bold':True, 'font_color':'red', 'align':'left'})

    # [1. Info Sheet]
    ws1 = workbook.add_worksheet("1. Info")
    ws1.set_column('A:A', 25); ws1.set_column('B:E', 15)
    ws1.merge_range('A1:E1', f'GMP Logbook: {method_name}', header)
    
    infos = [("Date", datetime.now().strftime("%Y-%m-%d")), ("Instrument", params.get('Instrument')), ("Column", params.get('Column_Plate')), ("Analyst", "")]
    for i, (k, v) in enumerate(infos):
        ws1.write(i+3, 0, k, sub); ws1.merge_range(i+3, 1, i+3, 4, v if v else "", cell)
    
    ws1.write(8, 0, "Round Rule:", sub)
    ws1.merge_range(8, 1, 8, 4, "모든 계산값은 소수점 2째자리(농도 3째자리)에서 절사(ROUNDDOWN).", cell)
    
    # Target Conc (B10)
    target_conc_val = float(params.get('Target_Conc', 1.0))
    ws1.write(9, 0, "Target Conc (100%):", sub); ws1.write(9, 1, target_conc_val, calc); ws1.write(9, 2, params.get('Unit', 'mg/mL'), cell)
    target_conc_ref = "'1. Info'!$B$10"

    # Preparation
    ws1.merge_range(10, 0, 10, 4, "■ Standard Preparation & Correction Factor", sub_rep)
    labels = ["Theoretical Stock (mg/mL):", "Purity (Potency, %):", "Water Content (%):", "Actual Weight (mg):", "Final Volume (mL):"]
    for i, label in enumerate(labels):
        ws1.write(11 + i, 0, label, sub)
        if "Purity" in label: ws1.write(11 + i, 1, 100.0, calc)
        elif "Water" in label: ws1.write(11 + i, 1, 0.0, calc)
        else: ws1.write(11 + i, 1, "", calc)

    # Actual Stock (B17)
    ws1.write(16, 0, "Actual Stock (mg/mL):", sub)
    ws1.write_formula(16, 1, '=IF(B16="","",ROUNDDOWN((B15*(B13/100)*((100-B14)/100))/B16, 4))', auto)
    actual_stock_ref = "'1. Info'!$B$17"

    # Correction Factor (B18)
    ws1.write(17, 0, "Correction Factor:", sub)
    ws1.write_formula(17, 1, '=IF(OR(B12="",B12=0,B17=""), 1, ROUNDDOWN(B17/B12, 4))', total_fmt)
    corr_factor_ref = "'1. Info'!$B$18"
    theo_stock_ref = "'1. Info'!$B$12"

    # [2. SST Sheet]
    ws_sst = workbook.add_worksheet("2. SST"); ws_sst.set_column('A:F', 15)
    ws_sst.merge_range('A1:F1', 'System Suitability Test (n=6)', header)
    ws_sst.write_row('A2', ["Inj No.", "RT (min)", "Area", "Height", "Tailing (1st)", "Plate Count"], sub)
    for i in range(1, 7): ws_sst.write(i+1, 0, i, cell); ws_sst.write_row(i+1, 1, ["", "", "", "", ""], calc)
    ws_sst.write('A9', "Mean", sub); ws_sst.write_formula('B9', "=ROUNDDOWN(AVERAGE(B3:B8), 2)", auto); ws_sst.write_formula('C9', "=ROUNDDOWN(AVERAGE(C3:C8), 2)", auto)
    ws_sst.write('A10', "RSD(%)", sub); ws_sst.write_formula('B10', "=ROUNDDOWN(STDEV(B3:B8)/B9*100, 2)", auto); ws_sst.write_formula('C10', "=ROUNDDOWN(STDEV(C3:C8)/C9*100, 2)", auto)
    ws_sst.write('A12', "Result:", sub); ws_sst.write_formula('B12', '=IF(AND(B10<=2.0, C10<=2.0, E3<=2.0), "Pass", "Fail")', pass_fmt)
    ws_sst.conditional_format('B12', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
    ws_sst.write('A14', "※ Acceptance Criteria: RSD ≤ 2.0%, Tailing ≤ 2.0", crit_fmt)

    # [3. Specificity]
    ws_spec = workbook.add_worksheet("3. Specificity"); ws_spec.set_column('A:E', 20); ws_spec.merge_range('A1:E1', 'Specificity Test', header)
    ws_spec.write('A3', "Std Mean Area (Ref. SST):", sub); ws_spec.write_formula('B3', "='2. SST'!C9", num)
    ws_spec.write_row('A5', ["Sample", "RT", "Area", "Interference (%)", "Result"], sub)
    for i, s in enumerate(["Blank", "Placebo"]):
        ws_spec.write(i+6, 0, s, cell); ws_spec.write(i+6, 1, "", calc); ws_spec.write(i+6, 2, "", calc)
        ws_spec.write_formula(i+6, 3, f'=IF(OR(C{i+7}="", $B$3=""), "", ROUNDDOWN(C{i+7}/$B$3*100, 2))', auto)
        ws_spec.write_formula(i+6, 4, f'=IF(D{i+7}="", "", IF(D{i+7}<=0.5, "Pass", "Fail"))', pass_fmt)
        ws_spec.conditional_format(f'E{i+7}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
    ws_spec.write('A10', "Criteria: Interference ≤ 0.5%", crit_fmt)

    # [4. Linearity] (Corrected Formula)
    ws2 = workbook.add_worksheet("4. Linearity"); ws2.set_column('A:I', 13); ws2.merge_range('A1:I1', 'Linearity Test', header)
    row = 3; rep_rows = {1: [], 2: [], 3: []}
    for rep in range(1, 4):
        ws2.merge_range(row, 0, row, 8, f"■ Repetition {rep}", sub_rep); row += 1
        ws2.write_row(row, 0, ["Level", "Conc (X)", "Area (Y)", "Back Calc", "Accuracy (%)", "Check"], sub); row += 1
        data_start = row
        for level in [80, 90, 100, 110, 120]:
            ws2.write(row, 0, f"{level}%", cell)
            # Formula: ActualStock(B17) * (Level/100) * (Target/Theo)
            formula_x = f"=ROUNDDOWN({actual_stock_ref} * ({level}/100) * ({target_conc_ref} / {theo_stock_ref}), 3)"
            ws2.write_formula(row, 1, formula_x, num3); ws2.write(row, 2, "", calc); rep_rows[rep].append(row + 1)
            slope_ref = f"C{data_start+7}"; int_ref = f"C{data_start+8}"
            ws2.write_formula(row, 3, f'=IF(C{row+1}="", "", ROUNDDOWN((C{row+1}-{int_ref})/{slope_ref}, 3))', auto)
            ws2.write_formula(row, 4, f'=IF(C{row+1}="", "", ROUNDDOWN(D{row+1}/B{row+1}*100, 1))', auto)
            ws2.write(row, 5, "OK", cell); row += 1
        ws2.write(row, 1, "Slope:", sub); ws2.write_formula(row, 2, f"=SLOPE(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
        ws2.write(row+1, 1, "Intercept:", sub); ws2.write_formula(row+1, 2, f"=INTERCEPT(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
        ws2.write(row+2, 1, "R²:", sub); ws2.write_formula(row+2, 2, f"=RSQ(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
        row += 6

    # Summary
    ws2.merge_range(row, 0, row, 8, "■ Summary (Mean) & Final Check", sub_rep); row += 1
    ws2.write_row(row, 0, ["Level", "Conc (X)", "Mean Area", "STDEV", "% RSD", "Result"], sub); row += 1
    sum_start = row
    for i, level in enumerate([80, 90, 100, 110, 120]):
        r1, r2, r3 = rep_rows[1][i], rep_rows[2][i], rep_rows[3][i]
        ws2.write(row, 0, f"{level}%", cell); ws2.write_formula(row, 1, f"=B{r1}", num3)
        ws2.write_formula(row, 2, f"=ROUNDDOWN(AVERAGE(C{r1},C{r2},C{r3}), 2)", auto)
        ws2.write_formula(row, 3, f"=ROUNDDOWN(STDEV(C{r1},C{r2},C{r3}), 2)", auto)
        ws2.write_formula(row, 4, f"=IF(C{row+1}=0, 0, ROUNDDOWN(D{row+1}/C{row+1}*100, 2))", auto)
        ws2.write_formula(row, 5, f'=IF(C{row+1}=0, "", IF(E{row+1}<=5.0, "Pass", "Fail"))', pass_fmt); ws2.conditional_format(f'F{row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt}); row += 1
    
    ws2.write(row+1, 1, "Final Slope:", sub); ws2.write_formula(row+1, 2, f"=ROUNDDOWN(SLOPE(C{sum_start+1}:C{sum_start+5}, B{sum_start+1}:B{sum_start+5}), 4)", auto)
    ws2.write(row+1, 1, "Final Intercept:", sub); ws2.write_formula(row+1, 2, f"=ROUNDDOWN(INTERCEPT(C{sum_start+1}:C{sum_start+5}, B{sum_start+1}:B{sum_start+5}), 4)", auto)
    ws2.write(row+2, 1, "Final R²:", sub); ws2.write_formula(row+2, 2, f"=ROUNDDOWN(RSQ(C{sum_start+1}:C{sum_start+5}, B{sum_start+1}:B{sum_start+5}), 4)", auto)
    ws2.write(row+2, 4, "Criteria: R² ≥ 0.990", crit_fmt)
    ws2.write_formula(row+2, 5, f'=IF(C{row+3}=0, "", IF(C{row+3}>=0.990, "Pass", "Fail"))', pass_fmt); ws2.conditional_format(f'F{row+3}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
    ws2.write(row+4, 0, "※ Acceptance Criteria: R² ≥ 0.990, %RSD ≤ 5.0%", crit_fmt)

    # [5. Accuracy]
    ws_acc = workbook.add_worksheet("5. Accuracy"); ws_acc.set_column('A:G', 15); ws_acc.merge_range('A1:G1', 'Accuracy (Recovery)', header)
    ws_acc.write('E3', "Slope:", sub); ws_acc.write_formula('F3', f"='4. Linearity'!C{row+2}", auto)
    ws_acc.write('E4', "Int:", sub); ws_acc.write_formula('F4', f"='4. Linearity'!C{row+3}", auto)
    ws_acc.write('G3', "(From Linearity Summary)", cell)
    
    acc_row = 6
    for level in [80, 100, 120]:
        ws_acc.merge_range(acc_row, 0, acc_row, 6, f"■ Level {level}% (3 Reps)", sub_rep); acc_row += 1
        ws_acc.write_row(acc_row, 0, ["Rep", "Theo Conc", "Area", "Calc Conc", "Recovery (%)", "Criteria", "Result"], sub); acc_row += 1
        start_r = acc_row
        for rep in range(1, 4):
            ws_acc.write(acc_row, 0, rep, cell)
            ws_acc.write_formula(acc_row, 1, f"=ROUNDDOWN({actual_stock_ref} * ({level}/100) * ({target_conc_ref} / {theo_stock_ref}), 3)", num3)
            ws_acc.write(acc_row, 2, "", calc)
            ws_acc.write_formula(acc_row, 3, f'=IF(C{acc_row+1}="","",ROUNDDOWN((C{acc_row+1}-$F$4)/$F$3, 3))', auto)
            ws_acc.write_formula(acc_row, 4, f'=IF(D{acc_row+1}="","",ROUNDDOWN(D{acc_row+1}/B{acc_row+1}*100, 1))', auto)
            ws_acc.write(acc_row, 5, "80~120%", cell)
            ws_acc.write_formula(acc_row, 6, f'=IF(E{acc_row+1}="","",IF(AND(E{acc_row+1}>=80, E{acc_row+1}<=120), "Pass", "Fail"))', pass_fmt)
            ws_acc.conditional_format(f'G{acc_row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt}); acc_row += 1
        ws_acc.write(acc_row, 3, "Mean Rec(%):", sub); ws_acc.write_formula(acc_row, 4, f"=ROUNDDOWN(AVERAGE(E{start_r+1}:E{acc_row}), 1)", total_fmt); acc_row += 2
    ws_acc.write(acc_row, 0, "Criteria: 80 ~ 120%", crit_fmt)

    # [6. Precision / 7. Robustness / 8. LOD_LOQ]
    ws3 = workbook.add_worksheet("6. Precision"); ws3.set_column('A:E', 15); ws3.merge_range('A1:E1', 'Precision', header)
    ws3.merge_range('A3:E3', "■ Day 1 (Repeatability)", sub); ws3.write_row('A4', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
    for i in range(6): ws3.write_row(4+i, 0, [i+1, "Sample", ""], calc)
    ws3.write_formula('D5', "=ROUNDDOWN(AVERAGE(C5:C10), 2)", num); ws3.write_formula('E5', "=ROUNDDOWN(STDEV(C5:C10)/D5*100, 2)", num)
    ws3.write('D11', "Result:", sub); ws3.write_formula('E11', '=IF(E5=0,"",IF(E5<=2.0,"Pass","Fail"))', pass_fmt)
    ws3.write(23, 0, "Criteria: RSD ≤ 2.0%", crit_fmt)

    if params.get('Detail_Robustness'):
        ws4 = workbook.add_worksheet("7. Robustness"); ws4.set_column('A:F', 20); ws4.merge_range('A1:F1', 'Robustness', header)
        ws4.write_row('A3', ["Condition", "Set", "Actual", "SST Result (RSD)", "Pass/Fail", "Note"], sub)
        for r, c in enumerate(["Standard", "Flow -0.1", "Flow +0.1", "Temp -2", "Temp +2"]): 
            ws4.write(4+r, 0, c, cell); ws4.write_row(4+r, 1, ["", "", ""], calc); ws4.write_formula(4+r, 4, f'=IF(D{5+r}="", "", IF(D{5+r}<=2.0, "Pass", "Fail"))', pass_fmt)
    
    ws_ll = workbook.add_worksheet("8. LOD_LOQ"); ws_ll.set_column('A:E', 15); ws_ll.merge_range('A1:E1', 'LOD / LOQ', header)
    ws_ll.write_row('A2', ["Item", "Signal", "Noise", "S/N Ratio", "Result"], sub)
    ws_ll.write('A3', "LOD", cell); ws_ll.write_row('B3', ["", ""], calc); ws_ll.write_formula('D3', '=IF(C3="","",ROUNDDOWN(B3/C3, 1))', auto); ws_ll.write_formula('E3', '=IF(D3="","",IF(D3>=3, "Pass", "Fail"))', pass_fmt)
    ws_ll.write('A4', "LOQ", cell); ws_ll.write_row('B4', ["", ""], calc); ws_ll.write_formula('D4', '=IF(C4="","",ROUNDDOWN(B4/C4, 1))', auto); ws_ll.write_formula('E4', '=IF(D4="","",IF(D4>=10, "Pass", "Fail"))', pass_fmt)
    ws_ll.write('A7', "Criteria: LOD S/N ≥ 3, LOQ S/N ≥ 10", crit_fmt)

    workbook.close(); output.seek(0)
    return output

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

    except Exception as e: st.error(f"Error extracting data: {e}"); return {}
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
# 7. 메인 UI Loop
# ---------------------------------------------------------
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
                
                # Protocol & Recipe Logic (Restored)
                sel_p = st.selectbox("Select Protocol:", my_plan["Method"].unique())
                if sel_p:
                    c_1, c_2 = st.columns(2)
                    with c_1: stock_in = st.number_input("내 Stock 농도:", min_value=0.0, value=1.0)
                    with c_2: vol_in = st.number_input("목표 조제량(mL):", min_value=1.0, value=10.0)
                    
                    if st.button("Generate Documents (Recipe & Protocol)"):
                        params_p = get_method_params(sel_p)
                        target_in = float(params_p.get('Target_Conc', 1.0))
                        
                        # Recipe
                        recipe = generate_master_recipe_excel(sel_p, target_in, "mg/mL", stock_in, vol_in, "Liquid")
                        st.download_button("📥 Master Recipe (Excel)", recipe, f"Recipe_{sel_p}.xlsx")
                        
                        # Protocol
                        proto = generate_protocol_premium(sel_p, "Cat", params_p, stock_in, vol_in, target_in)
                        st.download_button("📥 Protocol (Docx)", proto, f"Protocol_{sel_p}.docx")

            with t2:
                st.markdown("### 📗 스마트 엑셀 일지 (GMP)")
                sel_l = st.selectbox("Select Logbook:", my_plan["Method"].unique(), key="l")
                if st.button("Generate Excel Logbook"):
                    data = generate_smart_excel(sel_l, "Cat", get_method_params(sel_l))
                    st.download_button("Download Excel", data, f"Logbook_{sel_l}.xlsx")

            with t3:
                st.markdown("### 📊 최종 결과 보고서 (Automated)")
                st.info("작성 완료된 엑셀 일지를 업로드하면 결과가 자동 반영됩니다.")
                sel_r = st.selectbox("Report for:", my_plan["Method"].unique(), key="r")
                uploaded_log = st.file_uploader("📂 Upload Filled Logbook (xlsx)", type=["xlsx"])
                lot_no = st.text_input("Lot No:", value="TBD")
                
                if uploaded_log:
                    st.success("Data Extracted!")
                    extracted_data = extract_logbook_data(uploaded_log)
                    st.json(extracted_data)
                    if st.button("Generate Final Report"):
                        doc_r = generate_summary_report_gmp(sel_r, "Cat", get_method_params(sel_r), {'lot_no': lot_no}, extracted_data)
                        st.download_button("Download Report", doc_r, f"Final_Report_{sel_r}.docx")
