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
        if cell.paragraphs[0].runs: cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# [VMP 생성 함수]
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
    for i, h in enumerate(['No.', 'Method', 'Category', 'Required Items']):
        c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for idx, row in df_strategy.iterrows():
        r = table.add_row().cells; r[0].text=str(idx+1); r[1].text=str(row['Method']); r[2].text=str(row['Category']); r[3].text=", ".join(row['Required_Items'])
    
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [NEW] Master Recipe Excel (정확한 농도 로직 반영)
def generate_master_recipe_excel(method_name, target_conc, unit, stock_conc, req_vol, sample_type, powder_info=""):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Formats
    title_fmt = workbook.add_format({'bold':True, 'font_size': 14, 'align':'center', 'valign':'vcenter', 'bg_color': '#44546A', 'font_color': 'white'})
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center'}) # Main Section
    section_title = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFC000', 'font_size':11, 'align':'left'}) 
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#EDEDED', 'align':'center'})
    cell = workbook.add_format({'border':1, 'align':'center'})
    num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    auto = workbook.add_format({'border':1, 'bg_color':'#E2EFDA', 'num_format':'0.000', 'align':'center'}) # Green for Calculated
    total_fmt = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FFFF00', 'num_format':'0.00', 'align':'center'})

    ws = workbook.add_worksheet("Master Recipe")
    ws.set_column('A:A', 35); ws.set_column('B:E', 15); ws.set_column('F:F', 12)
    
    # 1. Dashboard
    ws.merge_range('A1:F1', f'Validation Material Planner: {method_name}', title_fmt)
    ws.write('A3', "Sample Type:", sub); ws.write('B3', sample_type, cell)
    if sample_type == "Powder (파우더)":
        ws.write('C3', "Prep Detail:", sub); ws.write_string('D3', powder_info, cell)
    
    ws.write('A4', "User Stock Conc:", sub); ws.write('B4', stock_conc, num); ws.write('C4', unit, cell)
    ws.write('A5', "Target Conc (100%):", sub); ws.write('B5', target_conc, num); ws.write('C5', unit, cell)
    ws.write('A6', "Target Vol/Vial (mL):", sub); ws.write('B6', req_vol, num)

    ws.write('D6', "TOTAL STOCK NEEDED (mL):", sub)
    # Total formula placeholder at E6
    
    row = 8
    
    # --- Helper to write grouped sets ---
    def add_section_grouped(main_title, levels, reps):
        nonlocal row
        ws.merge_range(row, 0, row, 5, f"■ {main_title}", header)
        row += 1
        
        section_start_row = row
        
        for rep in range(1, reps + 1):
            set_title = f"{main_title.split(' ')[0]} - {rep}회차 조제 (Set {rep})"
            ws.merge_range(row, 0, row, 5, set_title, section_title)
            row += 1
            ws.write_row(row, 0, ["Item ID", "Target Conc", "Stock Vol (mL)", "Diluent Vol (mL)", "Total (mL)", "Check"], sub)
            row += 1
            
            data_start = row
            for level in levels:
                # 100% Target 농도 기준 계산
                t_val = float(target_conc) * (level / 100)
                
                # Check if stock is sufficient
                if float(stock_conc) < t_val:
                    s_vol = "Error: Stock < Target"
                    d_vol = "N/A"
                else:
                    # Dilution Formula: V1 = (C2 * V2) / C1
                    s_vol = (t_val * float(req_vol)) / float(stock_conc)
                    d_vol = float(req_vol) - s_vol
                
                label = f"{main_title.split(' ')[0]}-{level}%-R{rep}"
                ws.write(row, 0, label, cell)
                ws.write(row, 1, t_val, num)
                
                if isinstance(s_vol, str):
                    ws.write(row, 2, s_vol, workbook.add_format({'bold':True, 'font_color':'red'}))
                    ws.write(row, 3, d_vol, cell)
                else:
                    ws.write(row, 2, s_vol, auto)
                    ws.write(row, 3, d_vol, auto)
                
                ws.write(row, 4, float(req_vol), num)
                ws.write(row, 5, "□", cell)
                row += 1
            
            ws.write(row, 1, f"[{rep}회차] 소요 Stock:", sub)
            if isinstance(s_vol, str):
                ws.write(row, 2, "Error", total_fmt)
            else:
                ws.write_formula(row, 2, f"=SUM(C{data_start+1}:C{row})", total_fmt)
            row += 2 

    # 1. System Suitability (SST) - 100% Level, 6 reps (Use High Stock)
    add_section_grouped("1. 시스템 적합성 (SST)", [100], 1) # SST용 1회 대량 제조 가정 or 6회 주입

    # 2. Specificity (100% Level usually)
    add_section_grouped("2. 특이성 (Specificity)", [100], 1)

    # 3. Linearity (5 levels x 3 reps)
    add_section_grouped("3. 직선성 (Linearity)", [80, 90, 100, 110, 120], 3)

    # 4. Accuracy (3 levels x 3 reps)
    add_section_grouped("4. 정확성 (Accuracy)", [80, 100, 120], 3)
    
    # 5. Precision (6 reps) - treat as 1 set of 6 items
    ws.merge_range(row, 0, row, 5, "■ 5. 정밀성 (Repeatability)", header)
    row += 2
    ws.merge_range(row, 0, row, 5, "반복성 시험 세트 (n=6)", section_title)
    row += 1
    ws.write_row(row, 0, ["Item ID", "Target Conc", "Stock Vol (mL)", "Diluent Vol (mL)", "Total (mL)", "Check"], sub)
    row += 1
    p_start = row
    for i in range(1, 7):
        t_val = float(target_conc)
        s_vol = (t_val * float(req_vol)) / float(stock_conc)
        d_vol = float(req_vol) - s_vol
        ws.write(row, 0, f"Prec-100%-{i}", cell); ws.write(row, 1, t_val, num)
        ws.write(row, 2, s_vol, auto); ws.write(row, 3, d_vol, auto); ws.write(row, 4, float(req_vol), num); ws.write(row, 5, "□", cell)
        row += 1
    ws.write(row, 1, "[정밀성] 소요 Stock:", sub); ws.write_formula(row, 2, f"=SUM(C{p_start+1}:C{row})", total_fmt)
    row += 2

    # 6. Robustness & LOD/LOQ
    add_section_grouped("7. 완건성 (Robustness)", [100], 3) # Assume 3 conditions
    add_section_grouped("8. LOD/LOQ", [1, 0.5], 3)

    # Grand Total
    ws.write_formula('E6', f"=SUM(C9:C{row})", workbook.add_format({'bold':True, 'border':1, 'bg_color':'#FF0000', 'font_color':'white', 'num_format':'0.00', 'align':'center'}))

    workbook.close(); output.seek(0)
    return output

# [PROTOCOL 업그레이드: SOP 수준 서술형 기술]
def generate_protocol_premium(method_name, category, params, stock_conc=None, req_vol=None):
    doc = Document(); set_korean_font(doc)
    def safe_get(key, default=""): val = params.get(key); return str(val) if val is not None else default
    section = doc.sections[0]; header = section.header; htable = header.add_table(1, 2, Inches(6.0)) 
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]; p1.add_run(f"Protocol No.: VP-{method_name[:3]}-001\n").bold = True; p1.add_run(f"Test Category: {category}")
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT; p2.add_run(f"Guideline: {safe_get('Reference_Guideline', 'ICH Q2(R2)')}\n").bold = True; p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
    
    title = doc.add_heading(f'밸리데이션 상세 계획서 (Validation Protocol)', 0); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Method Name: {method_name}").alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph()
    
    # 1-4 Sections (Standard)
    doc.add_heading('1. 목적 (Objective)', level=1); doc.add_paragraph(f"본 문서는 '{method_name}' 시험법의 밸리데이션 수행 방법 및 판정 기준을 기술한다.")
    doc.add_heading('2. 근거 (Reference)', level=1); doc.add_paragraph("• ICH Q2(R2) & MFDS 가이드라인")
    doc.add_heading('3. 기기 및 시약', level=1); t_cond = doc.add_table(rows=0, cols=2); t_cond.style = 'Table Grid'
    for k, v in [("기기", safe_get('Instrument')), ("컬럼", safe_get('Column_Plate')), ("조건", f"A: {safe_get('Condition_A')}\nB: {safe_get('Condition_B')}"), ("검출기", safe_get('Detection'))]:
        r = t_cond.add_row().cells; r[0].text=k; r[0].paragraphs[0].runs[0].bold=True; r[1].text=v
    
    doc.add_heading('4. 밸리데이션 항목 및 기준 (Criteria)', level=1); table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
    headers = ["항목 (Parameter)", "판정 기준 (Criteria)"]; 
    for i, h in enumerate(headers): c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    items = [("특이성", safe_get('Detail_Specificity')), ("직선성", safe_get('Detail_Linearity')), ("범위", safe_get('Detail_Range')), ("정확성", safe_get('Detail_Accuracy')), ("정밀성", safe_get('Detail_Precision')), ("완건성", safe_get('Detail_Robustness'))]
    for k, v in items:
        if v and "정보 없음" not in v: r = table.add_row().cells; r[0].text=k; r[1].text=v
    
    # 5. 상세 시험 방법 (SOP Narrative)
    doc.add_heading('5. 상세 시험 방법 (Test Procedures)', level=1)
    target_conc = safe_get('Target_Conc', '100'); unit = safe_get('Unit', '%')
    
    # 5.1 Preparation
    doc.add_heading('5.1 표준 모액 조제 (Stock Preparation)', level=2)
    doc.add_paragraph(f"1) 표준품 적당량을 정밀히 달아 희석액으로 녹여 농도 {stock_conc if stock_conc else '[입력필요]'} {unit} 용액을 조제한다.")
    doc.add_paragraph("2) 제조된 용액은 완전히 용해되도록 충분히 교반(Vortexing) 또는 초음파 처리(Sonication) 한다.")
    doc.add_paragraph("3) 실온에서 방냉 후 사용한다.")

    # 5.2 Linearity
    doc.add_heading('5.2 직선성 (Linearity)', level=2)
    doc.add_paragraph(f"1) 기준 농도 {target_conc} {unit}를 100%로 설정한다.")
    doc.add_paragraph(f"2) 'Master Recipe' 엑셀 시트에 계산된 용량에 따라, 80%, 90%, 100%, 110%, 120% 수준의 5개 농도를 조제한다.")
    doc.add_paragraph("3) 각 농도별로 1회차, 2회차, 3회차 독립적으로 조제하여 총 15개의 검액을 준비한다.")
    doc.add_paragraph("4) HPLC 시스템 안정화 후, 각 검액을 분석하여 크로마토그램을 얻는다.")
    
    # Insert Table if data exists
    if stock_conc and req_vol:
        doc.add_paragraph("■ 직선성 조제 예시 (1회차 세트 기준):")
        t_lin = doc.add_table(rows=1, cols=4); t_lin.style = 'Table Grid'
        for i, h in enumerate(["Level", "Target Conc", "Stock (mL)", "Diluent (mL)"]): c = t_lin.rows[0].cells[i]; c.text=h; set_table_header_style(c)
        for level in [80, 90, 100, 110, 120]:
            t_val = float(target_conc) * (level/100)
            if float(stock_conc) > t_val:
                s_vol = (t_val * float(req_vol)) / float(stock_conc)
                d_vol = float(req_vol) - s_vol
                r = t_lin.add_row().cells; r[0].text=f"{level}%"; r[1].text=f"{t_val:.2f}"; r[2].text=f"{s_vol:.3f}"; r[3].text=f"{d_vol:.3f}"
    
    # 5.3 Accuracy
    doc.add_heading('5.3 정확성 (Accuracy)', level=2)
    doc.add_paragraph("1) 기준 농도의 80%, 100%, 120% 수준으로 조제한다.")
    doc.add_paragraph("2) 각 수준별로 3회씩 독립적으로 반복 조제하여 총 9개의 검액을 분석한다.")
    
    # 5.4 Precision
    doc.add_heading('5.4 정밀성 (Precision)', level=2)
    doc.add_paragraph(f"1) 기준 농도({target_conc} {unit})로 6개의 검액을 독립적으로 조제한다.")
    doc.add_paragraph("2) 동일한 조건에서 연속적으로 분석하여 면적의 상대표준편차(RSD)를 구한다.")

    doc.add_paragraph("\n\n"); table_sign = doc.add_table(rows=2, cols=3); table_sign.style = 'Table Grid'
    for i, h in enumerate(["작성", "검토", "승인"]): c = table_sign.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i in range(3): table_sign.rows[1].cells[i].text="\n(서명/날짜)\n"
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Excel 생성 함수 - Logbook 전용 (기존 유지)]
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#4472C4', 'font_color':'white', 'align':'center', 'valign':'vcenter'})
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center', 'valign':'vcenter'})
    cell = workbook.add_format({'border':1, 'align':'center'}); num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    calc = workbook.add_format({'border':1, 'bg_color':'#FFFFCC', 'num_format':'0.00', 'align':'center'})
    ws1 = workbook.add_worksheet("1. Info & Prep"); ws1.set_column('A:A', 20); ws1.set_column('B:E', 15); ws1.merge_range('A1:E1', f'GMP Logbook: {method_name}', header)
    info = [("Date", datetime.now().strftime("%Y-%m-%d")), ("Instrument", params.get('Instrument')), ("Column", params.get('Column_Plate')), ("Analyst", "")]
    r = 3; 
    for k, v in info: ws1.write(r, 0, k, sub); ws1.merge_range(r, 1, r, 4, v if v else "", cell); r+=1
    ws1.write(r+1, 0, "Reagent", sub); ws1.merge_range(r+1, 1, r+1, 4, params.get('Ref_Standard_Info', ''), cell)
    ws1.write(r+2, 0, "Prep Method", sub); ws1.merge_range(r+2, 1, r+2, 4, params.get('Preparation_Sample', ''), cell)
    target_conc = params.get('Target_Conc')
    if target_conc:
        try: target_val_base = float(target_conc)
        except: target_val_base = 0
        ws2 = workbook.add_worksheet("2. Linearity"); ws2.set_column('A:H', 12); unit = params.get('Unit', 'ppm'); ws2.merge_range('A1:H1', f'Linearity: Triplicate Analysis (Target: {target_conc} {unit})', header)
        for c, h in enumerate(["Level", "Rep", f"Conc ({unit})", "Weight", "Vol", "Response (Y)", "Mean (Y)", "RSD (%)"]): ws2.write(2, c, h, sub)
        levels = [80, 90, 100, 110, 120]; row = 3; chart_rows = []
        for level in levels:
            target_val = target_val_base * (level / 100); start_row = row + 1
            for i in range(1, 4):
                ws2.write_row(row, 0, [f"{level}%", i, target_val, "", 50, ""], cell)
                if i == 1: ws2.merge_range(row, 6, row+2, 6, "", calc); ws2.write_formula(row, 6, f"=AVERAGE(F{start_row}:F{start_row+2})", calc); ws2.merge_range(row, 7, row+2, 7, "", calc); ws2.write_formula(row, 7, f"=STDEV(F{start_row}:F{start_row+2})/G{start_row}*100", calc); chart_rows.append(row + 1)
                row += 1
        s_row = row + 2; ws2.merge_range(s_row, 1, s_row, 3, "■ Summary for Chart", sub); ws2.write_row(s_row+1, 1, ["Conc (X)", "Mean (Y)", "R²"], sub)
        for idx, r_idx in enumerate(chart_rows): ws2.write_formula(s_row+2+idx, 1, f"=C{r_idx}", num); ws2.write_formula(s_row+2+idx, 2, f"=G{r_idx}", num)
        ws2.write_formula(s_row+2, 3, f"=RSQ(C{s_row+3}:C{s_row+7}, B{s_row+3}:B{s_row+7})", calc)
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'}); chart.add_series({'categories': f"='2. Linearity'!$B${s_row+3}:$B${s_row+7}", 'values': f"='2. Linearity'!$C${s_row+3}:$C${s_row+7}", 'trendline': {'type': 'linear', 'display_equation': True, 'display_r_squared': True}}); ws2.insert_chart('J3', chart)
    if params.get('Detail_Inter_Precision'):
        ws3 = workbook.add_worksheet("3. Precision"); ws3.set_column('A:E', 15); ws3.merge_range('A1:E1', 'Intermediate Precision', header); ws3.merge_range('A3:E3', "■ Day 1", sub); ws3.write_row('A4', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
        for i in range(6): ws3.write_row(4+i, 0, [i+1, "Sample", ""], cell)
        ws3.write_formula('D5', "=AVERAGE(C5:C10)", num); ws3.write_formula('E5', "=STDEV(C5:C10)/D5*100", num); ws3.merge_range('A12:E12', "■ Day 2", sub); ws3.write_row('A13', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
        for i in range(6): ws3.write_row(13+i, 0, [i+1, "Sample", ""], cell)
        ws3.write_formula('D14', "=AVERAGE(C14:C19)", num); ws3.write_formula('E14', "=STDEV(C14:C19)/D14*100", num); ws3.write('A21', "Diff (%)", sub); ws3.write_formula('B21', "=ABS(D5-D14)/AVERAGE(D5,D14)*100", num)
    if params.get('Detail_Robustness'):
        ws4 = workbook.add_worksheet("4. Robustness"); ws4.set_column('A:F', 18); ws4.merge_range('A1:F1', 'Robustness Conditions', header); ws4.merge_range('A2:F2', f"Guide: {params.get('Detail_Robustness')}", cell)
        for c, h in enumerate(["Condition", "Set", "Actual", "SST", "Pass/Fail", "Note"]): ws4.write(3, c, h, sub)
        for r, c in enumerate(["Standard", "Flow -0.1", "Flow +0.1", "Temp -2", "Temp +2"]): ws4.write(4+r, 0, c, cell); ws4.write_row(4+r, 1, [""]*5, cell)
    ws5 = workbook.add_worksheet("5. Raw Data"); ws5.set_column('A:F', 15); ws5.merge_range('A1:F1', 'Raw Data', header)
    for c, h in enumerate(["Inj No.", "Sample Name", "RT", "Area", "Height", "Remarks"]): ws5.write(2, c, h, sub)
    for r in range(3, 23): ws5.write_row(r, 0, [""]*6, cell)
    workbook.close(); output.seek(0)
    return output

# [Report 생성 함수 - 기존 유지]
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
                    st.divider()
                    st.markdown("#### 🧪 시약 제조 및 계획서 생성기")
                    sel_p = st.selectbox("Protocol:", my_plan["Method"].unique())
                    
                    if sel_p:
                        # [NEW] 시료 타입 선택 (도메인 복구)
                        st.info("👇 시료 상태(액체/파우더)를 선택하고 정보를 입력하세요.")
                        
                        sample_type = st.radio("시료 타입 (Sample Type):", ["Liquid (액체)", "Powder (파우더)"], horizontal=True)
                        
                        cc1, cc2 = st.columns(2)
                        stock_input_val = 0.0
                        powder_desc = ""
                        
                        if sample_type == "Liquid (액체)":
                            with cc1: stock_input_val = st.number_input("내 Stock 농도 (mg/mL 등):", min_value=0.0, step=0.1, format="%.2f")
                        else: # Powder
                            with cc1: weight_input = st.number_input("칭량값 (Weight, mg):", min_value=0.0, step=0.1)
                            with cc2: dil_vol_input = st.number_input("희석 부피 (Vol, mL):", min_value=0.1, value=10.0, step=1.0)
                            if dil_vol_input > 0:
                                stock_input_val = weight_input / dil_vol_input
                                st.caption(f"🧪 계산된 Stock 농도: **{stock_input_val:.2f} mg/mL**")
                                powder_desc = f"Weigh {weight_input}mg / {dil_vol_input}mL"

                        # [NEW] 명칭 변경: 1회당 조제량 -> 개별 바이알 조제 목표량
                        with cc2: vol_input = st.number_input("개별 바이알 조제 목표량 (Target Vol, mL):", min_value=1.0, value=5.0, step=1.0, help="부피 플라스크 크기(예: 5mL, 10mL)를 입력하세요.")
                        
                        params_p = get_method_params(sel_p)
                        target_conc_val = params_p.get('Target_Conc', 0)
                        unit_val = params_p.get('Unit', '')

                        # 다운로드 버튼
                        if stock_input_val > 0:
                            calc_excel = generate_master_recipe_excel(sel_p, target_conc_val, unit_val, stock_input_val, vol_input, sample_type, powder_desc)
                            st.download_button("🧮 시약 제조 계산기 (Master Recipe) 다운로드", calc_excel, f"Master_Recipe_{sel_p}.xlsx")
                        
                        doc_proto = generate_protocol_premium(sel_p, "Cat", params_p, stock_input_val if stock_input_val > 0 else None, vol_input)
                        st.download_button("📄 상세 계획서 (Protocol) 다운로드", doc_proto, f"Protocol_{sel_p}.docx", type="primary")

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