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
# 1. 설정 및 데이터 로딩 (Notion API)
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

# ---------------------------------------------------------
# [New] 최종 결과 보고서 생성 (Pre-filled Template)
# ---------------------------------------------------------
def generate_summary_report_gmp(method_name, category, params, context):
    doc = Document(); set_korean_font(doc)
    
    # 1. Header Information
    section = doc.sections[0]; header = section.header; htable = header.add_table(1, 2, Inches(6.0))
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]
    p1.add_run(f"Final Report: {method_name}\n").bold = True
    p1.add_run(f"Lot No.: {context.get('lot_no', 'N/A')}")
    
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}\nDoc No.: VR-{method_name[:3]}-001")

    # 2. Title & Approval
    title = doc.add_heading('시험법 밸리데이션 최종 보고서', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"(Method Validation Final Report for {method_name})").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    table_sign = doc.add_table(rows=2, cols=3); table_sign.style = 'Table Grid'
    headers = ["Written By (Analyzed)", "Reviewed By", "Approved By (QA)"]
    for i, h in enumerate(headers): 
        c = table_sign.rows[0].cells[i]; c.text = h; set_table_header_style(c)
    for i in range(3): table_sign.rows[1].cells[i].text = "\n\n(Sign/Date)\n"
    doc.add_paragraph()

    # 3. Objective & Method Summary (자동 입력됨)
    doc.add_heading('1. 개요 및 시험 방법 (Summary)', level=1)
    doc.add_paragraph("본 문서는 해당 시험법의 밸리데이션 결과를 요약하고 적합성을 판정한다.")
    
    t_sum = doc.add_table(rows=0, cols=2); t_sum.style = 'Table Grid'
    summary_data = [
        ("시험명 (Method)", method_name),
        ("시험 목적 (Category)", category),
        ("사용 기기 (Instrument)", params.get('Instrument', 'N/A')),
        ("컬럼 (Column)", params.get('Column_Plate', 'N/A')),
        ("검출 조건 (Detection)", params.get('Detection', 'N/A')),
        ("기준 농도 (Target)", f"{params.get('Target_Conc', '')} {params.get('Unit', '')}")
    ]
    for k, v in summary_data:
        r = t_sum.add_row().cells
        r[0].text = k; r[0].paragraphs[0].runs[0].bold = True
        r[1].text = str(v)
    
    # 4. Validation Results Summary (기준 자동 입력, 결과란은 공란)
    doc.add_heading('2. 밸리데이션 결과 요약 (Result Summary)', level=1)
    doc.add_paragraph("각 항목별 판정 기준 및 결과는 다음과 같다.")
    
    t_res = doc.add_table(rows=1, cols=4); t_res.style = 'Table Grid'
    res_headers = ["Test Item", "Acceptance Criteria", "Result Summary", "Pass/Fail"]
    for i, h in enumerate(res_headers): 
        c = t_res.rows[0].cells[i]; c.text = h; set_table_header_style(c)
    
    # 항목별 기준 불러오기 & 행 추가
    items_map = [
        ("System Suitability", params.get('SST_Criteria', "RSD ≤ 2.0%")),
        ("Specificity", params.get('Detail_Specificity', "No Interference")),
        ("Linearity", params.get('Detail_Linearity', "R² ≥ 0.990")),
        ("Accuracy", params.get('Detail_Accuracy', "80 ~ 120%")),
        ("Precision", params.get('Detail_Precision', "RSD ≤ 2.0%")),
        ("LOD/LOQ", params.get('Detail_LOD', "S/N ≥ 3, 10"))
    ]
    
    for item, criteria in items_map:
        if criteria and "정보 없음" not in criteria:
            row = t_res.add_row().cells
            row[0].text = item
            row[1].text = criteria # 기준 자동 입력
            row[2].text = "" # 결과는 사용자가 엑셀 보고 입력하도록 비워둠
            row[3].text = "□ Pass  □ Fail"

    # 5. Detailed Results (상세 섹션 생성)
    doc.add_heading('3. 상세 결과 (Detailed Results)', level=1)
    doc.add_paragraph("※ 첨부된 엑셀 로우데이터(Raw Data) 및 크로마토그램 참조.")

    # 각 항목별 섹션 자동 생성
    for item, criteria in items_map:
        if criteria and "정보 없음" not in criteria:
            doc.add_heading(f"3.{items_map.index((item,criteria))+1} {item}", level=2)
            doc.add_paragraph(f"■ Acceptance Criteria: {criteria}")
            doc.add_paragraph("■ Result:")
            # 빈 표 삽입 (사용자가 엑셀 표 복붙하기 좋게)
            t_dummy = doc.add_table(rows=5, cols=3); t_dummy.style = 'Table Grid'
            t_dummy.rows[0].cells[0].text = "Parameter"
            t_dummy.rows[0].cells[1].text = "Value"
            t_dummy.rows[0].cells[2].text = "Note"
            set_table_header_style(t_dummy.rows[0].cells[0])
            doc.add_paragraph()

    # 6. Conclusion
    doc.add_heading('4. 종합 결론 (Conclusion)', level=1)
    doc.add_paragraph(f"상기 밸리데이션 수행 결과, '{method_name}' 시험법은 설정된 모든 판정 기준을 만족하였으므로 공정서 시험법으로서 적합함을 보증한다.")
    doc.add_paragraph("\n[ End of Document ]").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Excel 생성 함수 - Smart Logbook (ACTUAL WEIGHT & CORRECTION LOGIC)]
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # 1. 스타일 정의 (최상단 배치로 NameError 방지)
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

    # -----------------------------------------------------------
    # 1. Info Sheet (스크린샷 레이아웃 완벽 일치 작업)
    # -----------------------------------------------------------
    ws1 = workbook.add_worksheet("1. Info")
    ws1.set_column('A:A', 25); ws1.set_column('B:E', 15)
    ws1.merge_range('A1:E1', f'GMP Logbook: {method_name}', header)
    
    # Rows 3~6: Basic Info
    info_rows = [("Date", datetime.now().strftime("%Y-%m-%d")), ("Instrument", params.get('Instrument')), ("Column", params.get('Column_Plate')), ("Analyst", "")]
    for i, (k, v) in enumerate(info_rows):
        ws1.write(i+3, 0, k, sub)
        ws1.merge_range(i+3, 1, i+3, 4, v if v else "", cell)
    
    # Row 9 (Excel B9): Round Rule (텍스트)
    ws1.write(8, 0, "Round Rule:", sub)
    ws1.merge_range(8, 1, 8, 4, "모든 계산값은 소수점 2째자리(농도 3째자리)에서 절사(ROUNDDOWN).", cell)
    
    # Row 10 (Excel B10): Target Conc (숫자 값)
    target_conc_val = float(params.get('Target_Conc', 1.0))
    ws1.write(9, 0, "Target Conc (100%):", sub)
    ws1.write(9, 1, target_conc_val, calc) # B10
    ws1.write(9, 2, params.get('Unit', 'mg/mL'), cell)
    
    target_conc_ref = "'1. Info'!$B$10" 

    # Row 11: Header
    ws1.merge_range(10, 0, 10, 4, "■ Standard Preparation & Correction Factor", sub_rep)
    
    # Row 12~16: Inputs
    labels = ["Theoretical Stock (mg/mL):", "Purity (Potency, %):", "Water Content (%):", "Actual Weight (mg):", "Final Volume (mL):"]
    # B12, B13, B14, B15, B16
    for i, label in enumerate(labels):
        ws1.write(11 + i, 0, label, sub)
        if "Purity" in label: ws1.write(11 + i, 1, 100.0, calc)
        elif "Water" in label: ws1.write(11 + i, 1, 0.0, calc)
        else: ws1.write(11 + i, 1, "", calc)

    # Row 17 (Excel B17): Actual Stock
    ws1.write(16, 0, "Actual Stock (mg/mL):", sub)
    # 수식: (Weight(B15) * Purity(B13) * Water(B14)) / Vol(B16)
    ws1.write_formula(16, 1, '=IF(B16="","",ROUNDDOWN((B15*(B13/100)*((100-B14)/100))/B16, 4))', auto)
    actual_stock_ref = "'1. Info'!$B$17"

    # Row 18 (Excel B18): Correction Factor
    ws1.write(17, 0, "Correction Factor:", sub)
    ws1.write_formula(17, 1, '=IF(OR(B12="",B12=0,B17=""), 1, ROUNDDOWN(B17/B12, 4))', total_fmt)
    corr_factor_ref = "'1. Info'!$B$18"
    theo_stock_ref = "'1. Info'!$B$12" # [중요] B12 참조

    # -----------------------------------------------------------
    # 2. SST Sheet
    # -----------------------------------------------------------
    ws_sst = workbook.add_worksheet("2. SST"); ws_sst.set_column('A:F', 15)
    ws_sst.merge_range('A1:F1', 'System Suitability Test (n=6)', header)
    ws_sst.write_row('A2', ["Inj No.", "RT (min)", "Area", "Height", "Tailing (1st)", "Plate Count"], sub)
    for i in range(1, 7): ws_sst.write(i+1, 0, i, cell); ws_sst.write_row(i+1, 1, ["", "", "", "", ""], calc)
    ws_sst.write('A9', "Mean", sub); ws_sst.write_formula('B9', "=ROUNDDOWN(AVERAGE(B3:B8), 2)", auto); ws_sst.write_formula('C9', "=ROUNDDOWN(AVERAGE(C3:C8), 2)", auto)
    ws_sst.write('A10', "RSD(%)", sub); ws_sst.write_formula('B10', "=ROUNDDOWN(STDEV(B3:B8)/B9*100, 2)", auto); ws_sst.write_formula('C10', "=ROUNDDOWN(STDEV(C3:C8)/C9*100, 2)", auto)
    ws_sst.write('A12', "Criteria (RSD):", sub); ws_sst.write('B12', "≤ 2.0%", cell)
    ws_sst.write('C12', "Criteria (Tail):", sub); ws_sst.write('D12', "≤ 2.0 (Inj #1)", cell) 
    ws_sst.write('E12', "Result:", sub)
    ws_sst.write_formula('F12', '=IF(AND(B10<=2.0, C10<=2.0, E3<=2.0), "Pass", "Fail")', pass_fmt)
    ws_sst.conditional_format('F12', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
    
    # [Criteria 명시]
    ws_sst.write('A14', "※ Acceptance Criteria:", crit_fmt)
    ws_sst.write('A15', "1) RSD of RT & Area ≤ 2.0%")
    ws_sst.write('A16', "2) Tailing Factor (1st Inj) ≤ 2.0")

    # -----------------------------------------------------------
    # 3. Specificity Sheet
    # -----------------------------------------------------------
    ws_spec = workbook.add_worksheet("3. Specificity"); ws_spec.set_column('A:E', 20)
    ws_spec.merge_range('A1:E1', 'Specificity Test', header)
    ws_spec.write('A3', "Std Mean Area (Ref. SST):", sub); ws_spec.write_formula('B3', "='2. SST'!C9", num)
    ws_spec.write_row('A5', ["Sample", "RT", "Area", "Interference (%)", "Result"], sub)
    for i, s in enumerate(["Blank", "Placebo"]):
        row = i + 6
        ws_spec.write(row, 0, s, cell); ws_spec.write_row(row, 1, ["", ""], calc)
        ws_spec.write_formula(row, 3, f'=IF(OR(C{row+1}="", $B$3=""), "", ROUNDDOWN(C{row+1}/$B$3*100, 2))', auto)
        ws_spec.write_formula(row, 4, f'=IF(D{row+1}="", "", IF(D{row+1}<=0.5, "Pass", "Fail"))', pass_fmt)
        ws_spec.conditional_format(f'E{row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
    ws_spec.write(9, 0, "Criteria: Interference ≤ 0.5%", crit_fmt)

    # [Criteria 명시]
    ws_spec.write(9, 0, "※ Acceptance Criteria:", crit_fmt)
    ws_spec.write(10, 0, "1) Interference Peak Area ≤ 0.5% of Standard Area")

    # -----------------------------------------------------------
    # 4. Linearity Sheet (수식 참조 수정 완료)
    # -----------------------------------------------------------
    ws2 = workbook.add_worksheet("4. Linearity")
    ws2.set_column('A:I', 13)
    ws2.merge_range('A1:I1', 'Linearity Test', header)
    
    row = 3
    rep_rows = {1: [], 2: [], 3: []}
    
    for rep in range(1, 4):
        ws2.merge_range(row, 0, row, 8, f"■ Repetition {rep}", sub_rep); row += 1
        ws2.write_row(row, 0, ["Level", "Conc (X)", "Area (Y)", "Back Calc", "Accuracy (%)", "Check"], sub); row += 1
        data_start = row
        for level in [80, 90, 100, 110, 120]:
            ws2.write(row, 0, f"{level}%", cell)
            
            # [수정] Info 시트 B9(텍스트) -> B10(Target Conc)로 참조 주소 변경
            # Target(B10) / Theo Stock(B12) 비율을 적용하여 실제 농도 산출
            formula_x = f"=ROUNDDOWN({actual_stock_ref} * ({level}/100) * ('1. Info'!$B$10 / '1. Info'!$B$12), 3)"
            
            ws2.write_formula(row, 1, formula_x, num3)
            ws2.write(row, 2, "", calc)
            rep_rows[rep].append(row + 1)
            
            # Back Calc & Accuracy Formulas
            slope = f"C{data_start+7}"; intercept = f"C{data_start+8}"
            ws2.write_formula(row, 3, f'=IF(C{row+1}="", "", ROUNDDOWN((C{row+1}-{intercept})/{slope}, 3))', auto)
            ws2.write_formula(row, 4, f'=IF(C{row+1}="", "", ROUNDDOWN(D{row+1}/B{row+1}*100, 1))', auto)
            ws2.write(row, 5, "OK", cell)
            row += 1
        
        # Regression Logic
        ws2.write(row, 1, "Slope:", sub); ws2.write_formula(row, 2, f"=SLOPE(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
        ws2.write(row+1, 1, "Intercept:", sub); ws2.write_formula(row+1, 2, f"=INTERCEPT(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
        ws2.write(row+2, 1, "R²:", sub); ws2.write_formula(row+2, 2, f"=RSQ(C{data_start+1}:C{row}, B{data_start+1}:B{row})", auto)
        
        # Chart
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
        chart.add_series({'name': f'Rep {rep}', 'categories': f"='4. Linearity'!$B${data_start+1}:$B${row}", 'values': f"='4. Linearity'!$C${data_start+1}:$C${row}", 'trendline': {'type': 'linear', 'display_equation': True, 'display_r_squared': True}})
        chart.set_size({'width': 350, 'height': 220})
        ws2.insert_chart(f'G{data_start}', chart)
        row += 6

    # Summary Section
    ws2.merge_range(row, 0, row, 8, "■ Summary (Mean of 3 Reps) & Final Check", sub_rep); row += 1
    ws2.write_row(row, 0, ["Level", "Conc (X)", "Mean Area", "STDEV", "% RSD", "Criteria (RSD≤5%)"], sub); row += 1
    sum_start = row
    for i, level in enumerate([80, 90, 100, 110, 120]):
        r1 = rep_rows[1][i]; r2 = rep_rows[2][i]; r3 = rep_rows[3][i]
        ws2.write(row, 0, f"{level}%", cell); ws2.write_formula(row, 1, f"=B{r1}", num3)
        ws2.write_formula(row, 2, f"=ROUNDDOWN(AVERAGE(C{r1},C{r2},C{r3}), 2)", auto)
        ws2.write_formula(row, 3, f"=ROUNDDOWN(STDEV(C{r1},C{r2},C{r3}), 2)", auto)
        ws2.write_formula(row, 4, f"=IF(C{row+1}=0, 0, ROUNDDOWN(D{row+1}/C{row+1}*100, 2))", auto)
        ws2.write_formula(row, 5, f'=IF(C{row+1}=0, "", IF(E{row+1}<=5.0, "Pass", "Fail"))', pass_fmt)
        ws2.conditional_format(f'F{row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt}); row += 1
    
    row += 1
    ws2.write(row, 1, "Final Slope:", sub); ws2.write_formula(row, 2, f"=ROUNDDOWN(SLOPE(C{sum_start+1}:C{sum_start+5}, B{sum_start+1}:B{sum_start+5}), 4)", auto)
    ws2.write(row+1, 1, "Final Intercept:", sub); ws2.write_formula(row+1, 2, f"=ROUNDDOWN(INTERCEPT(C{sum_start+1}:C{sum_start+5}, B{sum_start+1}:B{sum_start+5}), 4)", auto)
    ws2.write(row+2, 1, "Final R²:", sub); ws2.write_formula(row+2, 2, f"=ROUNDDOWN(RSQ(C{sum_start+1}:C{sum_start+5}, B{sum_start+1}:B{sum_start+5}), 4)", auto)
    ws2.write(row+2, 4, "Criteria: R² ≥ 0.990", crit_fmt)
    ws2.write_formula(row+2, 5, f'=IF(C{row+3}=0, "", IF(C{row+3}>=0.990, "Pass", "Fail"))', pass_fmt)
    ws2.conditional_format(f'F{row+3}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})

    # [Criteria 명시]
    ws2.write(row+4, 0, "※ Acceptance Criteria:", crit_fmt)
    ws2.write(row+5, 0, "1) Coefficient of determination (R²) ≥ 0.990")
    ws2.write(row+6, 0, "2) %RSD of peak areas at each level ≤ 5.0%")

    # -----------------------------------------------------------
    # 5. Accuracy Sheet
    # -----------------------------------------------------------
    ws_acc = workbook.add_worksheet("5. Accuracy")
    ws_acc.set_column('A:G', 15)
    ws_acc.merge_range('A1:G1', 'Accuracy Test (Recovery)', header)
    
    # [수정] Linearity 시트의 Summary 결과 위치인 C51(Slope), C52(Intercept)로 주소 변경
    # 기존 C62, C63은 빈 셀이라 0.00이 나왔습니다.
    ws_acc.write('E3', "Slope:", sub)
    ws_acc.write_formula('F3', "='4. Linearity'!C51", auto) 
    
    ws_acc.write('E4', "Intercept:", sub)
    ws_acc.write_formula('F4', "='4. Linearity'!C52", auto)
    
    ws_acc.write('G3', "(From Linearity)", cell)
    acc_row = 6
    for level in [80, 100, 120]:
        ws_acc.merge_range(acc_row, 0, acc_row, 6, f"■ Level {level}% (3 Reps)", sub_rep); acc_row += 1
        ws_acc.write_row(acc_row, 0, ["Rep", "Theo Conc", "Area", "Calc Conc", "Recovery (%)", "Criteria", "Result"], sub); acc_row += 1
        start_r = acc_row
        for rep in range(1, 4):
            ws_acc.write(acc_row, 0, rep, cell)
            # Theo Conc 수식 (직선성과 동일하게 보정 반영)
            ws_acc.write_formula(acc_row, 1, f"=ROUNDDOWN({actual_stock_ref} * ({level}/100) * ({target_conc_ref} / {theo_stock_ref}), 3)", num3)
            ws_acc.write(acc_row, 2, "", calc)
            ws_acc.write_formula(acc_row, 3, f'=IF(C{acc_row+1}="","",ROUNDDOWN((C{acc_row+1}-$F$4)/$F$3, 3))', auto)
            ws_acc.write_formula(acc_row, 4, f'=IF(D{acc_row+1}="","",ROUNDDOWN(D{acc_row+1}/B{acc_row+1}*100, 1))', auto)
            ws_acc.write(acc_row, 5, "80~120%", cell)
            ws_acc.write_formula(acc_row, 6, f'=IF(E{acc_row+1}="","",IF(AND(E{acc_row+1}>=80, E{acc_row+1}<=120), "Pass", "Fail"))', pass_fmt)
            ws_acc.conditional_format(f'G{acc_row+1}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt}); acc_row += 1
        ws_acc.write(acc_row, 3, "Mean Rec(%):", sub)
        ws_acc.write_formula(acc_row, 4, f"=ROUNDDOWN(AVERAGE(E{start_r+1}:E{acc_row}), 1)", total_fmt); acc_row += 2

    # [Criteria 명시]
    ws_acc.write(acc_row, 0, "※ Acceptance Criteria:", crit_fmt)
    ws_acc.write(acc_row+1, 0, "1) Individual & Mean Recovery: 80.0 ~ 120.0%") 

    # -----------------------------------------------------------
    # [Sheet 6] 6. Precision
    # -----------------------------------------------------------
    ws3 = workbook.add_worksheet("6. Precision")
    ws3.set_column('A:E', 15)
    ws3.merge_range('A1:E1', 'Precision', header)
    ws3.merge_range('A3:E3', "■ Day 1 (Repeatability)", sub)
    ws3.write_row('A4', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
    for i in range(6):
        ws3.write_row(4+i, 0, [i+1, "Sample", ""], calc)
    ws3.write_formula('D5', "=ROUNDDOWN(AVERAGE(C5:C10), 2)", num)
    ws3.write_formula('E5', "=ROUNDDOWN(STDEV(C5:C10)/D5*100, 2)", num)
    ws3.write('D11', "Result:", sub)
    ws3.write_formula('E11', '=IF(E5=0,"",IF(E5<=2.0,"Pass","Fail"))', pass_fmt)
    
    ws3.merge_range('A14:E14', "■ Day 2 (Intermediate Precision)", sub)
    ws3.write_row('A15', ["Inj", "Sample", "Result", "Mean", "RSD"], sub)
    for i in range(6):
        ws3.write_row(15+i, 0, [i+1, "Sample", ""], calc)
    ws3.write_formula('D16', "=ROUNDDOWN(AVERAGE(C16:C21), 2)", num)
    ws3.write_formula('E16', "=ROUNDDOWN(STDEV(C16:C21)/D16*100, 2)", num)
    
    ws3.write(23, 0, "※ Acceptance Criteria: RSD ≤ 2.0%", crit_fmt)

    # -----------------------------------------------------------
    # [Sheet 7] 7. Robustness
    # -----------------------------------------------------------
    ws4 = workbook.add_worksheet("7. Robustness")
    ws4.set_column('A:F', 20)
    ws4.merge_range('A1:F1', 'Robustness Conditions', header)
    ws4.write_row('A3', ["Condition", "Set", "Actual", "SST Result (RSD)", "Pass/Fail", "Note"], sub)
    for r, c in enumerate(["Standard", "Flow -0.1", "Flow +0.1", "Temp -2", "Temp +2"]): 
        ws4.write(4+r, 0, c, cell); ws4.write_row(4+r, 1, ["", "", ""], calc)
        ws4.write_formula(4+r, 4, f'=IF(D{5+r}="", "", IF(D{5+r}<=2.0, "Pass", "Fail"))', pass_fmt)
        ws4.conditional_format(f'E{5+r}', {'type': 'cell', 'criteria': '==', 'value': '"Fail"', 'format': fail_fmt})
    
    ws4.write(10, 0, "※ Acceptance Criteria: SST Criteria must be met (RSD ≤ 2.0%)", crit_fmt)

    # -----------------------------------------------------------
    # [Sheet 8] 8. LOD_LOQ
    # -----------------------------------------------------------
    ws_ll = workbook.add_worksheet("8. LOD_LOQ")
    ws_ll.set_column('A:E', 15)
    ws_ll.merge_range('A1:E1', 'LOD / LOQ Determination', header)
    ws_ll.write_row('A2', ["Item", "Signal", "Noise", "S/N Ratio", "Result"], sub)
    ws_ll.write('A3', "LOD Sample", cell); ws_ll.write('B3', "", calc); ws_ll.write('C3', "", calc)
    ws_ll.write_formula('D3', '=IF(C3="","",ROUNDDOWN(B3/C3, 1))', auto)
    ws_ll.write_formula('E3', '=IF(D3="","",IF(D3>=3, "Pass", "Fail"))', pass_fmt)
    
    ws_ll.write('A4', "LOQ Sample", cell); ws_ll.write('B4', "", calc); ws_ll.write('C4', "", calc)
    ws_ll.write_formula('D4', '=IF(C4="","",ROUNDDOWN(B4/C4, 1))', auto)
    ws_ll.write_formula('E4', '=IF(D4="","",IF(D4>=10, "Pass", "Fail"))', pass_fmt)
    
    ws_ll.write(6, 0, "※ Acceptance Criteria:", crit_fmt)
    ws_ll.write(7, 0, "1) LOD: S/N Ratio ≥ 3")
    ws_ll.write(8, 0, "2) LOQ: S/N Ratio ≥ 10")

    workbook.close()
    output.seek(0)
    return output

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
                        st.info("👇 시료 상태와 농도를 입력하세요. (Target 농도가 100% 기준이 됩니다)")
                        sample_type = st.radio("시료 타입 (Sample Type):", ["Liquid (액체)", "Powder (파우더)"], horizontal=True)
                        cc1, cc2 = st.columns(2)
                        stock_input_val = 0.0; powder_desc = ""
                        if sample_type == "Liquid (액체)":
                            with cc1: stock_input_val = st.number_input("내 Stock 농도 (mg/mL 등):", min_value=0.0, step=0.1, format="%.2f")
                        else: 
                            with cc1: weight_input = st.number_input("칭량값 (Weight, mg):", min_value=0.0, step=0.1)
                            with cc2: dil_vol_input = st.number_input("희석 부피 (Vol, mL):", min_value=0.1, value=10.0, step=1.0)
                            if dil_vol_input > 0:
                                stock_input_val = weight_input / dil_vol_input
                                st.caption(f"🧪 계산된 Stock 농도: **{stock_input_val:.2f} mg/mL**")
                                powder_desc = f"Weigh {weight_input}mg / {dil_vol_input}mL"
                        params_p = get_method_params(sel_p); db_target = params_p.get('Target_Conc', 0.0)
                        with cc1: target_input_val = st.number_input("기준 농도 (Target 100%, mg/mL):", min_value=0.001, value=float(db_target) if db_target else 1.0, format="%.3f")
                        with cc2: vol_input = st.number_input("개별 바이알 조제 목표량 (Target Vol, mL):", min_value=1.0, value=5.0, step=1.0)
                        unit_val = params_p.get('Unit', '')
                        if stock_input_val > 0 and target_input_val > 0:
                            if stock_input_val < target_input_val * 1.2: st.error("⚠️ Stock 농도가 Target 농도(120% 범위)보다 낮습니다! 더 진한 Stock을 준비하세요.")
                            else:
                                calc_excel = generate_master_recipe_excel(sel_p, target_input_val, unit_val, stock_input_val, vol_input, sample_type, powder_desc)
                                st.download_button("🧮 시약 제조 계산기 (Master Recipe) 다운로드", calc_excel, f"Master_Recipe_{sel_p}.xlsx")
                                doc_proto = generate_protocol_premium(sel_p, "Cat", params_p, stock_input_val, vol_input, target_input_val)
                                st.download_button("📄 상세 계획서 (Protocol) 다운로드", doc_proto, f"Protocol_{sel_p}.docx", type="primary")

            with t2:
                st.markdown("### 📗 스마트 엑셀 일지 (Final Fixed)")
                st.info("✅ SST(Tailing Check), 특이성(Std 기준), 직선성(회차별 그래프), 정확성(자동 참조) 기능 탑재")
                sel_l = st.selectbox("Logbook:", my_plan["Method"].unique(), key="l")
                if st.button("Download Excel Logbook"):
                    data = generate_smart_excel(sel_l, "Cat", get_method_params(sel_l))
                    st.download_button("📊 Excel Logbook 다운로드", data, f"Logbook_{sel_l}.xlsx")

            with t3:
                st.markdown("### 📊 최종 결과 보고서 (Final Report)")
                st.info("💡 기기 정보, 분석 조건, 판정 기준이 포함된 **결과 보고서 초안(Draft)**을 생성합니다. 실험 완료 후 엑셀의 결과값을 붙여넣기 하세요.")
                
                sel_r = st.selectbox("Report for:", my_plan["Method"].unique(), key="r")
                
                # 입력 최소화: Lot 번호 정도만 입력 (선택 사항)
                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    lot_no = st.text_input("Lot No. (Optional):", value="TBD")
                
                # 버튼 클릭 시 즉시 다운로드
                if st.button("📥 최종 보고서 초안 다운로드 (Generate Report Docx)"):
                    # DB에서 파라미터 가져오기
                    param_data = get_method_params(sel_r)
                    cat_data = "Unknown Category" # 필요 시 로직 추가
                    
                    # 보고서 생성 함수 호출
                    doc_report = generate_summary_report_gmp(sel_r, cat_data, param_data, {'lot_no': lot_no})
                    
                    st.download_button(
                        label="📄 Word 파일 다운로드",
                        data=doc_report,
                        file_name=f"Validation_Report_{sel_r}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
