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

# [NEW] 시약 제조 계산기 엑셀 생성 함수 (별도 다운로드용)
def generate_calculator_excel(method_name, target_conc, unit, stock_conc, req_vol):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Formats
    header = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#4472C4', 'font_color':'white', 'align':'center', 'valign':'vcenter'})
    sub = workbook.add_format({'bold':True, 'border':1, 'bg_color':'#D9E1F2', 'align':'center'})
    cell = workbook.add_format({'border':1, 'align':'center'}); num = workbook.add_format({'border':1, 'num_format':'0.00', 'align':'center'})
    auto = workbook.add_format({'border':1, 'bg_color':'#E2EFDA', 'num_format':'0.000', 'align':'center'}) # Green for Result

    ws = workbook.add_worksheet("Calculator")
    ws.set_column('A:E', 15)
    
    ws.merge_range('A1:E1', f'Solution Preparation Calculator: {method_name}', header)
    
    # User Inputs (Display)
    ws.write('A3', "Stock Conc:", sub); ws.write('B3', stock_conc, num); ws.write('C3', unit, cell)
    ws.write('A4', "Req Vol (mL):", sub); ws.write('B4', req_vol, num)
    ws.write('A5', "Target (100%):", sub); ws.write('B5', target_conc, num); ws.write('C5', unit, cell)
    
    # Recipe Table
    ws.write_row('A7', ["Level (%)", "Target Conc", "Stock Vol (mL)", "Diluent Vol (mL)", "Check Vol"], sub)
    
    levels = [80, 90, 100, 110, 120]
    for i, level in enumerate(levels):
        r = 8 + i
        t_val = float(target_conc) * (level / 100)
        s_vol = (t_val * float(req_vol)) / float(stock_conc)
        d_vol = float(req_vol) - s_vol
        
        ws.write(r, 0, f"{level}%", cell)
        ws.write(r, 1, t_val, num)
        ws.write(r, 2, s_vol, auto) # Calculated Stock
        ws.write(r, 3, d_vol, auto) # Calculated Diluent
        ws.write(r, 4, s_vol + d_vol, num)

    ws.merge_range('A15:E15', "※ 이 시트를 인쇄하여 실험 노트에 부착하십시오.", cell)
    
    workbook.close(); output.seek(0)
    return output

# [PROTOCOL 업그레이드: 계산된 레시피를 표로 삽입]
def generate_protocol_premium(method_name, category, params, stock_conc=None, req_vol=None):
    doc = Document(); set_korean_font(doc)
    
    def safe_get(key, default=""):
        val = params.get(key); return str(val) if val is not None else default

    # 머리글
    section = doc.sections[0]; header = section.header
    htable = header.add_table(1, 2, Inches(6.0)) 
    ht_c1 = htable.cell(0, 0); p1 = ht_c1.paragraphs[0]
    p1.add_run(f"Protocol No.: VP-{method_name[:3]}-001\n").bold = True; p1.add_run(f"Test Category: {category}")
    ht_c2 = htable.cell(0, 1); p2 = ht_c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"Guideline: {safe_get('Reference_Guideline', 'ICH Q2(R2)')}\n").bold = True; p2.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d')}")

    # 제목
    title = doc.add_heading(f'밸리데이션 상세 계획서 (Validation Protocol)', 0); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Method Name: {method_name}").alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph()

    # 1~3 섹션 (기존 동일)
    doc.add_heading('1. 목적 (Objective)', level=1)
    doc.add_paragraph(f"본 문서는 '{method_name}' 시험법의 밸리데이션 수행 방법 및 판정 기준을 기술한다.")
    doc.add_heading('2. 근거 (Reference)', level=1)
    doc.add_paragraph("• ICH Q2(R2) & MFDS 가이드라인")
    doc.add_heading('3. 기기 및 시약', level=1)
    t_cond = doc.add_table(rows=0, cols=2); t_cond.style = 'Table Grid'
    for k, v in [("기기", safe_get('Instrument')), ("컬럼", safe_get('Column_Plate')), 
                 ("조건", f"A: {safe_get('Condition_A')}\nB: {safe_get('Condition_B')}"), ("검출기", safe_get('Detection'))]:
        r = t_cond.add_row().cells; r[0].text=k; r[0].paragraphs[0].runs[0].bold=True; r[1].text=v
    
    # 4. 판정 기준 (표에는 기준만)
    doc.add_heading('4. 밸리데이션 항목 및 기준 (Criteria)', level=1)
    table = doc.add_table(rows=1, cols=2); table.style = 'Table Grid'
    headers = ["항목 (Parameter)", "판정 기준 (Criteria)"]
    for i, h in enumerate(headers): c = table.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    items = [("특이성", safe_get('Detail_Specificity')), ("직선성", safe_get('Detail_Linearity')), ("범위", safe_get('Detail_Range')),
             ("정확성", safe_get('Detail_Accuracy')), ("정밀성", safe_get('Detail_Precision')), ("완건성", safe_get('Detail_Robustness'))]
    for k, v in items:
        if v and "정보 없음" not in v: r = table.add_row().cells; r[0].text=k; r[1].text=v

    # 5. 상세 시험 방법 (Actionable SOP Style)
    doc.add_heading('5. 상세 시험 방법 (Test Procedures)', level=1)
    
    target_conc = safe_get('Target_Conc', '100'); unit = safe_get('Unit', '%')
    
    doc.add_heading('5.1 용액 조제 (Preparation)', level=2)
    doc.add_paragraph(f"1) 표준 모액(Stock Solution): 농도 {stock_conc if stock_conc else '[입력필요]'} {unit} 용액을 준비한다.")
    
    # [핵심] 사용자가 입력값을 넣었으면, 계산된 표를 문서에 박아버림
    if stock_conc and req_vol:
        doc.add_paragraph("2) 직선성 표준액 조제표 (Dilution Scheme): 아래 표에 따라 희석한다.")
        t_recipe = doc.add_table(rows=1, cols=4); t_recipe.style = 'Table Grid'
        rh = ["Level (%)", "Target Conc", "Stock Vol (mL)", "Diluent Vol (mL)"]
        for i, h in enumerate(rh): c = t_recipe.rows[0].cells[i]; c.text=h; set_table_header_style(c)
        
        levels = [80, 90, 100, 110, 120]
        for level in levels:
            t_val = float(target_conc) * (level/100)
            s_vol = (t_val * float(req_vol)) / float(stock_conc)
            d_vol = float(req_vol) - s_vol
            r = t_recipe.add_row().cells
            r[0].text = f"{level}%"
            r[1].text = f"{t_val:.2f}"
            r[2].text = f"{s_vol:.3f}" # 소수점 3자리까지 상세하게
            r[3].text = f"{d_vol:.3f}"
    else:
        doc.add_paragraph("2) 직선성 표준액 조제: 기준 농도의 80 ~ 120% 범위 내에서 5개 농도를 단계적으로 희석하여 조제한다.")

    doc.add_heading('5.2 직선성 시험 (Linearity)', level=2)
    doc.add_paragraph("1) 위에서 조제한 5개 농도의 표준액을 준비한다.")
    doc.add_paragraph("2) 각 농도별로 3회 반복 주입(Triplicate Injection)한다.")
    doc.add_paragraph("3) 크로마토그램의 면적을 구하여 검량선을 작성하고 R²를 평가한다.")

    doc.add_heading('5.3 정밀성 및 정확성', level=2)
    doc.add_paragraph(f"1) 정밀성: 기준 농도({target_conc} {unit})를 6회 반복 조제하여 분석한다.")
    doc.add_paragraph("2) 정확성: 80%, 100%, 120% 농도로 각 3회씩 조제하여 회수율을 구한다.")

    doc.add_paragraph("\n\n")
    table_sign = doc.add_table(rows=2, cols=3); table_sign.style = 'Table Grid'
    for i, h in enumerate(["작성", "검토", "승인"]): c = table_sign.rows[0].cells[i]; c.text=h; set_table_header_style(c)
    for i in range(3): table_sign.rows[1].cells[i].text="\n(서명/날짜)\n"

    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# [Excel 생성 함수 - Logbook 전용]
def generate_smart_excel(method_name, category, params):
    output = io.BytesIO(); workbook = xlsxwriter.Workbook(output, {'in_memory': True})
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
    if target_conc:
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
                    st.markdown("#### 🧪 상세 계획서 생성 설정")
                    sel_p = st.selectbox("Protocol:", my_plan["Method"].unique())
                    
                    if sel_p:
                        # [NEW] 입력란: Stock 농도와 필요량 입력
                        st.info("👇 여기에 시약 정보를 입력하면 [계산기]와 [계획서]에 자동 반영됩니다.")
                        cc1, cc2 = st.columns(2)
                        with cc1: stock_input = st.number_input("내 Stock 농도 (mg/mL 등):", min_value=0.0, step=0.1, format="%.2f")
                        with cc2: vol_input = st.number_input("필요한 양 (mL):", min_value=1.0, value=5.0, step=1.0)
                        
                        params_p = get_method_params(sel_p)
                        target_conc_val = params_p.get('Target_Conc', 0)
                        unit_val = params_p.get('Unit', '')

                        # 다운로드 버튼 1: 계산기 엑셀
                        if stock_input > 0:
                            calc_excel = generate_calculator_excel(sel_p, target_conc_val, unit_val, stock_input, vol_input)
                            st.download_button("🧮 시약 제조 계산기 (Excel) 다운로드", calc_excel, f"Calculator_{sel_p}.xlsx")
                        
                        # 다운로드 버튼 2: 상세 계획서 (워드)
                        doc_proto = generate_protocol_premium(sel_p, "Cat", params_p, stock_input if stock_input > 0 else None, vol_input)
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