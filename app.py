import streamlit as st
import pandas as pd
import requests
import io
import xlsxwriter  # 엑셀 생성을 위한 필수 라이브러리
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------
# 1. 설정 및 보안 (API 키 로딩)
# ---------------------------------------------------------
try:
    NOTION_API_KEY = st.secrets["NOTION_API_KEY"]
    CRITERIA_DB_ID = st.secrets["CRITERIA_DB_ID"]
    STRATEGY_DB_ID = st.secrets["STRATEGY_DB_ID"]
    PARAM_DB_ID = st.secrets.get("PARAM_DB_ID", "") 
except:
    # 로컬 테스트용 (Secrets가 없을 경우 방어)
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
# 2. 노션 데이터 로딩 함수 (Backend)
# ---------------------------------------------------------
@st.cache_data
def get_criteria_map():
    """판정 기준 DB에서 카테고리별 필수 항목 매핑"""
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
    """전략 DB에서 시험 항목 리스트 추출"""
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
    """상세 파라미터 DB(8번)에서 시험법별 세부 정보 추출 (GMP 항목 포함)"""
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
                
                # 상세 밸리데이션 정보
                "Reference_Guideline": get_text("Reference_Guideline"),
                "Detail_Specificity": get_text("Detail_Specificity"),
                "Detail_Linearity": get_text("Detail_Linearity"),
                "Detail_Accuracy": get_text("Detail_Accuracy"),
                "Detail_Precision": get_text("Detail_Precision"),
                
                # GMP 일지 및 보고서용 정보
                "Reagent_List": get_text("Reagent_List"),
                "Ref_Standard_Info": get_text("Ref_Standard_Info"),
                "Preparation_Std": get_text("Preparation_Std"),
                "Preparation_Sample": get_text("Preparation_Sample"),
                "Calculation_Formula": get_text("Calculation_Formula"),
                "Logic_Statement": get_text("Logic_Statement"),
                
                # 엑셀 자동 계산용 숫자
                "Target_Conc": get_number("Target_Conc"),
                "Unit": get_text("Unit")
            }
    return None

# ---------------------------------------------------------
# 3. 문서 생성 엔진 (Word & Excel)
# ---------------------------------------------------------
def set_korean_font(doc):
    """한글 폰트 설정"""
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    style.font.size = Pt(10)

def generate_vmp_premium(modality, phase, df_strategy):
    """VMP 생성 (Word)"""
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
    """상세 계획서 생성 (Word)"""
    doc = Document()
    set_korean_font(doc)
    doc.add_heading(f'Validation Protocol: {method_name}', 0)
    doc.add_paragraph(f"Guideline: {params.get('Reference_Guideline', 'SOP')}")
    doc.add_heading('1. 기기 및 조건', level=1)
    doc.add_paragraph(f"기기: {params['Instrument']}\n컬럼: {params['Column_Plate']}\n조건: {params['Condition_A']} / {params['Condition_B']}")
    doc.add_heading('2. 밸리데이션 계획', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = "항목"; table.rows[0].cells[1].text = "절차 및 기준"
    
    items = [("특이성", params.get('Detail_Specificity')), ("직선성", params.get('Detail_Linearity')), 
             ("정확성", params.get('Detail_Accuracy')), ("정밀성", params.get('Detail_Precision'))]
    for k, v in items:
        if v:
            row = table.add_row().cells
            row[0].text = k; row[1].text = v
            
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

def generate_smart_excel(method_name, category, params):
    """스마트 엑셀 일지 생성 (Excel) - 수식 및 농도 자동 계산"""
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Logbook")

    # 스타일
    bold = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    cell_fmt = workbook.add_format({'border': 1})
    num_fmt = workbook.add_format({'border': 1, 'num_format': '0.00'})
    calc_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFFFCC', 'num_format': '0.00'})

    # 1. 헤더 정보
    worksheet.merge_range('A1:E1', f'GMP Analytical Logbook: {method_name}', bold)
    info_data = [("Method", method_name), ("Date", datetime.now().strftime("%Y-%m-%d")), 
                 ("Instrument", params.get('Instrument', '')), ("Column", params.get('Column_Plate', ''))]
    row = 2
    for k, v in info_data:
        worksheet.write(row, 0, k, bold)
        worksheet.merge_range(row, 1, row, 4, v, cell_fmt)
        row += 1
    
    row += 2
    # 2. 직선성 자동 농도 계산 (Target_Conc가 있을 경우)
    target_conc = params.get('Target_Conc')
    unit = params.get('Unit', 'ppm')
    
    if target_conc:
        worksheet.merge_range(row, 0, row, 4, f"■ 직선성 시험 (Linearity) - 기준 농도: {target_conc} {unit}", bold)
        row += 1
        headers = ["Level (%)", f"Target ({unit})", "실제 칭량값 (mg)", "희석 부피 (mL)", "실제 농도 (Calc)"]
        for col, h in enumerate(headers):
            worksheet.write(row, col, h, bold)
        
        row += 1
        levels = [80, 90, 100, 110, 120]
        for level in levels:
            target_val = float(target_conc) * (level / 100)
            worksheet.write(row, 0, f"{level}%", cell_fmt)
            worksheet.write(row, 1, target_val, num_fmt)
            worksheet.write(row, 2, "", cell_fmt) # 사용자 입력 (칭량)
            worksheet.write(row, 3, 50, cell_fmt) # 기본 부피
            
            # 엑셀 수식: (칭량 / 부피) * 1000 (단위 변환 가정)
            xl_row = row + 1
            formula = f"=C{xl_row}/D{xl_row}*1000"
            worksheet.write_formula(row, 4, formula, calc_fmt)
            row += 1
        worksheet.write(row+1, 0, "※ 노란색 셀은 값 입력 시 자동 계산됩니다.", cell_fmt)
        row += 3
    else:
        worksheet.merge_range(row, 0, row, 4, "⚠️ 노션에 'Target_Conc' 값이 없어 자동 계산 생략", cell_fmt)
        row += 3

    # 3. Raw Data
    worksheet.merge_range(row, 0, row, 4, "■ 데이터 기록 (Raw Data)", bold)
    row += 1
    headers = ["Inj No.", "Sample Name", "RT (min)", "Area", "Height"]
    for col, h in enumerate(headers):
        worksheet.write(row, col, h, bold)
    for _ in range(10): # 빈 칸 10줄
        row += 1
        for col in range(5):
            worksheet.write(row, col, "", cell_fmt)

    workbook.close()
    output.seek(0)
    return output

def generate_summary_report_gmp(method_name, category, params, user_inputs):
    """최종 보고서 생성 (Word) - 로직 포함"""
    doc = Document()
    set_korean_font(doc)
    doc.add_heading(f'Validation Summary Report: {method_name}', 0)
    
    # 1. 헤더
    info_table = doc.add_table(rows=3, cols=2)
    info_table.style = 'Table Grid'
    data = [("Test Category", category), ("Lot No / Date", f"{user_inputs['lot_no']} / {user_inputs['date']}"),
            ("Analyst", user_inputs['analyst'])]
    for i, (k, v) in enumerate(data):
        info_table.rows[i].cells[0].text = k
        info_table.rows[i].cells[1].text = str(v)

    # 2. SST
    doc.add_heading('1. 시스템 적합성 (System Suitability)', level=1)
    sst_table = doc.add_table(rows=2, cols=3)
    sst_table.style = 'Table Grid'
    sst_table.rows[0].cells[0].text = "기준"; sst_table.rows[0].cells[1].text = "결과"; sst_table.rows[0].cells[2].text = "판정"
    sst_table.rows[1].cells[0].text = params['SST_Criteria']
    sst_table.rows[1].cells[1].text = user_inputs['sst_result']
    sst_table.rows[1].cells[2].text = "Pass"

    # 3. 상세 결과 (로직 포함)
    doc.add_heading('2. 결과 산출 및 판정 (Calculation & Logic)', level=1)
    doc.add_paragraph(f"■ 계산식: {params.get('Calculation_Formula', 'SOP 참조')}")
    doc.add_paragraph(f"■ 판정 로직: {params.get('Logic_Statement', '기준 만족 시 적합')}")
    
    res_table = doc.add_table(rows=2, cols=2)
    res_table.style = 'Table Grid'
    res_table.rows[0].cells[0].text = "최종 결과값"; res_table.rows[0].cells[1].text = "판정 기준"
    res_table.rows[1].cells[0].text = user_inputs['main_result']
    res_table.rows[1].cells[1].text = params.get('Detail_Accuracy', 'SOP 참조')

    doc.add_heading('3. 결론 (Conclusion)', level=1)
    doc.add_paragraph("상기 결과는 설정된 기준을 만족하므로 적합(Pass)으로 판정함.")
    
    doc_io = io.BytesIO(); doc.save(doc_io); doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. 메인 UI (Streamlit App)
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD GMP Suite", layout="wide")
st.title("🧪 AtheraCLOUD: GMP Validation Suite")
st.markdown("##### Strategy · Protocol · Smart Excel Logbook · Report")

col1, col2 = st.columns([1, 3])
with col1:
    st.header("📂 Project")
    sel_modality = st.selectbox("Modality", ["mAb", "Cell Therapy", "Gene Therapy"])
    sel_phase = st.selectbox("Phase", ["Phase 1", "Phase 3"])

with col2:
    try:
        criteria_map = get_criteria_map()
        df_full = get_strategy_list(criteria_map)
    except:
        st.error("Notion 연결 실패. API Key와 DB ID를 확인하세요.")
        df_full = pd.DataFrame()

    if sel_modality == "mAb" and not df_full.empty:
        my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
        
        if not my_plan.empty:
            tab1, tab2, tab3 = st.tabs(["📑 Step 1: Protocol", "📗 Step 2: Excel Logbook", "📊 Step 3: Report"])
            
            # --- Tab 1: Protocol ---
            with tab1:
                st.subheader("전략 및 계획서 생성")
                st.dataframe(my_plan[["Method", "Category"]], use_container_width=True)
                doc_vmp = generate_vmp_premium(sel_modality, sel_phase, my_plan)
                st.download_button("📥 VMP 다운로드", doc_vmp, "VMP_Master.docx")
                
                st.divider()
                sel_proto = st.selectbox("상세 계획서 선택:", my_plan["Method"].unique())
                if sel_proto:
                    params = get_method_params(sel_proto)
                    if params:
                        doc_proto = generate_protocol_premium(sel_proto, "Category", params)
                        st.download_button(f"📄 {sel_proto} Protocol 다운로드", doc_proto, f"Protocol_{sel_proto}.docx")

            # --- Tab 2: Smart Excel Logbook ---
            with tab2:
                st.subheader("📗 스마트 엑셀 일지 (Smart Excel)")
                st.info("기준 농도(Target_Conc)에 맞춰 5포인트 직선성 농도와 수식이 자동 계산된 엑셀을 생성합니다.")
                sel_log = st.selectbox("일지 생성 시험법:", my_plan["Method"].unique(), key="log")
                
                params_log = get_method_params(sel_log)
                if params_log:
                    excel_data = generate_smart_excel(sel_log, "Category", params_log)
                    st.download_button(
                        label=f"📊 {sel_log} Excel 일지 다운로드",
                        data=excel_data,
                        file_name=f"Logbook_{sel_log}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

            # --- Tab 3: Report (Secure) ---
            with tab3:
                st.subheader("📊 최종 결과 보고서 (보안 모드)")
                sel_rep = st.selectbox("보고서 생성 시험법:", my_plan["Method"].unique(), key="rep")
                params_rep = get_method_params(sel_rep)
                
                if params_rep:
                    if "generated_doc" not in st.session_state:
                        st.session_state.generated_doc = None

                    with st.form("report_form"):
                        st.write(f"**[{sel_rep}] 결과 입력 (서버 저장 안됨)**")
                        c1, c2 = st.columns(2)
                        with c1:
                            input_lot = st.text_input("Lot No.")
                            input_date = st.date_input("시험일자")
                        with c2:
                            input_analyst = st.text_input("시험자")
                            input_sst = st.text_input("SST 결과")
                        input_main = st.text_input("최종 결과값")
                        
                        submitted = st.form_submit_button("🚀 보고서 생성")
                        
                        if submitted:
                            cat = my_plan[my_plan["Method"] == sel_rep].iloc[0]["Category"]
                            user_data = {"lot_no": input_lot, "date": input_date, "analyst": input_analyst,
                                         "sst_result": input_sst, "main_result": input_main}
                            st.session_state.generated_doc = generate_summary_report_gmp(sel_rep, cat, params_rep, user_data)
                    
                    if st.session_state.generated_doc:
                        st.success("보고서 생성 완료")
                        st.download_button("📥 보고서 다운로드", st.session_state.generated_doc, f"Report_{sel_rep}.docx")