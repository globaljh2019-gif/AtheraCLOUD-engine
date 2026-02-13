import streamlit as st
import pandas as pd
import requests
import io
from datetime import datetime

# ---------------------------------------------------------
# [필수] 워드 문서 생성 라이브러리 (배경색 에러 수정 완료)
# ---------------------------------------------------------
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement  # ★ 이 부분이 핵심 (에러 해결)

# ---------------------------------------------------------
# 1. 설정 (API 키와 DB ID를 여기에 입력하세요)
# ---------------------------------------------------------
NOTION_API_KEY = st.secrets["NOTION_API_KEY"]
CRITERIA_DB_ID = st.secrets["CRITERIA_DB_ID"]
STRATEGY_DB_ID = st.secrets["STRATEGY_DB_ID"]

headers = {
    "Authorization": "Bearer " + NOTION_API_KEY,
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28"
}

# ---------------------------------------------------------
# 2. 노션 데이터 로딩 함수
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

# ---------------------------------------------------------
# 3. 고품질 VMP 문서 생성 함수 (Premium Version)
# ---------------------------------------------------------
def generate_vmp_premium(modality, phase, df_strategy):
    doc = Document()
    
    # [0] 스타일 설정 (한글 폰트 깨짐 방지 & 가독성)
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    style.font.size = Pt(11)
    
    # [1] 표지 및 헤더
    header = doc.sections[0].header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.add_run(f"Document No.: VMP-{phase}-{modality}-001 (Ver. 1.0)")

    doc.add_paragraph("\n\n") 
    title = doc.add_heading('밸리데이션 종합 계획서\n(Validation Master Plan)', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("\n")
    
    # 프로젝트 정보 요약
    info_table = doc.add_table(rows=3, cols=2)
    info_table.style = 'Table Grid'
    info_table.rows[0].cells[0].text = "제품 명칭 (Product)"
    info_table.rows[0].cells[1].text = f"{modality} Candidate (TBD)"
    info_table.rows[1].cells[0].text = "개발 단계 (Phase)"
    info_table.rows[1].cells[1].text = phase
    info_table.rows[2].cells[0].text = "작성 일자 (Date)"
    info_table.rows[2].cells[1].text = datetime.now().strftime("%Y년 %m월 %d일")

    doc.add_page_break() 

    # [2] 본문 섹션 시작
    
    # 1. 개요
    doc.add_heading('1. 개요 (Introduction)', level=1)
    p1 = doc.add_paragraph()
    p1.add_run(f"본 밸리데이션 종합 계획서(VMP)는 '{modality}' 의약품의 '{phase}' 임상시험 승인(IND)을 목표로 한다. ").bold = True
    p1.add_run(
        "본 문서는 의약품의 품질 관리(Quality Control)에 사용되는 시험방법이 "
        "의도된 목적에 적합함을 입증하기 위한 밸리데이션 수행 전략, 범위, 절차 및 판정 기준을 기술한다."
    )

    # 2. 적용 범위
    doc.add_heading('2. 적용 범위 (Scope)', level=1)
    doc.add_paragraph(
        "본 계획서는 원료의약품(Drug Substance) 및 완제의약품(Drug Product)의 "
        "출하 시험(Release Test) 및 안정성 시험(Stability Test)에 적용되는 모든 분석법에 적용된다."
    )

    # 3. 관련 가이드라인
    doc.add_heading('3. 관련 가이드라인 (References)', level=1)
    doc.add_paragraph("본 밸리데이션은 다음의 최신 가이드라인을 준수하여 수행된다:", style='List Bullet')
    doc.add_paragraph("ICH Q2(R2): Validation of Analytical Procedures", style='List Bullet')
    doc.add_paragraph("ICH Q6B: Specifications for Biotechnological/Biological Products", style='List Bullet')
    doc.add_paragraph("MFDS(식약처) 의약품 등 시험방법 밸리데이션 가이드라인", style='List Bullet')

    # 4. 밸리데이션 수행 전략 (핵심 표)
    doc.add_heading('4. 밸리데이션 수행 전략 (Validation Strategy)', level=1)
    doc.add_paragraph(
        "각 시험법의 특성(확인, 순도, 정량 등)과 목적에 따라, "
        "ICH Q2(R2) 가이드라인에 근거한 필수 검증 항목(Validation Characteristics)을 다음과 같이 설정한다."
    )

    # --- [표 그리기] ---
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False
    
    # 헤더 설정
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '시험법 명칭 (Method)'
    hdr_cells[1].text = '시험 구분 (Category)'
    hdr_cells[2].text = '필수 검증 항목 (Parameters)'
    
    # 헤더 스타일 (배경색 에러 수정됨)
    for cell in hdr_cells:
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 배경색 설정 (안전한 방식)
        tcPr = cell._element.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'E7E6E6') # 옅은 회색
        tcPr.append(shd)

    # 너비 설정 (17cm 기준)
    widths = [Cm(4.5), Cm(4.0), Cm(8.5)]
    for i in range(3):
        table.columns[i].width = widths[i]
        hdr_cells[i].width = widths[i]

    # 데이터 입력
    for index, row in df_strategy.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Method'])
        row_cells[1].text = str(row['Category'])
        row_cells[2].text = ", ".join(row['Required_Items'])
        
        # 너비 재적용
        for i in range(3):
            row_cells[i].width = widths[i]
            row_cells[i].vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("\n")

    # 5. 판정 기준 및 절차
    doc.add_heading('5. 판정 기준 및 절차 (Criteria & Procedure)', level=1)
    doc.add_paragraph(
        "각 검증 항목에 대한 세부 판정 기준은 개별 밸리데이션 계획서(Validation Protocol)에 명시하며, "
        "일반적인 허용 기준은 다음과 같다."
    )
    doc.add_paragraph("특이성 (Specificity): 주성분과 불순물 간의 간섭이 없을 것", style='List Bullet')
    doc.add_paragraph("직선성 (Linearity): 결정계수(R²) ≥ 0.990", style='List Bullet')
    doc.add_paragraph("정밀성 (Precision): 반복성 및 실험실 내 정밀성 RSD ≤ 2.0%", style='List Bullet')
    doc.add_paragraph("정확성 (Accuracy): 회수율 98.0 ~ 102.0% 범위 내", style='List Bullet')

    # 6. 종합 결론
    doc.add_heading('6. 종합 결론 (Conclusion)', level=1)
    doc.add_paragraph(
        "본 계획서에 기술된 전략에 따라 수행된 밸리데이션 결과는 최종 보고서(Validation Report)로 문서화되며, "
        "이는 IND 신청 시 CTD Module 3.2.S.4.3의 근거 자료로서 시험법의 과학적 타당성을 입증하는 데 사용된다."
    )

    # 7. 승인 서명란
    doc.add_paragraph("\n\n")
    doc.add_paragraph("승인 (Approval)", style='Heading 2')
    
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.style = 'Table Grid'
    
    sig_hdr = sig_table.rows[0].cells
    sig_hdr[0].text = "작성 (Prepared by)"
    sig_hdr[1].text = "검토 (Reviewed by)"
    sig_hdr[2].text = "승인 (Approved by)"
    
    sig_body = sig_table.rows[1].cells
    sig_body[0].text = "\n\n(서명)\nDate: "
    sig_body[1].text = "\n\n(서명)\nDate: "
    sig_body[2].text = "\n\n(서명)\nDate: "

    # 저장
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ---------------------------------------------------------
# 4. Streamlit UI (메인 화면)
# ---------------------------------------------------------
st.set_page_config(page_title="AtheraCLOUD Engine", layout="wide")

st.title("🧪 AtheraCLOUD: Validation Master Plan")
st.markdown("##### The First Step to IND Filing: Generate Your Strategy")

st.sidebar.header("📂 Project Setup")
sel_modality = st.sidebar.selectbox("Modality", ["mAb", "Cell Therapy", "Gene Therapy", "Exosome"])
sel_phase = st.sidebar.selectbox("Phase", ["Phase 1", "Phase 3"])

# 데이터 로딩
try:
    criteria_map = get_criteria_map()
    df_full = get_strategy_list(criteria_map)
except Exception as e:
    st.error(f"System Error: {e}")
    df_full = pd.DataFrame()

# ---------------------------------------------------------
# 5. 로직 분기
# ---------------------------------------------------------

if sel_modality == "mAb":
    # 필터링
    if not df_full.empty:
        my_plan = df_full[(df_full["Modality"] == sel_modality) & (df_full["Phase"] == sel_phase)]
    else:
        my_plan = pd.DataFrame()

    if my_plan.empty:
        st.warning(f"⚠️ {sel_modality} {sel_phase}에 대한 전략 데이터가 노션에 아직 입력되지 않았습니다.")
        st.info("Validation_Strategy_DB에 데이터를 추가해주세요.")
    else:
        st.success(f"✅ **{sel_modality} {sel_phase}** 밸리데이션 전략이 수립되었습니다.")
        
        # 전략 미리보기 (표)
        st.dataframe(
            my_plan[["Method", "Category", "Required_Items"]],
            use_container_width=True,
            column_config={"Required_Items": st.column_config.ListColumn("필수 수행 항목")}
        )
        
        st.write("---")
        
        st.info("💡 아래 버튼을 누르면 전체 전략이 포함된 '밸리데이션 종합 계획서(VMP)'가 생성됩니다.")
            
        # 문서 생성 (여기가 에러 났던 부분 - 수정됨)
        doc_file = generate_vmp_premium(sel_modality, sel_phase, my_plan)
        
        st.download_button(
            label="📄 밸리데이션 종합 계획서(VMP) 다운로드",
            data=doc_file,
            file_name=f"VMP_{sel_modality}_{sel_phase}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )

else:
    st.info(f"🚧 **{sel_modality}** 모듈은 현재 개발 중입니다.")
    st.markdown("AtheraCLOUD 팀이 최신 가이드라인(FDA/EMA)을 반영하여 전략 엔진을 구축하고 있습니다.")