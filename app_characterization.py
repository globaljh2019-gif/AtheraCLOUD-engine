import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from datetime import datetime

# ==========================================
# 1. Notion Master Blueprint 기반 지식 베이스
# ==========================================
def get_notion_master_db(lang_code):
    """
    노션 라이브러리의 03_Analytical_Library 로직을 반영한 마스터 DB
    """
    if lang_code == "KR":
        return [
            {"Category": "1. 구조적 특성", "Attribute": "1차 구조 (아미노산 서열)", "Method": "Peptide Mapping (LC-MS/MS)", "Tier": "필수 (Tier 1)", "Rationale": "아미노산 서열 일치성 및 PTM 확인 필수", "Dev_Strategy": "Trypsin 소화 효율 최적화 및 Coverage 95% 이상 확보 전략."},
            {"Category": "1. 구조적 특성", "Attribute": "당쇄 프로파일 (N-Glycan)", "Method": "HILIC-FLD / MS", "Tier": "필수 (Tier 1)", "Rationale": "면역원성 및 이펙터 기능(ADCC) 영향 분석", "Dev_Strategy": "2-AB 라벨링 효율 및 주요 당쇄(G0F, G1F 등) 분리능 최적화."},
            {"Category": "2. 물리화학적 성질", "Attribute": "전하 변이체 (Charge Variants)", "Method": "CEX-HPLC / cIEF", "Tier": "필수 (Tier 1)", "Rationale": "단백질 안정성 및 불순물 프로파일 확인", "Dev_Strategy": "pH Gradient를 이용한 Acidic/Basic 변이체 분리능 극대화."},
            {"Category": "2. 물리화학적 성질", "Attribute": "크기 변이체 (응집체)", "Method": "SEC-HPLC", "Tier": "필수 (Tier 1)", "Rationale": "단백질 응집에 따른 안전성 위험 관리", "Dev_Strategy": "비특이적 결합 방지를 위한 이동상 염 농도 및 유속 최적화."},
            {"Category": "3. 생물학적 활성", "Attribute": "결합 역가 (Binding Affinity)", "Method": "SPR (Biacore) / ELISA", "Tier": "필수 (Tier 1)", "Rationale": "항원-항체 결합력(KD) 및 특이성 입증", "Dev_Strategy": "Chip 표면 고정화 농도 최적화 및 Kinetics 분석 정밀도 확보."},
        ]
    else:
        return [
            {"Category": "1. Structural", "Attribute": "Primary Structure", "Method": "Peptide Mapping (LC-MS/MS)", "Tier": "Tier 1", "Rationale": "Sequence confirmation and PTM site mapping", "Dev_Strategy": "Optimize digestion and target >95% sequence coverage."},
            {"Category": "1. Structural", "Attribute": "Glycan Profile (N-linked)", "Method": "HILIC-FLD / MS", "Tier": "Tier 1", "Rationale": "Impact on immunogenicity and ADCC activity", "Dev_Strategy": "Maximize labeling efficiency and resolve major glycoforms."},
            {"Category": "2. Physicochemical", "Attribute": "Charge Variants", "Method": "CEX-HPLC / cIEF", "Tier": "Tier 1", "Rationale": "Assessment of stability and variant profile", "Dev_Strategy": "Optimize pH gradient for acidic/basic peak resolution."},
            {"Category": "2. Physicochemical", "Attribute": "Size Variants (Aggregates)", "Method": "SEC-HPLC", "Tier": "Tier 1", "Rationale": "Safety risk management for protein aggregation", "Dev_Strategy": "Screen mobile phase salt concentration to prevent non-specific binding."},
            {"Category": "3. Biological", "Attribute": "Binding Affinity", "Method": "SPR (Biacore) / ELISA", "Tier": "Tier 1", "Rationale": "Demonstrate antigen-antibody binding (KD)", "Dev_Strategy": "Optimize ligand density and ensure kinetic data quality."},
        ]

# ==========================================
# 2. 문서 생성 엔진
# ==========================================
def set_cell_background(cell, color_hex):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    cell._element.get_or_add_tcPr().append(shd)

def generate_plan_report(product_name, phase, selected_df, lang):
    doc = Document()
    font_name = 'Malgun Gothic' if lang == "KR" else 'Arial'
    style = doc.styles['Normal']
    style.font.name = font_name
    if lang == "KR": style._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    
    title = "의약품 특성분석 종합 계획서" if lang == "KR" else "Comprehensive Characterization Plan"
    doc.add_heading(title, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading("1. 개요 (Project Overview)", level=1)
    doc.add_paragraph(f"제품명: {product_name} / 개발 단계: {phase}")

    doc.add_heading("2. 시험 항목 및 선정 근거 (Test Items & Rationale)", level=1)
    table = doc.add_table(rows=1, cols=4, style='Table Grid')
    headers = ["분류", "항목", "시험법", "선정근거"] if lang == "KR" else ["Category", "Attribute", "Method", "Rationale"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        set_cell_background(cell, 'E7E6E6')

    for _, row in selected_df.iterrows():
        cells = table.add_row().cells
        cells[0].text, cells[1].text, cells[2].text, cells[3].text = row['Category'], row['Attribute'], row['Method'], row['Rationale']

    doc.add_heading("3. 개발 전략 (Development Strategy)", level=1)
    for _, row in selected_df.iterrows():
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(f"{row['Method']}: ").bold = True
        p.add_run(row['Dev_Strategy'])

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# 3. 메인 UI
# ==========================================
def main():
    st.set_page_config(page_title="AtheraCLOUD - Characterization", layout="wide")
    
    with st.sidebar:
        st.title("🧬 AtheraCLOUD")
        lang = st.radio("Language Select / 언어 선택", ["Korean (국문)", "English (영문)"])
        lang_code = "KR" if "Korean" in lang else "EN"
        product_name = st.text_input("제품명 (Product Name)", "Athera-mAb-001")
        phase = st.selectbox("개발 단계 (Phase)", ["비임상", "임상 1상", "임상 3상", "BLA"])

    st.header(f"🧪 {lang_code} 특성분석 엔진 (Characterization Engine)")
    st.info("노션 마스터 블루프린트 로직 기반 종합 계획서 생성 시스템")

    # 원본 데이터 로드
    db_list = get_notion_master_db(lang_code)
    master_df = pd.DataFrame(db_list)
    
    # 탭 구성
    tab1, tab2, tab3 = st.tabs(["📋 종합계획서 (Summary Plan)", "🔬 시험항목 선정 (Decision)", "💡 개발 가이드 (Strategy)"])

    # [Step 1] 항목 선정 (Tab 2)
    with tab2:
        st.subheader("시험 항목 선정 (Method Decision)")
        # 체크박스 선택용 데이터프레임 생성
        display_df = master_df.copy()
        display_df.insert(0, '선택 (Select)', True)
        
        edited_df = st.data_editor(
            display_df[['선택 (Select)', 'Category', 'Attribute', 'Method', 'Rationale']], 
            use_container_width=True, 
            hide_index=True
        )
        
        # 사용자가 선택한 행의 'Attribute' 리스트 추출
        selected_attributes = edited_df[edited_df['선택 (Select)'] == True]['Attribute'].tolist()
        # 원본 데이터에서 선택된 행만 필터링 (에러 방지 핵심)
        selected_df = master_df[master_df['Attribute'].isin(selected_attributes)].copy()

    # [Step 2] 종합계획서 (Tab 1)
    with tab1:
        st.subheader("종합계획서 미리보기 (Master Plan Preview)")
        if not selected_df.empty:
            st.dataframe(selected_df[['Category', 'Attribute', 'Method']], use_container_width=True, hide_index=True)
            
            # 리포트 파일 생성
            doc_file = generate_plan_report(product_name, phase, selected_df, lang_code)
            
            st.success("종합 계획서 생성이 완료되었습니다.")
            st.download_button(
                label=f"📥 {lang_code} 종합계획서 다운로드 (.docx)",
                data=doc_file,
                file_name=f"Characterization_Plan_{lang_code}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.warning("시험항목 선정 탭에서 항목을 선택해주세요.")

    # [Step 3] 개발 가이드 (Tab 3)
    with tab3:
        st.subheader("상세 개발 가이드 (Development Guide)")
        if not selected_df.empty:
            for _, row in selected_df.iterrows():
                with st.expander(f"📌 {row['Attribute']} - {row['Method']}"):
                    st.success(f"Strategy: {row['Dev_Strategy']}")
        else:
            st.warning("항목을 선택하면 가이드가 표시됩니다.")

if __name__ == "__main__":
    main()