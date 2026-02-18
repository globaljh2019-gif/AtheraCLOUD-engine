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
# 0. Notion Master Blueprint 기반 지식 베이스
# ==========================================
def get_notion_master_db(lang_code):
    if lang_code == "KR":
        return [
            {"Category": "1. 구조적 특성", "Attribute": "1차 구조 (아미노산 서열)", "Method": "Peptide Mapping (LC-MS/MS)", "Tier": "Tier 1", "Rationale": "아미노산 서열 일치성 및 PTM 확인 필수", "Dev_Strategy": "Trypsin 소화 효율 최적화 및 Coverage 95% 이상 확보 전략."},
            {"Category": "1. 구조적 특성", "Attribute": "당쇄 프로파일 (N-Glycan)", "Method": "HILIC-FLD / MS", "Tier": "Tier 1", "Rationale": "면역원성 및 이펙터 기능(ADCC) 영향 분석", "Dev_Strategy": "2-AB 라벨링 효율 및 주요 당쇄(G0F, G1F 등) 분리능 최적화."},
            {"Category": "2. 물리화학적 성질", "Attribute": "전하 변이체 (Charge Variants)", "Method": "CEX-HPLC / cIEF", "Tier": "Tier 1", "Rationale": "단백질 안정성 및 불순물 프로파일 확인", "Dev_Strategy": "pH Gradient를 이용한 Acidic/Basic 변이체 분리능 극대화."},
            {"Category": "2. 물리화학적 성질", "Attribute": "크기 변이체 (응집체)", "Method": "SEC-HPLC", "Tier": "Tier 1", "Rationale": "단백질 응집에 따른 안전성 위험 관리", "Dev_Strategy": "비특이적 결합 방지를 위한 이동상 염 농도 및 유속 최적화."},
            {"Category": "3. 생물학적 활성", "Attribute": "결합 역가 (Binding Affinity)", "Method": "SPR (Biacore) / ELISA", "Tier": "Tier 1", "Rationale": "항원-항체 결합력(KD) 및 특이성 입증", "Dev_Strategy": "Chip 표면 고정화 농도 최적화 및 Kinetics 분석 정밀도 확보."},
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
# 1. 지식 베이스 (Database - Dual Language)
# ==========================================
def get_method_database(modality, lang):
    """
    모달리티별 시험 항목 DB (국문/영문 스위칭)
    """
    if modality == "Monoclonal Antibody (mAb)":
        if lang == "KR":
            # [국문 데이터]
            data = [
                {"Category": "1. 구조적 특성", "Attribute": "1차 구조 (아미노산 서열)", "Method": "Peptide Mapping (LC-MS/MS)", "Tier": "필수 (Tier 1)", "Dev_Strategy": "Trypsin 소화 효율 최적화 (4시간 vs Overnight). Sequence Coverage 95% 이상 목표."},
                {"Category": "1. 구조적 특성", "Attribute": "고차 구조 (2차/3차)", "Method": "CD (Far/Near UV) & DSC", "Tier": "심화 (Tier 2)", "Dev_Strategy": "Buffer 간섭 최소화 및 Reference와의 스펙트럼 중첩성(Similarity) 비교."},
                {"Category": "1. 구조적 특성", "Attribute": "이황화 결합", "Method": "Non-reduced / Reduced Peptide Mapping", "Tier": "필수 (Tier 1)", "Dev_Strategy": "Free Thiol 측정 병행. Scrambled disulfide bond 유무 확인."},
                {"Category": "1. 구조적 특성", "Attribute": "당쇄 프로파일 (N-Glycan)", "Method": "HILIC-FLD / MS", "Tier": "필수 (Tier 1)", "Dev_Strategy": "주요 당쇄(G0F, G1F 등) 정량 및 면역원성 당쇄(Man5, G0) 모니터링."},
                {"Category": "2. 물리화학적 성질", "Attribute": "전하 변이체", "Method": "CEX-HPLC (Salt/pH Gradient)", "Tier": "필수 (Tier 1)", "Dev_Strategy": "Acidic/Basic peak 분리능 확보. 등전점(pI) 확인."},
                {"Category": "2. 물리화학적 성질", "Attribute": "크기 변이체 (응집체)", "Method": "SEC-HPLC", "Tier": "필수 (Tier 1)", "Dev_Strategy": "비특이적 결합 방지(염 농도 조절). HMW/Monomer 분리능 확인."},
                {"Category": "2. 물리화학적 성질", "Attribute": "크기 변이체 (분해물)", "Method": "CE-SDS (Non-reduced)", "Tier": "필수 (Tier 1)", "Dev_Strategy": "샘플 전처리 온도/시간 최적화로 인위적 분해 방지."},
                {"Category": "3. 생물학적 활성", "Attribute": "결합 활성 (Binding)", "Method": "ELISA / SPR", "Tier": "필수 (Tier 1)", "Dev_Strategy": "항원 코팅 농도 최적화 및 평행성(Parallelism) 입증."},
                {"Category": "3. 생물학적 활성", "Attribute": "작용 기전 역가 (Potency)", "Method": "Cell-based Assay", "Tier": "필수 (Tier 1)", "Dev_Strategy": "세포주 민감도 확인 및 4-PL 커브 피팅 적합성 평가."},
                {"Category": "4. 불순물", "Attribute": "공정 유래 불순물", "Method": "HCP ELISA & qPCR", "Tier": "필수 (Tier 1)", "Dev_Strategy": "공정 특이적 키트 선정 및 DNA 추출 효율 확인."},
            ]
        else:
            # [English Data]
            data = [
                {"Category": "1. Structure", "Attribute": "Primary Structure", "Method": "Peptide Mapping (LC-MS/MS)", "Tier": "Tier 1", "Dev_Strategy": "Optimize digestion efficiency (4h vs overnight). Target >95% sequence coverage."},
                {"Category": "1. Structure", "Attribute": "Higher Order Structure", "Method": "CD (Far/Near UV) & DSC", "Tier": "Tier 2", "Dev_Strategy": "Minimize buffer interference. Compare spectral similarity with reference standard."},
                {"Category": "1. Structure", "Attribute": "Disulfide Bond", "Method": "Non-reduced / Reduced Mapping", "Tier": "Tier 1", "Dev_Strategy": "Check free thiols (Ellman's). Confirm absence of scrambled bonds."},
                {"Category": "1. Structure", "Attribute": "Glycan Profile", "Method": "HILIC-FLD / MS", "Tier": "Tier 1", "Dev_Strategy": "Quantify major glycans (G0F, G1F) and monitor immunogenic species (Man5)."},
                {"Category": "2. Physicochemical", "Attribute": "Charge Variants", "Method": "CEX-HPLC", "Tier": "Tier 1", "Dev_Strategy": "Ensure resolution of Acidic/Basic peaks. Verify pI consistency."},
                {"Category": "2. Physicochemical", "Attribute": "Size Variants (Aggregates)", "Method": "SEC-HPLC", "Tier": "Tier 1", "Dev_Strategy": "Control salt conc. to prevent non-specific binding. Check resolution."},
                {"Category": "2. Physicochemical", "Attribute": "Size Variants (Fragments)", "Method": "CE-SDS (Non-reduced)", "Tier": "Tier 1", "Dev_Strategy": "Optimize sample prep temp/time to minimize artificial degradation."},
                {"Category": "3. Biological Activity", "Attribute": "Binding Activity", "Method": "ELISA / SPR", "Tier": "Tier 1", "Dev_Strategy": "Optimize coating concentration. Demonstrate parallelism."},
                {"Category": "3. Biological Activity", "Attribute": "Potency (MoA)", "Method": "Cell-based Assay", "Tier": "Tier 1", "Dev_Strategy": "Check cell line sensitivity. Evaluate 4-PL curve fit suitability."},
                {"Category": "4. Impurities", "Attribute": "Process Impurities", "Method": "HCP ELISA & qPCR", "Tier": "Tier 1", "Dev_Strategy": "Select process-specific kit. Verify DNA recovery efficiency."},
            ]
        return pd.DataFrame(data)
    else:
        return pd.DataFrame()

# ==========================================
# 2. 문서 생성 엔진 (Report Generator - Dual)
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
    headers = ["Category", "Attribute", "Method", "Rationale"]
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
# 3. 메인 UI (Streamlit - Dual)
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

    db = get_notion_master_db(lang_code)
    df = pd.DataFrame(db)
    df['Include'] = True
    
    # 탭 순서: 종합계획서가 가장 먼저 나오도록 배치
    tab1, tab2, tab3 = st.tabs(["📋 종합계획서 (Summary Plan)", "🔬 시험항목 선정 (Decision)", "💡 개발 가이드 (Strategy)"])

    # 로직 상 Decision 탭의 데이터를 먼저 정의해야 함
    with tab2:
        st.subheader("시험 항목 선정 (Method Decision)")
        edited_df = st.data_editor(df[['Include', 'Category', 'Attribute', 'Method', 'Rationale']], use_container_width=True, hide_index=True)
        selected_rows = edited_df[edited_df['Include'] == True]

    with tab1:
        st.subheader("종합계획서 미리보기 (Master Plan Preview)")
        if not selected_rows.empty:
            st.dataframe(selected_rows[['Category', 'Attribute', 'Method']], use_container_width=True, hide_index=True)
            
            # 리포트 생성
            final_df = pd.merge(selected_rows, df, on=['Category', 'Attribute', 'Method', 'Rationale'])
            doc = generate_plan_report(product_name, phase, final_df, lang_code)
            
            st.success("종합 계획서 생성이 완료되었습니다.")
            st.download_button(
                label=f"📥 {lang_code} 종합계획서 다운로드 (.docx)",
                data=doc,
                file_name=f"Characterization_Plan_{lang_code}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.warning("선택 탭에서 시험 항목을 하나 이상 선택해주세요.")

    with tab3:
        st.subheader("상세 개발 가이드 (Development Guide)")
        if not selected_rows.empty:
            final_df = pd.merge(selected_rows, df, on=['Category', 'Attribute', 'Method', 'Rationale'])
            for _, row in final_df.iterrows():
                with st.expander(f"📌 {row['Attribute']} - {row['Method']}"):
                    st.success(f"Strategy: {row['Dev_Strategy_y']}")
        else:
            st.warning("선택 탭에서 항목을 선택하면 가이드가 표시됩니다.")

if __name__ == "__main__":
    main()