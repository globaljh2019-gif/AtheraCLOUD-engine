import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
from datetime import datetime

# ==========================================
# 1. Knowledge Base (ICH Q6B & Development Guide)
# ==========================================
def get_method_database(modality):
    """
    모달리티별 시험 항목 및 개발 가이드 DB
    """
    if modality == "Monoclonal Antibody (mAb)":
        data = [
            {
                "Category": "1. Structure", "Attribute": "Primary Structure", 
                "Method": "Peptide Mapping (LC-MS/MS)", "Tier": "Tier 1",
                "Dev_Strategy": "Optimization of digestion time (4h vs overnight) & Enzyme:Substrate ratio (1:20 vs 1:50). Target >95% coverage."
            },
            {
                "Category": "1. Structure", "Attribute": "Glycan Profile", 
                "Method": "HILIC-FLD / MS", "Tier": "Tier 1",
                "Dev_Strategy": "Fluorescent labeling efficiency check (2-AB vs RapiFluor). Column temp optimization (45-60°C) for sialylated species resolution."
            },
            {
                "Category": "2. Physicochemical", "Attribute": "Charge Variants", 
                "Method": "CEX-HPLC (Salt Gradient)", "Tier": "Tier 1",
                "Dev_Strategy": "Buffer pH screening (pH 5.5 - 7.0). Gradient slope optimization to separate acidic/basic variants from main peak."
            },
            {
                "Category": "2. Physicochemical", "Attribute": "Size Variants (Aggregates)", 
                "Method": "SEC-HPLC", "Tier": "Tier 1",
                "Dev_Strategy": "Mobile phase salt conc. (200-500mM) screening to minimize non-specific binding. Flow rate study for resolution."
            },
            {
                "Category": "2. Physicochemical", "Attribute": "Size Variants (Fragments)", 
                "Method": "CE-SDS (Non-reduced)", "Tier": "Tier 1",
                "Dev_Strategy": "Sample preparation temp/time (70°C 10min vs 3min) to prevent artificial fragmentation. Alkylation condition check."
            },
            {
                "Category": "3. Biological Activity", "Attribute": "Binding Activity", 
                "Method": "ELISA / SPR", "Tier": "Tier 1",
                "Dev_Strategy": "Plate coating concentration optimization. Specificity test against other mAbs and blocking buffers."
            },
             {
                "Category": "3. Biological Activity", "Attribute": "Potency (MoA)", 
                "Method": "Cell-based Assay", "Tier": "Tier 2",
                "Dev_Strategy": "Cell line sensitivity selection. Incubation time and cell density optimization. (Expect high variability, n=3 required)."
            },
        ]
        return pd.DataFrame(data)
    else:
        return pd.DataFrame() 

# ==========================================
# 2. Document Generator (Report Structure Updated)
# ==========================================
def generate_ind_report(product_name, modality, phase, selected_methods):
    doc = Document()
    
    # 스타일 설정
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)

    # 헤더
    doc.add_heading(f'Characterization Study Plan', 0)
    doc.add_paragraph(f"Product: {product_name} ({modality})")
    doc.add_paragraph(f"Target Phase: {phase}")
    doc.add_paragraph("-" * 70)

    # -------------------------------------------------------
    # 1. Comprehensive Plan (종합 계획서) - 가장 먼저 배치
    # -------------------------------------------------------
    doc.add_heading('1. Comprehensive Characterization Plan', level=1)
    doc.add_paragraph(f"The following test items have been established for the characterization of {product_name}.")

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    headers = ['Category', 'Quality Attribute', 'Test Method']
    
    # 테이블 헤더 스타일
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell._element.tcPr.append(qn('w:shd', {'w:fill': 'E7E6E6'}))

    # 테이블 내용 (Decision)
    for idx, row in selected_methods.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(row['Category'])
        cells[1].text = str(row['Attribute'])
        cells[2].text = str(row['Method'])

    # -------------------------------------------------------
    # 2. Method Decision Rationale (선정 근거)
    # -------------------------------------------------------
    doc.add_heading('2. Method Decision Rationale', level=1)
    doc.add_paragraph("The selection of characterization methods is based on ICH Q6B guidelines and the specific critical quality attributes (CQAs) of the molecule.")
    
    doc.add_paragraph("Rationale for Selection:", style='List Bullet')
    for idx, row in selected_methods.iterrows():
        p = doc.add_paragraph(style='List Bullet')
        runner = p.add_run(f"{row['Attribute']}: ")
        runner.bold = True
        p.add_run(f"Selected {row['Method']} as the primary method for {row['Category']} assessment (ICH Tier {row['Tier']}).")

    # -------------------------------------------------------
    # 3. Method Development Strategy (개발 전략)
    # -------------------------------------------------------
    doc.add_heading('3. Method Development Strategy', level=1)
    doc.add_paragraph("The following development strategies will be applied to optimize method performance:")
    
    for idx, row in selected_methods.iterrows():
        p = doc.add_paragraph()
        runner = p.add_run(f"[{row['Method']}] Development:")
        runner.bold = True
        doc.add_paragraph(f"   ► Strategy: {row['Dev_Strategy']}")
        doc.add_paragraph("") 

    # 저장
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# 3. UI Implementation
# ==========================================
def main():
    st.set_page_config(page_title="Characterization Engine", layout="wide")
    
    with st.sidebar:
        st.title("🧬 AtheraCLOUD")
        st.subheader("Project Info")
        
        modality = st.selectbox(
            "Modality", 
            ["Monoclonal Antibody (mAb)", "ADC (Coming Soon)", "Bispecific Ab (Coming Soon)"]
        )
        product_name = st.text_input("Product Name", "Athera-mAb-001")
        phase = st.selectbox("Phase", ["Pre-clinical", "Phase 1", "Phase 3", "BLA"])

    st.markdown(f"## 🧪 {modality} Characterization Engine")
    st.markdown("**Process Flow:** Plan Overview ➔ Method Decision ➔ Development Strategy")

    if "Coming Soon" in modality:
        st.warning(f"🚧 {modality} module is under development.")
        return

    # 데이터 로드
    df_db = get_method_database(modality)

    # 탭 순서 재배치 (Plan -> Decision -> Development)
    tab1, tab2, tab3 = st.tabs(["1️⃣ Comprehensive Plan (Output)", "2️⃣ Method Decision (Select)", "3️⃣ Method Development (Guide)"])

    # ------------------------------------------------------------------
    # 중요: Streamlit의 실행 순서상, 'Method Decision'(Tab2)의 입력값을
    # 'Comprehensive Plan'(Tab1)에서 보여주려면
    # 코드 상에서는 Tab 2 로직을 먼저 처리해야 합니다.
    # ------------------------------------------------------------------

    # --- [Logic for Tab 2] Method Decision (Selection) ---
    with tab2:
        st.subheader("Method Decision (Test Item Selection)")
        st.markdown("Select test items based on ICH Q6B CQAs.")
        
        df_db['Include'] = True 
        edited_df = st.data_editor(
            df_db[['Include', 'Category', 'Attribute', 'Method', 'Tier']],
            column_config={
                "Include": st.column_config.CheckboxColumn("Select", help="Include in Plan?"),
                "Tier": st.column_config.TextColumn("Tier", help="Tier 1: Mandatory"),
            },
            use_container_width=True,
            hide_index=True
        )
        selected_rows = edited_df[edited_df['Include'] == True]

    # --- [Logic for Tab 1] Comprehensive Plan (Output) ---
    with tab1:
        st.subheader("Comprehensive Characterization Plan")
        st.markdown("Based on your selection in Tab 2, here is the final plan.")
        
        if len(selected_rows) > 0:
            # 결과 미리보기 (깔끔한 테이블)
            st.dataframe(
                selected_rows[['Category', 'Attribute', 'Method']], 
                use_container_width=True,
                hide_index=True
            )
            
            # 리포트 생성 준비
            final_selection = pd.merge(selected_rows, df_db, on=['Category', 'Attribute', 'Method', 'Tier'], how='left')
            # merge시 중복 컬럼 처리
            if 'Dev_Strategy_y' in final_selection.columns:
                 final_selection['Dev_Strategy'] = final_selection['Dev_Strategy_y']

            doc_file = generate_ind_report(product_name, modality, phase, final_selection)
            
            st.success("The comprehensive plan is ready.")
            st.download_button(
                label="📄 Download Comprehensive Plan (.docx)",
                data=doc_file,
                file_name=f"{product_name}_Characterization_Plan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.warning("Please select at least one method in 'Method Decision' tab.")

    # --- [Logic for Tab 3] Method Development ---
    with tab3:
        st.subheader("Method Development Strategy")
        st.markdown("Technical guidelines for the selected methods.")
        
        if len(selected_rows) > 0:
            final_selection = pd.merge(selected_rows, df_db, on=['Category', 'Attribute', 'Method', 'Tier'], how='left')
            
            for index, row in final_selection.iterrows():
                strategy_text = row.get('Dev_Strategy_y', row.get('Dev_Strategy', ''))
                with st.expander(f"📌 {row['Attribute']} - {row['Method']}"):
                    st.write(f"**Tier:** {row['Tier']}")
                    st.info(f"**Optimization Strategy:**\n\n{strategy_text}")
        else:
            st.info("Select methods in Tab 2 to see development strategies.")

if __name__ == "__main__":
    main()