import streamlit as st
import google.generativeai as genai
import os
from datetime import date
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

# ==========================================
# 1. Page & API Configuration
# ==========================================
st.set_page_config(
    page_title="AI Tutor Feedback System",
    layout="centered",
    page_icon="🎓"
)

try:
    os.environ["GOOGLE_API_VERSION"] = "v1"
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=api_key)
    else:
        st.error("❌ API Key not found in secrets.toml")
except Exception as e:
    st.error(f"❌ Configuration Error: {e}")

# ==========================================
# 2. Dynamic Model Loading
# ==========================================
@st.cache_resource
def get_available_model():
    try:
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        priority_list = ['models/gemini-1.5-flash', 'models/gemini-1.5-pro', 'gemini-1.5-flash', 'gemini-1.5-pro']
        for target in priority_list:
            if target in available_models:
                return genai.GenerativeModel(target)
        if available_models:
            return genai.GenerativeModel(available_models[0])
    except Exception as e:
        st.error(f"Model detection failed: {e}")
    return None

model = get_available_model()

# ==========================================
# 3. Word Document Generation
# ==========================================
NAVY   = RGBColor(0x1C, 0x2B, 0x4B)
ORANGE = RGBColor(0xC8, 0x5A, 0x1A)
GREY   = RGBColor(0x99, 0x99, 0x99)

def _set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def create_word_doc(text, student_name, teacher_name):
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # --- Branding header ---
    hdr_table = doc.add_table(rows=1, cols=2)
    hdr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    left_cell = hdr_table.cell(0, 0)
    lp = left_cell.paragraphs[0]
    lr = lp.add_run("★  StarsEdu Online Tuition")
    lr.bold = True
    lr.font.size = Pt(12)
    lr.font.color.rgb = ORANGE
    lp.alignment = WD_ALIGN_PARAGRAPH.LEFT

    right_cell = hdr_table.cell(0, 1)
    rp = right_cell.paragraphs[0]
    rr = rp.add_run("PLACE\nLOGO")
    rr.font.size = Pt(9)
    rr.font.color.rgb = ORANGE
    rp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph()

    # --- Title ---
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title_para.add_run("Homework Feedback")
    tr.bold = True
    tr.font.size = Pt(24)
    tr.font.color.rgb = NAVY

    rule_para = doc.add_paragraph()
    rule_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr2 = rule_para.add_run("─────")
    rr2.font.color.rgb = ORANGE
    rr2.font.size = Pt(14)

    doc.add_paragraph()

    # --- Info band ---
    today_str = date.today().strftime("%B %d, %Y")
    info_table = doc.add_table(rows=2, cols=3)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["STUDENT NAME", "ASSIGNED TUTOR", "DATE"]
    values = [student_name, teacher_name, today_str]
    for col_idx, (lbl, val) in enumerate(zip(labels, values)):
        lc = info_table.cell(0, col_idx)
        vc = info_table.cell(1, col_idx)
        _set_cell_bg(lc, "F5F0E8")
        _set_cell_bg(vc, "F5F0E8")
        lp2 = lc.paragraphs[0]
        lp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lr2 = lp2.add_run(lbl)
        lr2.font.size = Pt(7.5)
        lr2.font.color.rgb = GREY
        vp = vc.paragraphs[0]
        vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        vr2 = vp.add_run(val)
        vr2.bold = True
        vr2.font.size = Pt(12)
        vr2.font.color.rgb = NAVY

    doc.add_paragraph()

    # --- Body: render AI text with basic markdown support ---
    for line in text.strip().split("\n"):
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph()
            continue

        is_bullet = stripped.startswith(("• ", "- ", "* "))
        is_heading = stripped.startswith("#")

        if is_heading:
            heading_text = re.sub(r"^#+\s*", "", stripped)
            para = doc.add_paragraph()
            pPr = para._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "4")
            bottom.set(qn("w:color"), "CCCCCC")
            pBdr.append(bottom)
            pPr.append(pBdr)
            run = para.add_run(heading_text)
            run.bold = True
            run.font.size = Pt(13)
            run.font.color.rgb = NAVY
            continue

        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if is_bullet:
            para.style = doc.styles["List Bullet"]
            stripped = re.sub(r"^[•\-\*]\s*", "", stripped)

        parts = re.split(r"(\*\*[^*]+\*\*)", stripped)
        for part in parts:
            if part.startswith("**") and part.endswith("**"):
                r = para.add_run(part[2:-2])
                r.bold = True
                r.font.size = Pt(11)
            else:
                r = para.add_run(part)
                r.font.size = Pt(11)

    # --- Footer ---
    doc.add_paragraph()
    footer_para = doc.add_paragraph()
    pPr = footer_para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "4")
    top.set(qn("w:color"), "CCCCCC")
    pBdr.append(top)
    pPr.append(pBdr)
    fr = footer_para.add_run(f"STARS_EDU // Academic Reporting System")
    fr.font.size = Pt(8)
    fr.font.color.rgb = GREY

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ==========================================
# 4. Main Application Interface
# ==========================================
st.title("🎓 Professional Feedback Portal")
st.write("Generate high-quality Word reports for your students instantly.")

with st.sidebar:
    st.header("📋 Report Details")
    teacher_name = st.text_input("Teacher Name")
    student_name = st.text_input("Student Name")

    st.write("---")
    st.header("📂 Upload Work")
    homework_file    = st.file_uploader("Upload Student Work (PDF/Image/Word)",  type=["pdf", "png", "jpg", "jpeg", "docx", "doc"])
    mark_scheme_file = st.file_uploader("Upload Mark Scheme (Optional)",          type=["pdf", "png", "jpg", "jpeg", "docx", "doc"])


def build_payload_part(uploaded_file, label):
    filename = uploaded_file.name.lower()
    if filename.endswith(".docx") or filename.endswith(".doc"):
        doc = Document(BytesIO(uploaded_file.getvalue()))
        lines = []
        for para in doc.paragraphs:
            if para.text.strip():
                lines.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                lines.append("\t".join(cell.text for cell in row.cells))
        return f"\n[{label}]\n" + "\n".join(lines)
    else:
        return {"mime_type": uploaded_file.type, "data": uploaded_file.getvalue()}


if homework_file:
    st.info(f"File '{homework_file.name}' ready for analysis.")

    if st.button("🚀 Process & Generate Word Report", use_container_width=True):
        if not teacher_name or not student_name:
            st.warning("⚠️ Please enter both Teacher and Student names in the sidebar.")
        elif not model:
            st.error("AI model is not responding. Check your API key.")
        else:
            with st.spinner(f"Analysing {student_name}'s work..."):
                prompt = f"""
                You are a senior international tutor named {teacher_name}.
                Write a personalized study report to your student, {student_name}.

                Context:
                - Use a professional yet encouraging first-person tone ("I observed", "You should focus on").
                - Today's date is {date.today().strftime("%B %d, %Y")}.

                Requirements:
                1. Greeting: Address {student_name} directly.
                2. Executive Summary: Brief overview of performance.
                3. Detailed Marking: List questions with [Correct] or [Incorrect]. Highlight conceptual errors (e.g., Ne vs Ni symbols).
                4. Key Gaps: Specific knowledge areas needing attention.
                5. Next Steps: Foundation (to consolidate) and Extension (to challenge).

                Language: English only.
                """

                payload = [prompt, build_payload_part(homework_file, "Student Work")]
                if mark_scheme_file:
                    payload.append(build_payload_part(mark_scheme_file, "Mark Scheme"))

                try:
                    response = model.generate_content(payload)
                    report_content = response.text

                    word_bytes = create_word_doc(report_content, student_name, teacher_name)

                    st.success(f"🎉 Report for {student_name} is ready!")
                    st.download_button(
                        label="📥 Download Word Report",
                        data=word_bytes,
                        file_name=f"{student_name}_Feedback_{date.today()}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error during report generation: {e}")
else:
    st.info("Upload a student's homework to begin.")

st.markdown("---")
st.caption(f"System Date: {date.today().strftime('%Y-%m-%d')} | Powered by Gemini 1.5 Pro")
