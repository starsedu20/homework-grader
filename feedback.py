import streamlit as st
import google.generativeai as genai
import os
from datetime import date
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
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
# 3. Word Document Helpers
# ==========================================
NAVY   = RGBColor(0x1C, 0x2B, 0x4B)
ORANGE = RGBColor(0xC8, 0x5A, 0x1A)
BEIGE  = "F5F0E8"
WHITE  = "FFFFFF"
DARK   = "1C2B4B"
LIGHT_GREEN = "D4EDDA"
LIGHT_BLUE  = "D0E8F5"
LIGHT_AMBER = "FFF3CD"

RATING_COLORS = {
    "exceeding": LIGHT_BLUE,
    "achieving": LIGHT_GREEN,
    "working towards": LIGHT_AMBER,
    "below": "FADADD",
}

def _set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def _set_cell_border_bottom(cell, hex_color="CCCCCC", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:color"), hex_color)
    tcBorders.append(bottom)
    tcPr.append(tcBorders)

def _para_border_bottom(para, hex_color="C85A1A", sz="12"):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:color"), hex_color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def _rating_color(rating_text):
    lower = rating_text.lower()
    for key, color in RATING_COLORS.items():
        if key in lower:
            return color
    return "EEEEEE"

def create_word_doc(sections, perf_table_rows, student_name, teacher_name, subject):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # --- Header band: branding left, logo placeholder right ---
    hdr_table = doc.add_table(rows=1, cols=2)
    hdr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_table.columns[0].width = Inches(4.5)
    hdr_table.columns[1].width = Inches(1.8)

    left_cell = hdr_table.cell(0, 0)
    left_para = left_cell.paragraphs[0]
    run = left_para.add_run("★  StarsEdu Online Tuition")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = ORANGE
    left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    right_cell = hdr_table.cell(0, 1)
    right_para = right_cell.paragraphs[0]
    run2 = right_para.add_run("PLACE\nLOGO")
    run2.font.size = Pt(9)
    run2.font.color.rgb = ORANGE
    right_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # --- Title ---
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run("Homework Feedback")
    title_run.bold = True
    title_run.font.size = Pt(26)
    title_run.font.color.rgb = NAVY

    # Orange decorative rule under title
    rule_para = doc.add_paragraph()
    rule_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rule_run = rule_para.add_run("─────")
    rule_run.font.color.rgb = ORANGE
    rule_run.font.size = Pt(14)

    doc.add_paragraph()

    # --- Info band table ---
    today_str = date.today().strftime("%B %Y")
    info_table = doc.add_table(rows=2, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["STUDENT NAME", "COURSE SUBJECT", "ASSIGNED TUTOR", "REPORTING PERIOD"]
    values = [student_name, subject, teacher_name, today_str]

    for col_idx, (lbl, val) in enumerate(zip(labels, values)):
        label_cell = info_table.cell(0, col_idx)
        value_cell = info_table.cell(1, col_idx)

        _set_cell_bg(label_cell, BEIGE)
        _set_cell_bg(value_cell, BEIGE)

        lp = label_cell.paragraphs[0]
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lr = lp.add_run(lbl)
        lr.font.size = Pt(7.5)
        lr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

        vp = value_cell.paragraphs[0]
        vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        vr = vp.add_run(val)
        vr.bold = True
        vr.font.size = Pt(12)
        vr.font.color.rgb = NAVY

    doc.add_paragraph()

    # --- Performance summary table (if AI provided rows) ---
    if perf_table_rows:
        section_heading = doc.add_paragraph()
        _para_border_bottom(section_heading, hex_color="CCCCCC", sz="4")
        hr = section_heading.add_run("Academic Performance Summary")
        hr.bold = True
        hr.font.size = Pt(14)
        hr.font.color.rgb = NAVY

        doc.add_paragraph()

        perf_table = doc.add_table(rows=1 + len(perf_table_rows), cols=2)
        perf_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Header row
        hrow = perf_table.rows[0]
        for cell, txt in zip(hrow.cells, ["EVALUATION DOMAIN", "CURRENT STANDING"]):
            _set_cell_bg(cell, DARK)
            p = cell.paragraphs[0]
            r = p.add_run(txt)
            r.bold = True
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        for row_idx, (domain, rating) in enumerate(perf_table_rows):
            row = perf_table.rows[row_idx + 1]
            # Domain cell
            _set_cell_border_bottom(row.cells[0])
            dp = row.cells[0].paragraphs[0]
            dr = dp.add_run(domain)
            dr.bold = True
            dr.font.size = Pt(10)
            dr.font.color.rgb = NAVY
            # Rating cell
            _set_cell_border_bottom(row.cells[1])
            bg = _rating_color(rating)
            _set_cell_bg(row.cells[1], bg)
            rp = row.cells[1].paragraphs[0]
            rp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr = rp.add_run(rating.upper())
            rr.bold = True
            rr.font.size = Pt(9)
            rr.font.color.rgb = NAVY

        doc.add_paragraph()

    # --- Narrative sections ---
    for sec_num, (sec_title, sec_body) in enumerate(sections, start=1):
        # Section heading with bottom border
        heading_para = doc.add_paragraph()
        _para_border_bottom(heading_para, hex_color="CCCCCC", sz="4")
        heading_run = heading_para.add_run(f"{sec_num}. {sec_title}")
        heading_run.bold = True
        heading_run.font.size = Pt(14)
        heading_run.font.color.rgb = NAVY

        doc.add_paragraph()

        # Body — support **bold** inline markup and bullet lines starting with "•" or "-"
        for line in sec_body.strip().split("\n"):
            line = line.strip()
            if not line:
                doc.add_paragraph()
                continue

            is_bullet = line.startswith("•") or line.startswith("-") or line.startswith("*")
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if is_bullet:
                para.style = doc.styles["List Bullet"]
                line = re.sub(r"^[•\-\*]\s*", "", line)

            # Split on **bold** markers
            parts = re.split(r"(\*\*[^*]+\*\*)", line)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    r = para.add_run(part[2:-2])
                    r.bold = True
                    r.font.size = Pt(11)
                    r.font.color.rgb = NAVY
                else:
                    r = para.add_run(part)
                    r.font.size = Pt(11)

        doc.add_paragraph()

    # --- Footer ---
    footer_para = doc.add_paragraph()
    _para_border_bottom(footer_para, hex_color="CCCCCC", sz="4")
    footer_run = footer_para.add_run(
        f"STARS_EDU // Academic Reporting System{' ' * 60}Page 1"
    )
    footer_run.font.size = Pt(8)
    footer_run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ==========================================
# 4. Parse AI output into sections + perf table
# ==========================================
def parse_ai_output(text):
    perf_rows = []
    sections = []

    # Extract performance table block
    table_match = re.search(
        r"===PERFORMANCE_TABLE===\s*(.*?)\s*===END_TABLE===",
        text, re.DOTALL
    )
    if table_match:
        table_text = table_match.group(1)
        text = text[:table_match.start()] + text[table_match.end():]
        for line in table_text.strip().split("\n"):
            if "|" in line:
                parts = [p.strip() for p in line.split("|")]
                if len(parts) >= 2 and parts[0].lower() not in ("domain", "evaluation domain", ""):
                    perf_rows.append((parts[0], parts[1]))

    # Extract sections by ===SECTION: Title=== delimiter
    section_pattern = re.split(r"===SECTION:\s*(.*?)===", text)
    if len(section_pattern) > 1:
        for i in range(1, len(section_pattern), 2):
            title = section_pattern[i].strip()
            body  = section_pattern[i + 1].strip() if i + 1 < len(section_pattern) else ""
            if title and body:
                sections.append((title, body))
    else:
        # Fallback: treat whole text as a single section
        sections.append(("Feedback", text.strip()))

    return perf_rows, sections


# ==========================================
# 5. Main Application Interface
# ==========================================
st.title("🎓 Professional Feedback Portal")
st.write("Generate high-quality Word reports for your students instantly.")

with st.sidebar:
    st.header("📋 Report Details")
    teacher_name = st.text_input("Teacher Name")
    student_name = st.text_input("Student Name")
    subject      = st.text_input("Subject")

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
                subject_line = f"Subject: {subject}" if subject else ""
                prompt = f"""
You are a senior international tutor named {teacher_name}.
Analyse the student work provided and produce a structured homework feedback report for {student_name}.
{subject_line}

IMPORTANT FORMATTING RULES — follow exactly:
- Do NOT write the student name, date, tutor name, or any header at the top. Those will be added automatically.
- Structure your output using the delimiters below — no other section markers.
- Use **bold** for key terms or sub-headings within body text.
- Use bullet points (starting with •) for lists.

Step 1 — Output a performance table block like this (keep the exact delimiters, list 3-5 skill domains relevant to the work):
===PERFORMANCE_TABLE===
Domain | Rating
Reading Comprehension | Achieving Expectations
Written Expression | Exceeding Expectations
===END_TABLE===

Ratings must be one of: Exceeding Expectations | Achieving Expectations | Working Towards Expectations | Below Expectations

Step 2 — Output narrative sections using this pattern (choose section titles appropriate to the subject):
===SECTION: Executive Summary & Approach to Learning===
Two or three paragraph overview of overall performance, attitude, and the key focus area for development.

===SECTION: Detailed Marking & Question Review===
Go through each question or task. For each: state what was attempted, what was correct, what was incorrect, and the conceptual gap if any.

===SECTION: Key Knowledge Gaps===
Specific topics or skills the student needs to strengthen, with brief explanation of why.

===SECTION: Strategic Objectives for the Upcoming Sessions===
• **Objective 1 name:** explanation
• **Objective 2 name:** explanation
• **Objective 3 name:** explanation

===SECTION: Actionable Guidance for Home Study===
• **Tip 1 name:** explanation
• **Tip 2 name:** explanation
• **Tip 3 name:** explanation

Language: English only. Be specific, encouraging, and professional.
"""
                payload = [prompt, build_payload_part(homework_file, "Student Work")]
                if mark_scheme_file:
                    payload.append(build_payload_part(mark_scheme_file, "Mark Scheme"))

                try:
                    response = model.generate_content(payload)
                    raw_text = response.text

                    perf_rows, sections = parse_ai_output(raw_text)
                    word_bytes = create_word_doc(
                        sections, perf_rows, student_name, teacher_name,
                        subject if subject else "General"
                    )

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
