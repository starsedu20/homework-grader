import streamlit as st
import google.generativeai as genai
import os
import io
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date

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

# Canonical section names — the code will auto-number these 1–6
SECTIONS = [
    "Executive Summary",
    "Detailed Marking",
    "Key Gaps",
    "Next Steps",
    "Foundation (To Consolidate)",
    "Extension (To Challenge)",
]

# Lines containing any of these strings are silently dropped
REMOVE_PATTERNS = [
    r"STARS_EDU",
    r"Academic Reporting System",
    r"Senior International Tutor",
    # Header block the AI sometimes adds
    r"^Study Report",
    r"^Student\s*:",
    r"^Tutor\s*:",
    r"^Date\s*:",
    r"^To\s*:",
    r"^From\s*:",
    r"^Homework Feedback$",   # title will be added by the code itself
]

def should_remove(line: str) -> bool:
    for pat in REMOVE_PATTERNS:
        if re.search(pat, line.strip(), re.IGNORECASE):
            return True
    return False

def match_section(line: str):
    """Return (section_index, canonical_label) if line is a section header, else None."""
    # Strip leading numbers like "1. " or "1) " that the AI might add
    clean = re.sub(r'^\d+[\.\)]\s*', '', line.strip()).rstrip(':').strip().lower()
    for i, name in enumerate(SECTIONS):
        # Exact match (case-insensitive)
        if clean == name.lower():
            return i, name
        # Prefix match — handles variants like "Key Gaps for Attention"
        if clean.startswith(name.lower()):
            return i, name
    return None

def is_numbered_item(line: str) -> bool:
    return bool(re.match(r'^\d+[\)\.]\s+\S', line.strip()))

def create_docx(text: str, student_name: str, teacher_name: str) -> bytes:
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    # Title
    title_p = doc.add_paragraph()
    run = title_p.add_run("Homework Feedback")
    run.bold = True
    run.font.size = Pt(22)
    title_p.paragraph_format.space_after = Pt(14)

    # Split AI text at "Best regards" so we control the sign-off
    body_text, _, _ = text.partition("Best regards")

    lines = body_text.split("\n")
    prev_blank = False

    for raw in lines:
        line = raw.strip()

        # Drop unwanted lines
        if should_remove(line):
            continue

        # Collapse consecutive blank lines
        if not line:
            if prev_blank:
                continue
            prev_blank = True
            doc.add_paragraph().paragraph_format.space_after = Pt(0)
            continue
        prev_blank = False

        # Section header
        m = match_section(line)
        if m is not None:
            idx, name = m
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after  = Pt(2)
            run = p.add_run(f"{idx + 1}.  {name}:")
            run.bold = True
            run.font.size = Pt(12)
            continue

        # Numbered sub-item  e.g. "1). Grammar…"
        if is_numbered_item(line):
            body = re.sub(r'^\d+[\)\.]\s*', '', line)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0)
            p.paragraph_format.space_after = Pt(4)
            # Write the number manually so formatting is consistent
            num_match = re.match(r'^(\d+[\)\.])(.+)', line)
            if num_match:
                num_part  = num_match.group(1)
                body_part = num_match.group(2).strip()
                r1 = p.add_run(num_part + "  ")
                r1.font.size = Pt(11)
                r2 = p.add_run(body_part)
                r2.font.size = Pt(11)
            else:
                p.add_run(line).font.size = Pt(11)
            continue

        # Normal paragraph
        p = doc.add_paragraph(line)
        for r in p.runs:
            r.font.size = Pt(11)
        p.paragraph_format.space_after = Pt(4)

    # Controlled sign-off
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Best regards,").font.size = Pt(11)

    p2 = doc.add_paragraph()
    r = p2.add_run(teacher_name)
    r.bold = True
    r.font.size = Pt(11)
    p2.paragraph_format.space_before = Pt(6)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ==========================================
# 4. Main Application Interface
# ==========================================
st.title("🎓 Professional Feedback Portal")
st.write("Generate high-quality feedback reports for your students instantly.")

with st.sidebar:
    st.header("📋 Report Details")
    teacher_name = st.text_input("Teacher Name", placeholder="e.g., Harry")
    student_name = st.text_input("Student Name", placeholder="e.g., Nedime")

    st.write("---")
    st.header("📂 Upload Work")
    homework_file    = st.file_uploader("Upload Student Work (PDF/Image)", type=["pdf", "png", "jpg", "jpeg"])
    mark_scheme_file = st.file_uploader("Upload Mark Scheme (Optional)",   type=["pdf", "png", "jpg", "jpeg"])

if homework_file:
    st.info(f"File '{homework_file.name}' ready for analysis.")

    if st.button("🚀 Process & Generate Feedback Report", use_container_width=True):
        if not teacher_name or not student_name:
            st.warning("⚠️ Please enter both Teacher and Student names in the sidebar.")
        elif not model:
            st.error("AI model is not responding. Check your API key.")
        else:
            with st.spinner(f"Analysing {student_name}'s work…"):

                prompt = f"""You are a tutor named {teacher_name}.
Write a personalised homework feedback report addressed to your student, {student_name}.

Tone: professional, encouraging, and direct (use "I" and "you").
Today's date: {date.today().strftime("%B %d, %Y")}.

IMPORTANT — begin the report immediately with the greeting below.
Do NOT add any title block, header, or metadata before the greeting.
Do NOT write lines such as "Study Report:", "Student:", "Tutor:", "Date:", or "Homework Feedback" at the top.

Start with:

Dear {student_name},

[One warm opening sentence.]

Then include each of these sections in order, using EXACTLY these headings (no numbering — the formatting system will add numbers):

Executive Summary:
[2–3 sentences summarising overall performance.]

Detailed Marking:
[For each question or paragraph, give it a clear label and mark it [Correct], [Correct – Minor Error], or [Incorrect]. Explain what was right or wrong and give a specific suggestion.]

Key Gaps:
[Number each gap 1). 2). 3). Focus on the most important areas needing improvement.]

Next Steps:
[One brief sentence introducing the steps below.]

Foundation (To Consolidate):
[Number each step 1). 2). 3). Practical actions to reinforce core understanding.]

Extension (To Challenge):
[Number each step 1). 2). 3). Stretch tasks for deeper learning.]

[One closing encouragement sentence addressed to {student_name} by name.]

Best regards,
{teacher_name}

Rules:
- Do NOT include any title, header block, or metadata at the top.
- Do NOT write your job title, "Senior International Tutor", "STARS_EDU", or any footer.
- Do NOT add more than one blank line between sections.
- Use plain numbered items 1). 2). 3). — no bullet symbols.
- Write in English only.
"""

                def prep(f):
                    return {"mime_type": f.type, "data": f.getvalue()}

                payload = [prompt, prep(homework_file)]
                if mark_scheme_file:
                    payload.append(prep(mark_scheme_file))

                try:
                    response = model.generate_content(payload)
                    report_text = response.text

                    docx_bytes = create_docx(report_text, student_name, teacher_name)

                    st.success(f"🎉 Feedback report for {student_name} is ready!")

                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📥 Download Word (.docx)",
                            data=docx_bytes,
                            file_name=f"{student_name}_Feedback_{date.today()}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )

                    with st.expander("Preview report text"):
                        st.text(report_text)

                except Exception as e:
                    st.error(f"Error during report generation: {e}")
else:
    st.info("👈 Upload the student's work in the sidebar to begin.")

st.markdown("---")
st.caption(f"System Date: {date.today().strftime('%Y-%m-%d')}")
