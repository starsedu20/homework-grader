import streamlit as st
import google.generativeai as genai
import os
from fpdf import FPDF
from datetime import date
from io import BytesIO

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
# 3. PDF Generation Function (Unicode Safe)
# ==========================================
def create_pdf(text, student_name, teacher_name):
    today = date.today().strftime("%B %d, %Y")
    pdf = FPDF()
    pdf.add_page()
    
    # Header Section
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, txt="STUDENT STUDY REPORT", ln=True, align='C')
    
    pdf.set_font("Arial", size=11)
    pdf.cell(0, 10, txt=f"Date: {today}", ln=True, align='R')
    pdf.cell(0, 7, txt=f"Student: {student_name}", ln=True, align='L')
    pdf.cell(0, 7, txt=f"Tutor: {teacher_name}", ln=True, align='L')
    pdf.ln(5)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)
    
    # Clean text for PDF (swapping Unicode for safe characters)
    clean_text = text.replace("✓", "[Correct]").replace("✗", "[Incorrect]")
    clean_text = clean_text.replace("\u2713", "[Correct]").replace("\u2717", "[Incorrect]")
    clean_text = clean_text.replace("**", "").replace("* ", "- ")
    
    pdf.set_font("Arial", size=11)
    # Wrap text automatically
    pdf.multi_cell(0, 8, txt=clean_text.encode('latin-1', 'replace').decode('latin-1'))
    
    return pdf.output(dest='S').encode('latin-1', 'ignore')

# ==========================================
# 4. Main Application Interface
# ==========================================
st.title("🎓 Professional Feedback Portal")
st.write("Generate high-quality PDF reports for your students instantly.")

with st.sidebar:
    st.header("📋 Report Details")
    teacher_name = st.text_input("Teacher Name")
    student_name = st.text_input("Student Name")
    
    st.write("---")
    st.header("📂 Upload Work")
    homework_file = st.file_uploader("Upload Student Work (PDF/Image/Word)", type=["pdf", "png", "jpg", "jpeg", "docx", "doc"])
    mark_scheme_file = st.file_uploader("Upload Mark Scheme (Optional)", type=["pdf", "png", "jpg", "jpeg", "docx", "doc"])

def build_payload_part(uploaded_file, label):
    filename = uploaded_file.name.lower()
    if filename.endswith(".docx") or filename.endswith(".doc"):
        from docx import Document
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
    
    if st.button("🚀 Process & Generate PDF Report", use_container_width=True):
        if not teacher_name or not student_name:
            st.warning("⚠️ Please enter both Teacher and Student names in the sidebar.")
        elif not model:
            st.error("AI model is not responding. Check your API key.")
        else:
            with st.spinner(f"Analyzing {student_name}'s work..."):
                # Professional Prompt addressing specific science/math needs
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
                    
                    # Generate the PDF in background
                    pdf_bytes = create_pdf(report_content, student_name, teacher_name)
                    
                    st.success(f"🎉 Report for {student_name} is ready!")
                    
                    # Centered Download Button
                    st.download_button(
                        label="📥 Download PDF Report",
                        data=pdf_bytes,
                        file_name=f"{student_name}_Report_{date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error during report generation: {e}")
else:
    st.info("Upload a student's homework to begin.")

st.markdown("---")
st.caption(f"System Date: {date.today().strftime('%Y-%m-%d')} | Powered by Gemini 1.5 Pro")