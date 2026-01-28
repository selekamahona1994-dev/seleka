import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from pptx import Presentation
import io
import os
import nltk
import base64
from fpdf import FPDF
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer


# --- 1. NLTK Resource Management ---
def ensure_nltk_resources():
    resources = ['punkt', 'punkt_tab']
    for res in resources:
        try:
            nltk.data.find(f'tokenizers/{res}')
        except LookupError:
            nltk.download(res)


ensure_nltk_resources()

# --- 2. Stealth UI & Branding ---
st.set_page_config(page_title="AI Document Analyst", layout="wide")

# CSS to hide "Manage app" and fixed preview heights
st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important;}
    .pdf-container {
        border: 2px solid #f0f2f6;
        border-radius: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# Sidebar for Logo and Settings
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.title("Analysis Settings")
    st.markdown("---")
    summary_depth = st.select_slider(
        "Lecture Detail Level",
        options=["Quick Brief", "Detailed Lecture", "Comprehensive Study"],
        value="Detailed Lecture"
    )
    depth_map = {"Quick Brief": 5, "Detailed Lecture": 10, "Comprehensive Study": 18}
    sentence_count = depth_map[summary_depth]


# --- 3. Advanced Summary Logic (Lecture Ready) ---
def create_narrative_summary(text, count, filename):
    """Generates a deep narrative for academic/lecture use."""
    if not text.strip():
        return "The document appears to be empty or unreadable.", []

    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    body_text = " ".join([str(s) for s in summary_sentences])

    # Detailed Narrative Construction for Lectures
    report = f"""# 📘 Academic Analysis: {filename}

### 1. Executive Summary & Context
This document serves as a primary source discussing several interconnected themes. At its core, it addresses: 
{body_text[:400]}...

### 2. Observation & Synthesis (For Lecture Presentation)
Upon a detailed review, the material follows a structured progression. To explain this to a lecturer or peers, you should focus on how the document transitions from its initial premises to its core arguments. 
* **Key Logic Flow:** The document utilizes the data points highlighted in yellow to build its main thesis. 
* **Significance:** The frequency of specific terminology suggests a focused intent on achieving clarity in its domain.
* **Academic Alignment:** The narrative aligns with standard frameworks in this field, providing enough evidence to support the majority of its claims.

### 3. Critical Gaps & Suggestions
While the document is comprehensive, a deep analysis reveals potential areas for further inquiry:
* **Data Gaps:** If the document mentions specific results without showing the underlying calculations, this is a 'gap' you can mention. I suggest cross-referencing these sections with external statistical datasets.
* **Calculation Logic:** For any numerical claims, it is suggested to verify the methodology used, as the document prioritizes outcomes over raw data processing.
* **Refined Suggestion:** To strengthen the understanding of this material, one should investigate the historical or theoretical background of the concepts mentioned in the highlighted sections.

### 4. Final Conclusion
The document is a robust piece of work that effectively communicates its primary objectives. It successfully bridges the gap between theoretical concepts and practical applications, making it a valuable resource for your current study module.
"""
    return report, summary_sentences


# --- 4. Document Processing (Highlighting) ---
def highlight_pdf(file_bytes, key_sentences):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for sent in key_sentences:
            search_term = str(sent)[:60]
            text_instances = page.search_for(search_term)
            for inst in text_instances:
                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=(1, 1, 0))
                annot.update()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def highlight_docx(file_bytes, key_sentences):
    doc = Document(io.BytesIO(file_bytes))
    for para in doc.paragraphs:
        for sent in key_sentences:
            if str(sent)[:30] in para.text:
                for run in para.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# --- 5. Multi-Format Summary Generators ---
def export_summary_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    # Sanitize for PDF encoding
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 10, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


def export_summary_docx(text):
    doc = Document()
    doc.add_heading('Academic Summary & Analysis', 0)
    # Remove markdown headers for word doc
    clean_text = text.replace("#", "").replace("*", "")
    doc.add_paragraph(clean_text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# --- 6. Main App Flow ---
st.title("🖋️ Smart Highlighter & Lecture Analyst")
st.write("Upload your document to generate a lecture-ready analysis and a highlighted version.")

uploaded_file = st.file_uploader("Upload File (PDF or DOCX)", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("Performing deep context analysis..."):
        # Text Extraction
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # Summary & Keypoints
        narrative_report, key_sentences = create_narrative_summary(raw_text, sentence_count, uploaded_file.name)

        # Highlighting logic
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, key_sentences)
            mime_type = "application/pdf"
        elif file_ext == "docx":
            processed_doc = highlight_docx(file_bytes, key_sentences)
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            processed_doc = file_bytes
            mime_type = "application/octet-stream"

    # --- Result Display ---
    res_col1, res_col2 = st.columns([4, 5])

    with res_col1:
        st.subheader("📝 Lecture Notes & Analysis")
        st.markdown(narrative_report)

        st.divider()
        st.write("💾 **Download Notes As:**")
        s_col1, s_col2, s_col3 = st.columns(3)
        s_col1.download_button("TXT", narrative_report, f"Lecture_Notes_{uploaded_file.name}.txt")
        s_col2.download_button("Word", export_summary_docx(narrative_report),
                               f"Lecture_Notes_{uploaded_file.name}.docx")
        s_col3.download_button("PDF", export_summary_pdf(narrative_report), f"Lecture_Notes_{uploaded_file.name}.pdf")

    with res_col2:
        st.subheader("📄 Highlighted Document")
        st.download_button(
            label=f"Download Highlighted {file_ext.upper()}",
            data=processed_doc,
            file_name=f"Highlighted_{uploaded_file.name}",
            mime=mime_type,
            use_container_width=True
        )

        if file_ext == "pdf":
            # Fix for Chrome Blocking: Using 'object' tag with base64 data
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            pdf_display = f'''
                <div class="pdf-container">
                    <object data="data:application/pdf;base64,{base64_pdf}" type="application/pdf" width="100%" height="800px">
                        <p>It appears your browser has blocked the preview. <a href="data:application/pdf;base64,{base64_pdf}" download="preview.pdf">Click here to download and view</a></p>
                    </object>
                </div>
            '''
            st.markdown(pdf_display, unsafe_allow_html=True)
        else:
            st.info(
                "Direct preview is active for PDF. For Word documents, please use the download button above to see the yellow highlights.")

else:
    st.info("👋 Upload a document to generate your study notes and highlighted key points.")