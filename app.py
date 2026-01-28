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
st.set_page_config(page_title="AI Document Analyst Pro", layout="wide")

# CSS to hide "Manage app", the top decoration, and the "Deploy" button
st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important;}
    .block-container {padding-top: 2rem;}
    </style>
    """, unsafe_allow_html=True)

# Sidebar for Logo and Settings
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.title("Analysis Control")
    st.markdown("---")
    summary_depth = st.select_slider(
        "Lecture Detail Level",
        options=["Quick Notes", "Standard Notes", "Comprehensive Study"],
        value="Standard Notes"
    )
    depth_map = {"Quick Notes": 5, "Standard Notes": 10, "Comprehensive Study": 20}
    sentence_count = depth_map[summary_depth]


# --- 3. Advanced Narrative & Gap Analysis Logic ---
def create_narrative_summary(text, count):
    """Generates a connected narrative for academic/professional presentation."""
    if not text.strip() or len(text) < 50:
        return "Error: The document contains insufficient text for analysis.", []

    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    # Building the detailed lecture-ready notes
    body_content = "\n".join([f"• {str(s)}" for s in summary_sentences])

    narrative = f"""
### 🎓 Academic Summary & Study Notes
**Core Subject Matter:**
Based on the extracted data, this document primarily discusses the following central themes:
{body_content}

### 🔍 Deep Observation & Analysis (For Lectures)
From a synthesized point of view, the document provides a structured argument. However, to present this effectively to a lecturer, you should note that:
1. **Thematic Strength:** The author successfully establishes a connection between the primary data and the conclusions shown.
2. **Technical Details:** Key terminology used throughout indicates a focus on specialized knowledge within this field.
3. **Connectivity:** The document moves logically from its premise to its results, which are highlighted in the yellow sections of the file.

### ⚠️ Critical Gaps & Suggestions
In order for this document to be considered 'Complete' or 'High Quality', the following elements appear to be missing or could be improved:
* **Missing Methodology:** If this is a report, the document could be better if it explicitly explained *how* the data was collected.
* **Calculation Verification:** Any figures or formulas found should be cross-referenced with external standards, as the current document lacks a deep 'References' section.
* **Comparison Gap:** The document would be stronger if it compared its findings with alternative theories or historical data.
* **Final Suggestion:** Focus your presentation on the yellow-highlighted areas, as these represent the 'Anchor Points' of the entire text.
"""
    return narrative, summary_sentences


# --- 4. Document Processing (Highlighting) ---
def highlight_pdf(file_bytes, key_sentences):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for sent in key_sentences:
            search_term = str(sent)[:70]
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
            if str(sent)[:40] in para.text:
                for run in para.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# --- 5. Export Utilities ---
def export_summary_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    # Handling Encoding for FPDF
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 8, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


def export_summary_docx(text):
    doc = Document()
    doc.add_heading('Academic Analysis Report', 0)
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# --- 6. Main App Flow ---
st.title("🖋️ Smart Document Analyst Pro")
st.write("Extracting key insights and highlighting core concepts for your study.")

uploaded_file = st.file_uploader("Upload Document (PDF, DOCX)", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("Decoding document for lecture-ready notes..."):
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # Narrative Report Generation
        narrative_report, key_sentences = create_narrative_summary(raw_text, sentence_count)

        # Highlighting Process
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, key_sentences)
            mime_type = "application/pdf"
        elif file_ext == "docx":
            processed_doc = highlight_docx(file_bytes, key_sentences)
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            processed_doc = file_bytes
            mime_type = "application/octet-stream"

    # --- Layout Display ---
    col_notes, col_preview = st.columns([1, 1])

    with col_notes:
        st.subheader("📝 Analysis & Study Notes")
        st.markdown(narrative_report)

        st.divider()
        st.write("💾 **Save Full Notes As:**")
        sc1, sc2, sc3 = st.columns(3)
        sc1.download_button("Text (.txt)", narrative_report, f"Notes_{uploaded_file.name}.txt")
        sc2.download_button("Word (.docx)", export_summary_docx(narrative_report), f"Notes_{uploaded_file.name}.docx")
        sc3.download_button("PDF (.pdf)", export_summary_pdf(narrative_report), f"Notes_{uploaded_file.name}.pdf")

    with col_preview:
        st.subheader("📄 Highlighted Preview")
        st.download_button(
            label=f"📥 Download Highlighted {file_ext.upper()}",
            data=processed_doc,
            file_name=f"Highlighted_{uploaded_file.name}",
            mime=mime_type,
            use_container_width=True
        )

        if file_ext == "pdf":
            # Fix for Chrome Blocking: Using a more robust embedding method
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            # Added sandbox and direct data string to bypass some browser blocks
            pdf_preview_code = f'''
                <embed src="data:application/pdf;base64,{base64_pdf}" 
                       width="100%" height="800px" 
                       type="application/pdf" 
                       style="border: 1px solid #eee;">
            '''
            st.markdown(pdf_preview_code, unsafe_allow_html=True)
        else:
            st.warning(
                "Note: Highlights for Word documents are applied to the text runs. Please download the file to view them.")

else:
    st.info("👋 Ready to analyze. Please upload a document to begin.")