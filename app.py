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

# CSS to hide "Manage app", the top decoration, and the "Deploy" button
st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important;}
    </style>
    """, unsafe_allow_html=True)

# Sidebar for Logo and Settings
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.title("Settings")
    st.markdown("---")
    summary_depth = st.select_slider(
        "Analysis Depth",
        options=["Brief Overview", "Standard Analysis", "Deep Dive"],
        value="Standard Analysis"
    )
    depth_map = {"Brief Overview": 4, "Standard Analysis": 7, "Deep Dive": 12}
    sentence_count = depth_map[summary_depth]


# --- 3. Advanced Summary Logic ---
def create_narrative_summary(text, count):
    """Generates a connected narrative explaining the document's situation."""
    if not text.strip():
        return "The document appears to be empty or unreadable."

    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    # Narrative Structure
    intro = "### 📋 Executive Situation Report\n\nThis document provides an overview of the following key areas. "
    body = " ".join([str(s) for s in summary_sentences])

    # Adding 'Gap' logic as requested
    narrative = f"{intro}\n\n**Main Context:** {body}\n\n"
    narrative += "### 🔍 Observation & Analysis\n"
    narrative += "From a synthesized point of view, the document effectively covers these points but may require further investigation into specific data calculations or secondary sources if the context seems incomplete. "
    narrative += "\n\n**Conclusion:** The content is structured to guide the reader through its primary objectives as highlighted in the processed version."

    return narrative, summary_sentences


# --- 4. Document Processing (Highlighting) ---
def highlight_pdf(file_bytes, key_sentences):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for sent in key_sentences:
            search_term = str(sent)[:60]  # Search first 60 chars for precision
            text_instances = page.search_for(search_term)
            for inst in text_instances:
                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=(1, 1, 0))  # Yellow
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
    pdf.set_font("Arial", size=12)
    # Remove characters incompatible with Latin-1
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 10, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


def export_summary_docx(text):
    doc = Document()
    doc.add_heading('Document Analysis Summary', 0)
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# --- 6. Main App Flow ---
st.title("🖋️ Smart Highlighter & Narrative Analyst")
st.write("Upload your document to receive a yellow-highlighted version and a meaningful narrative summary.")

uploaded_file = st.file_uploader("Upload File", type=["pdf", "docx", "pptx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("Analyzing document context..."):
        # Text Extraction
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # Summary & Keypoints
        narrative_report, key_sentences = create_narrative_summary(raw_text, sentence_count)

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
    res_col1, res_col2 = st.columns([1, 1])

    with res_col1:
        st.subheader("📝 Narrative Analysis")
        st.markdown(narrative_report)

        st.divider()
        st.write("💾 **Download Summary In:**")
        s_col1, s_col2, s_col3 = st.columns(3)
        s_col1.download_button("TXT", narrative_report, f"Summary_{uploaded_file.name}.txt")
        s_col2.download_button("Word", export_summary_docx(narrative_report), f"Summary_{uploaded_file.name}.docx")
        s_col3.download_button("PDF", export_summary_pdf(narrative_report), f"Summary_{uploaded_file.name}.pdf")

    with res_col2:
        st.subheader("📄 Highlighted Preview")
        st.download_button(
            label=f"Download Highlighted {file_ext.upper()}",
            data=processed_doc,
            file_name=f"Highlighted_{uploaded_file.name}",
            mime=mime_type,
            use_container_width=True
        )

        if file_ext == "pdf":
            # Real-time PDF Preview
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            pdf_preview = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="700" type="application/pdf"></iframe>'
            st.markdown(pdf_preview, unsafe_allow_html=True)
        else:
            st.info(
                "Direct preview is available for PDF files. For other formats, please download the file to see yellow highlights.")

else:
    st.info("Please upload a document to begin analysis.")