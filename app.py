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


# --- 1. Setup & NLTK Resources ---
def ensure_nltk_resources():
    resources = ['punkt', 'punkt_tab']
    for res in resources:
        try:
            nltk.data.find(f'tokenizers/{res}')
        except LookupError:
            nltk.download(res)


ensure_nltk_resources()

# --- 2. UI Customization & Hiding Manage App ---
st.set_page_config(page_title="Pro Doc Alchemist", layout="wide")

st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    .reportview-container .main .block-container { padding-top: 1rem; }
    </style>
    """, unsafe_allow_html=True)

# Sidebar with Logo
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.title("Control Panel")
    st.markdown("---")
    summary_depth = st.select_slider("Summary Detail Level", options=["Brief", "Standard", "Deep"], value="Standard")

    depth_map = {"Brief": 3, "Standard": 6, "Deep": 12}
    sentence_count = depth_map[summary_depth]


# --- 3. Logic Functions ---

def generate_narrative_summary(text, count):
    """Creates a connected, narrative-style summary."""
    if not text.strip(): return "No readable text found in document."

    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    # Building a connected narrative
    intro = "The analyzed document presents several critical findings. "
    body = " ".join([str(s) for s in summary_sentences])
    conclusion = "\n\nIn conclusion, the document emphasizes these core themes as vital for the reader's understanding."

    full_narrative = f"{intro}\n\n{body}\n\n{conclusion}"
    return full_narrative


def create_summary_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    # Sanitize text for FPDF (removes non-latin1 chars)
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 10, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


def create_summary_docx(text):
    doc = Document()
    doc.add_heading('Document Analysis Summary', 0)
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def highlight_pdf(file_bytes, key_sentences):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for sent in key_sentences:
            # Match shorter chunks for better hit rates in PDF
            search_term = str(sent)[:50]
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


# --- 4. Main Application Interface ---

st.title("🖋️ Smart Highlighter & Narrative Summarizer")
uploaded_file = st.file_uploader("Upload Document (PDF, DOCX, PPTX)", type=["pdf", "docx", "pptx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("Processing document logic..."):
        # 1. Text Extraction
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([page.get_text() for page in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # 2. Summary Generation (The Narrative)
        parser = PlaintextParser.from_string(raw_text, Tokenizer("english"))
        summarizer = LsaSummarizer()
        raw_sentences = summarizer(parser.document, sentence_count)
        full_summary = generate_narrative_summary(raw_text, sentence_count)

        # 3. Highlighting
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, raw_sentences)
            mime_type = "application/pdf"
        elif file_ext == "docx":
            processed_doc = highlight_docx(file_bytes, raw_sentences)
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            processed_doc = file_bytes
            mime_type = "application/octet-stream"

    # --- UI Display ---
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📝 Narrative Summary")
        st.write(full_summary)

        st.markdown("---")
        st.write("📥 **Download Summary As:**")
        dl_col1, dl_col2, dl_col3 = st.columns(3)

        dl_col1.download_button("TXT", full_summary, f"Summary_{uploaded_file.name}.txt")
        dl_col2.download_button("Word", create_summary_docx(full_summary), f"Summary_{uploaded_file.name}.docx")
        dl_col3.download_button("PDF", create_summary_pdf(full_summary), f"Summary_{uploaded_file.name}.pdf")

    with col2:
        st.subheader("👁️ Preview & Highlighted File")
        st.download_button(
            label=f"Download Highlighted {file_ext.upper()}",
            data=processed_doc,
            file_name=f"Highlighted_{uploaded_file.name}",
            mime=mime_type,
            use_container_width=True
        )

        if file_ext == "pdf":
            # PDF Preview using Base64
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
            st.markdown(pdf_display, unsafe_allow_html=True)
        else:
            st.info("Preview is optimized for PDF. Please download the file to view highlights for Word/PPTX.")

else:
    st.info("👋 Welcome! Please upload a file to begin the automated highlighting and narrative summary.")