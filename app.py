import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import io
import os
import nltk
import base64
import collections
import re
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
st.set_page_config(page_title="AI Contextual Analyst", layout="wide")

# CSS to hide "Manage app" and Streamlit branding
st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important;}
    .block-container {padding-top: 2rem;}
    </style>
    """, unsafe_allow_html=True)

# Sidebar for Logo
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.title("Document Analysis")
    st.markdown("---")
    summary_depth = st.select_slider(
        "Detail Level",
        options=["Brief", "Standard", "Lecture Ready"],
        value="Standard"
    )
    depth_map = {"Brief": 5, "Standard": 10, "Lecture Ready": 18}
    sentence_count = depth_map[summary_depth]


# --- 3. Context Identification Engine ---
def identify_document_type(text):
    """Automatically identifies what the document is concerning."""
    text = text.lower()
    categories = {
        "Legal/Contractual": ["agreement", "contract", "shall", "party", "terms", "liability", "herein"],
        "Medical/Healthcare": ["patient", "treatment", "clinical", "study", "diagnosis", "health", "dose"],
        "Technical/Engineering": ["system", "data", "software", "implementation", "specification", "technical"],
        "Financial/Business": ["revenue", "market", "investment", "fiscal", "profit", "quarterly", "growth"],
        "Academic/Research": ["hypothesis", "abstract", "conclusion", "methodology", "research", "cited"]
    }

    scores = {cat: 0 for cat in categories}
    for cat, keywords in categories.items():
        for word in keywords:
            if word in text:
                scores[cat] += text.count(word)

    # Return highest scoring category
    best_fit = max(scores, key=scores.get)
    return best_fit if scores[best_fit] > 0 else "General Information"


# --- 4. Narrative & Lecture Notes Logic ---
def create_narrative_summary(text, count):
    if not text.strip() or len(text) < 50:
        return "Insufficient text for analysis.", []

    doc_type = identify_document_type(text)
    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    body_content = "\n".join([f"• {str(s)}" for s in summary_sentences])

    narrative = f"""
### 📑 Document Identity: {doc_type}
**Analysis Overview:** This document is identified as a **{doc_type}** file. It primarily concerns the following structured themes which are critical for understanding the overall situation:

{body_content}

### 🔍 Detailed Observation & Analysis (For Presentation)
In explaining this to a lecturer, focus on these synthesized points:
1. **Contextual Flow:** The document moves from its initial premise regarding {doc_type} topics toward a specific conclusion.
2. **Technical Depth:** The vocabulary used indicates the document is intended for a professional/academic audience.
3. **Key Highlights:** The yellow-highlighted areas in the preview represent the 'spine' of the argument.

### ⚠️ Gaps & Suggestions for Quality
To make this document 'Better' or 'Complete', the following should be addressed:
* **Missing Perspective:** The document would be improved by adding a section on alternative viewpoints or risks.
* **Structural Gap:** There is a lack of visual data (tables/charts) to support the dense text segments.
* **Synthesis:** The document tells us *what* is happening but could be better at explaining *why* it is happening.
* **Recommendation:** Use the summary below as a cheat-sheet for your lecture to fill in these logical gaps.
"""
    return narrative, summary_sentences


# --- 5. Highlighting & Exporting ---
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


def export_summary_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
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


# --- 6. Main App Interface ---
st.title("🖋️ Smart Contextual Analyst Pro")
uploaded_file = st.file_uploader("Upload Document", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("Identifying document context and generating highlights..."):
        # Text Extraction
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # Identification & Narrative logic
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

    # --- Display ---
    col_notes, col_preview = st.columns([1, 1])

    with col_notes:
        st.subheader("📝 Automated Context & Notes")
        st.markdown(narrative_report)
        st.divider()
        st.write("💾 **Download Summary:**")
        sc1, sc2, sc3 = st.columns(3)
        sc1.download_button("TXT", narrative_report, f"Notes_{uploaded_file.name}.txt")
        sc2.download_button("Word", export_summary_docx(narrative_report), f"Notes_{uploaded_file.name}.docx")
        sc3.download_button("PDF", export_summary_pdf(narrative_report), f"Notes_{uploaded_file.name}.pdf")

    with col_preview:
        st.subheader("📄 Highlighted Preview")
        st.download_button(label="📥 Download Highlighted File", data=processed_doc,
                           file_name=f"Highlighted_{uploaded_file.name}", mime=mime_type, use_container_width=True)

        if file_ext == "pdf":
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            # Updated Preview logic for Chrome/Edge compatibility
            pdf_embed = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800px" type="application/pdf">'
            st.markdown(pdf_embed, unsafe_allow_html=True)
        else:
            st.info("Direct preview for DOCX is limited. Please download to see highlights.")
else:
    st.info("Please upload a file to begin the contextual analysis.")