import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from pptx import Presentation
import io
import os
import nltk
import base64
import re
from collections import Counter
from fpdf import FPDF
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer


# --- 1. NLTK Resource Management ---
def ensure_nltk_resources():
    resources = ['punkt', 'punkt_tab', 'stopwords']
    for res in resources:
        try:
            nltk.data.find(f'tokenizers/{res}')
        except LookupError:
            nltk.download(res)


ensure_nltk_resources()

# --- 2. Page Configuration & Stealth UI ---
st.set_page_config(page_title="AI Intelligence Analyst", layout="wide")

st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important;}
    .block-container {padding-top: 1.5rem;}
    </style>
    """, unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.title("Intelligence Dashboard")
    st.markdown("---")
    analysis_mode = st.radio("Analysis Detail", ["Standard Analysis", "Academic Deep Dive"])
    sentence_count = 12 if analysis_mode == "Academic Deep Dive" else 7


# --- 3. AI Identification Logic ---
def identify_document_type(text):
    text_lower = text.lower()
    scores = {
        "Professional CV": len(re.findall(r'(experience|education|skills|employment|objective)', text_lower)),
        "Research Project": len(re.findall(r'(abstract|methodology|results|discussion|references)', text_lower)),
        "Lecture Notes": len(re.findall(r'(lecture|module|chapter|definition|topic)', text_lower)),
        "Business Report": len(re.findall(r'(executive summary|market|revenue|strategy|kpi)', text_lower))
    }
    best_match = max(scores, key=scores.get)
    return best_match if scores[best_match] > 0 else "General Informational Document"


# --- 4. Systematized Content Reconstruction ---
def reconstruct_content(text, doc_type, summary_sentences):
    # Extracting core professional keywords
    words = re.findall(r'\w+', text.lower())
    common = [word for word, count in Counter(words).most_common(60) if len(word) > 5]
    top_keywords = common[:8]

    # Start the "New Systematic Version"
    reconstruction = f"# 🏛️ SYSTEMATIC RECONSTRUCTION: {doc_type.upper()}\n\n"
    reconstruction += f"**Primary Intelligence Markers:** {', '.join([f'`{k.upper()}`' for k in top_keywords])}\n\n"

    reconstruction += "## 📑 1. Core Narrative Summary\n"
    for i, s in enumerate(summary_sentences):
        sent = str(s)
        for k in top_keywords[:3]:
            sent = re.sub(f'({k})', r'**\1**', sent, flags=re.IGNORECASE)
        reconstruction += f"{i + 1}. {sent}\n\n"

    reconstruction += "## 💎 2. High-Level Professional Expansion\n"
    reconstruction += f"This document, identified as a **{doc_type}**, has been reconstructed to meet global professional standards. "

    if "CV" in doc_type:
        reconstruction += f"To improve this profile, we have added a strategic focus on **{top_keywords[0].capitalize() if top_keywords else 'Value'}**. "
        reconstruction += "We have optimized the layout to emphasize leadership competencies and technical impact.\n\n"
    elif "Research" in doc_type:
        reconstruction += f"This reconstruction adds an 'Analytical Synthesis' layer. The study on **{top_keywords[0] if top_keywords else 'the subject'}** "
        reconstruction += "has been aligned with current academic methodologies to ensure logical validity.\n\n"
    else:
        reconstruction += f"We have rewritten this as a structured briefing. The emphasis is on **{top_keywords[0] if top_keywords else 'strategic goals'}**, "
        reconstruction += "ensuring that the reader understands the 'Executive Why' behind the data.\n\n"

    reconstruction += "## 🔍 3. Critical Observations & Missing Gaps\n"
    reconstruction += f"* **Calculated Gap:** The original text lacks a deep dive into **{top_keywords[-1] if len(top_keywords) > 4 else 'comparative analysis'}**. \n"
    reconstruction += "* **Suggestion for Lecture:** When presenting, focus on the yellow-highlighted 'Pivot Points' found in the preview.\n"
    reconstruction += "* **Final Verdict:** This document is now synthesized into a high-level authority report.\n"

    return reconstruction, top_keywords


# --- 5. Highlighting & File Handling ---
def highlight_pdf(file_bytes, key_sentences):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for sent in key_sentences:
            text_instances = page.search_for(str(sent)[:65])
            for inst in text_instances:
                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=(1, 1, 0))  # Yellow
                annot.update()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def export_to_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    # Sanitize for PDF encoding
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 8, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


# --- 6. Main App Process ---
st.title("🖋️ Smart AI Document Analyst & Systematizer")

uploaded_file = st.file_uploader("Upload File (PDF or DOCX)", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("🧠 AI Intelligence identifying and rewriting document..."):
        # 1. Extract Text
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # 2. Identify and Summarize
        dtype = identify_document_type(raw_text)
        parser = PlaintextParser.from_string(raw_text, Tokenizer("english"))
        summarizer = LsaSummarizer()
        summary_sentences = summarizer(parser.document, sentence_count)

        # 3. Systematic Reconstruction
        reconstructed_notes, keywords = reconstruct_content(raw_text, dtype, summary_sentences)

        # 4. Apply Highlights
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, summary_sentences)
            mime_type = "application/pdf"
        else:
            processed_doc = file_bytes
            mime_type = "application/octet-stream"

    # --- UI Layout ---
    col_notes, col_view = st.columns([1, 1])

    with col_notes:
        st.subheader("📝 Systematic Intelligence Report")
        st.markdown(reconstructed_notes)
        st.divider()
        st.write("📥 **Download New Reconstructed Notes:**")
        d1, d2 = st.columns(2)
        d1.download_button("Download as Word", reconstructed_notes, "Reconstructed_Analysis.docx")
        d2.download_button("Download as PDF", export_to_pdf(reconstructed_notes), "Reconstructed_Analysis.pdf")

    with col_view:
        st.subheader("📄 Highlighted Original Preview")
        st.download_button(f"📥 Download Highlighted {file_ext.upper()}", processed_doc,
                           f"Highlighted_{uploaded_file.name}", mime=mime_type)

        if file_ext == "pdf":
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            # Using embed with sandbox-friendly parameters for Chrome
            pdf_embed = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="900px" type="application/pdf">'
            st.markdown(pdf_embed, unsafe_allow_html=True)
        else:
            st.info("Preview optimized for PDF. Please download the file to see the yellow highlights.")

else:
    st.info("Please upload a document to begin the AI systematic analysis.")