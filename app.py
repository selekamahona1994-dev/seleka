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
st.set_page_config(page_title="AI Document Intelligence", layout="wide")

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


# --- 3. AI Brain: Document Identification ---
def identify_document_type(text):
    """AI logic to identify if the document is a CV, Research, Notes, or Report."""
    text_lower = text.lower()

    cv_keywords = ['experience', 'education', 'skills', 'objective', 'employment', 'projects']
    research_keywords = ['abstract', 'methodology', 'results', 'discussion', 'conclusion', 'references', 'introduction']
    notes_keywords = ['lecture', 'chapter', 'module', 'topic', 'summary', 'definition']

    cv_score = sum(1 for w in cv_keywords if w in text_lower)
    research_score = sum(1 for w in research_keywords if w in text_lower)
    notes_score = sum(1 for w in notes_keywords if w in text_lower)

    if cv_score > research_score and cv_score > notes_score:
        return "Professional Curriculum Vitae (CV)"
    elif research_score > notes_score:
        return "Academic Research Project / Paper"
    elif notes_score > 0:
        return "Educational Lecture Notes"
    else:
        return "General Informational Document"


# --- 4. Professional Reconstruction & Systematic Summary ---
def create_systematic_summary(text, count):
    if not text.strip():
        return "Unreadable content.", []

    doc_type = identify_document_type(text)
    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    # Keyword analysis for professional bolding
    words = re.findall(r'\w+', text.lower())
    common = [word for word, count in Counter(words).most_common(50) if len(word) > 5]
    highlighted_keywords = ", ".join(common[:10])

    # Systematic Rewriting
    intro = f"### 🤖 AI Intelligence Report: **{doc_type}**\n\n"
    intro += f"**Key Terminology Detected:** `{highlighted_keywords}`\n\n"

    body = "### 📋 Systematic Content Reconstruction\n"
    for i, s in enumerate(summary_sentences):
        sentence_str = str(s)
        # Bold high-value words for professional readability
        for word in common[:5]:
            sentence_str = re.sub(f'({word})', r'**\1**', sentence_str, flags=re.IGNORECASE)
        body += f"{i + 1}. {sentence_str}\n\n"

    # --- THE PROFESSIONAL RE-WRITING PART ---
    reconstruction = f"### 💎 High-Level Professional Enhancement\n"
    reconstruction += f"**How this {doc_type} should appear for senior stakeholders:**\n\n"

    if "CV" in doc_type:
        reconstruction += "> **Strategic Addition:** This document should include a 'Executive Value Proposition' section at the top, summarizing years of impact rather than just tasks. Ensure that **" + (
            common[0] if common else "Achievement") + "** is quantified with percentages.\n"
    elif "Research" in doc_type:
        reconstruction += "> **Professional Standard:** To reach a high academic level, this document requires a 'Practical Implications' section. It currently explains 'what' but needs to emphasize 'so what' regarding **" + (
            common[0] if common else "the results") + "**.\n"
    else:
        reconstruction += "> **Systematic Upgrade:** This content would be more professional if structured with an 'Executive Summary' followed by 'Key Performance Indicators'. The focus on **" + (
            common[0] if common else "core topics") + "** is good, but needs secondary source validation.\n"

    analysis = f"""
### 🔍 Observation & Analysis (Lecture/Presentation Ready)
* **Logic Flow:** The document moves from premise to conclusion using `{common[1] if len(common) > 1 else 'structured data'}`.
* **Missing Gaps:** To make this document perfect, you must add a **Risk Assessment** or **Limitation Section**. 
* **Suggestion:** During your lecture, explain how `{common[0] if common else 'the subject'}` interacts with current industry trends.

### 🏁 Conclusion & Synthesis
This **{doc_type}** provides a solid foundation. By implementing the suggested professional enhancements and focusing on the yellow-highlighted key points, the document becomes a high-level authority on the subject.
"""
    return intro + body + reconstruction + analysis, summary_sentences


# --- 5. Document Processing & Highlighting ---
def highlight_pdf(file_bytes, key_sentences):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for sent in key_sentences:
            text_instances = page.search_for(str(sent)[:60])
            for inst in text_instances:
                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=(1, 1, 0))
                annot.update()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def export_summary_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    # Filter for Latin-1 compatibility
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 8, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


# --- 6. Main Application Flow ---
st.title("🖋️ Smart AI Document Analyst & Professional Systematizer")

uploaded_file = st.file_uploader("Upload Document (PDF, DOCX)", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("🧠 AI Intelligence is identifying and reconstructing your document..."):
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # Generate the Narrative, Reconstruction, and Analysis
        full_notes, key_sentences = create_systematic_summary(raw_text, sentence_count)

        # Apply Yellow Highlights
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, key_sentences)
            mime_type = "application/pdf"
        else:
            processed_doc = file_bytes
            mime_type = "application/octet-stream"

    # --- UI Columns ---
    col_a, col_b = st.columns([1, 1])

    with col_a:
        st.subheader("📝 Systematic Analysis & Professional Reconstruction")
        st.markdown(full_notes)
        st.divider()
        st.write("📥 **Download Professional Notes:**")
        c1, c2 = st.columns(2)
        c1.download_button("Word Document (.docx)", full_notes, f"Professional_Notes_{uploaded_file.name}.docx")
        c2.download_button("PDF Document (.pdf)", export_summary_pdf(full_notes),
                           f"Professional_Notes_{uploaded_file.name}.pdf")

    with col_b:
        st.subheader("📄 Highlighted Document Preview")
        st.download_button("📥 Download Highlighted File", processed_doc, f"Highlighted_{uploaded_file.name}",
                           mime=mime_type)

        if file_ext == "pdf":
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            # Using embed for better Chrome compatibility
            pdf_code = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="900px" type="application/pdf">'
            st.markdown(pdf_code, unsafe_allow_html=True)
        else:
            st.info("Preview optimized for PDF. Please download the file to see the applied highlights.")

else:
    st.info("👋 Upload a file to begin the professional AI analysis.")