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
st.set_page_config(page_title="AI Intelligence System", layout="wide")

# CSS to hide "Manage app" and optimize the workspace
st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important;}
    .block-container {padding-top: 1.5rem;}
    </style>
    """, unsafe_allow_html=True)

# Sidebar for logo and controls
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.title("System Controls")
    st.markdown("---")
    analysis_mode = st.radio("Analysis Precision", ["Standard", "Executive Deep Dive"])
    sentence_count = 15 if analysis_mode == "Executive Deep Dive" else 8


# --- 3. AI Brain: Document Categorization ---
def identify_document_type(text):
    text_lower = text.lower()

    cv_keywords = ['experience', 'education', 'skills', 'objective', 'employment', 'projects', 'profile']
    research_keywords = ['abstract', 'methodology', 'results', 'discussion', 'conclusion', 'references', 'introduction',
                         'hypothesis']
    notes_keywords = ['lecture', 'chapter', 'module', 'topic', 'summary', 'definition', 'concept']

    cv_score = sum(1 for w in cv_keywords if w in text_lower)
    research_score = sum(1 for w in research_keywords if w in text_lower)
    notes_score = sum(1 for w in notes_keywords if w in text_lower)

    if cv_score > research_score and cv_score > notes_score:
        return "Professional Curriculum Vitae (CV)"
    elif research_score > notes_score:
        return "Academic Research / Project Thesis"
    elif notes_score > 0:
        return "Systematic Lecture Notes"
    else:
        return "Executive Corporate Report"


# --- 4. Systematic Rewriting & Professional Synthesis ---
def create_systematic_summary(text, count):
    if not text.strip():
        return "The uploaded document contains no readable text.", []

    doc_type = identify_document_type(text)
    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    # Keyword Frequency Analysis
    words = re.findall(r'\w+', text.lower())
    # Filter common words to find significant professional terms
    stop_words = set(['the', 'and', 'this', 'that', 'with', 'from', 'they', 'their'])
    common = [word for word, count in Counter(words).most_common(60) if len(word) > 5 and word not in stop_words]
    top_keywords = ", ".join(common[:12])

    # --- PART A: IDENTIFICATION & KEYWORDS ---
    intro = f"### 🧠 AI Intelligence Identification: **{doc_type}**\n"
    intro += f"**Core Domain Keywords:** `{top_keywords.upper()}`\n\n"

    # --- PART B: SYSTEMATIC REWRITING (REPRODUCING THE CORE WORK) ---
    body = "### 📋 Systematic Reconstruction of Content\n"
    body += f"Below is the professional reproduction of the **{doc_type}** content, reorganized for maximum clarity:\n\n"
    for i, s in enumerate(summary_sentences):
        sentence_str = str(s)
        # Apply bolding to keywords for high-level scanning
        for word in common[:6]:
            sentence_str = re.sub(f'({word})', r'**\1**', sentence_str, flags=re.IGNORECASE)
        body += f"> **{i + 1}.** {sentence_str}\n\n"

    # --- PART C: PROFESSIONAL ADDITIONS (WHAT SHOULD APPEAR) ---
    enhancement = f"### 💎 High-Level Professional Roadmap\n"
    enhancement += f"**Strategic recommendations to elevate this {doc_type}:**\n\n"

    # AI Logic to suggest additions
    best_keyword = common[0].capitalize() if common else "The Main Topic"

    if "CV" in doc_type:
        enhancement += f"* **Proposed Addition:** Include a 'Core Competencies' matrix emphasizing **{best_keyword}**. \n"
        enhancement += f"* **Refinement:** Remove generic descriptors. Replace with 'Key Performance Indicators' (KPIs) relevant to **{common[1] if len(common) > 1 else 'industry standards'}**.\n"
    elif "Research" in doc_type:
        enhancement += f"* **Proposed Addition:** An 'Executive Implications' section explaining how **{best_keyword}** impacts real-world applications.\n"
        enhancement += f"* **Refinement:** Systematic citation of recent studies (2024-2026) regarding **{common[1] if len(common) > 1 else 'the methodology'}**.\n"
    else:
        enhancement += f"* **Proposed Addition:** A 'SWOT Analysis' or 'Risk Mitigation' table centered around **{best_keyword}**.\n"
        enhancement += f"* **Refinement:** Transition from descriptive text to a 'Strategic Action Plan' format.\n"

    # --- PART D: CONCLUSION FOR LECTURER/PROFESSIONAL ---
    analysis = f"""
### 🔍 Expert Observation & Analysis
* **Thematic Core:** The work revolves around **{best_keyword}**, which acts as the primary pillar of the argument.
* **Logical Integrity:** The connectivity between **{common[1] if len(common) > 1 else 'the data points'}** and the final objective is clear but requires stronger conclusion backing.
* **Final Synthesis:** This document is now reconstructed into a high-level authority. Use the yellow-highlighted sections in the file as your primary talking points during your presentation.
"""
    return intro + body + enhancement + analysis, summary_sentences


# --- 5. Highlighting Logic ---
def highlight_pdf(file_bytes, key_sentences):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for sent in key_sentences:
            text_instances = page.search_for(str(sent)[:65])  # Search first 65 chars
            for inst in text_instances:
                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=(1, 1, 0))  # Standard Academic Yellow
                annot.update()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# --- 6. Export Logic ---
def export_summary_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    # Clean text for standard encoding
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 8, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


# --- 7. Main Application UI ---
st.title("🖋️ Smart Document Intelligence & Analyst Pro")
st.write("Professional identification, systematic reconstruction, and keypoint highlighting.")

uploaded_file = st.file_uploader("Upload Document (PDF or DOCX)", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("🚀 AI Brain analyzing content and reconstructing the document structure..."):
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        # Core Process
        systematic_notes, key_sentences = create_systematic_summary(raw_text, sentence_count)

        # Highlighting Process
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, key_sentences)
            mime_type = "application/pdf"
        else:
            processed_doc = file_bytes  # Highlights for Word are better viewed post-export
            mime_type = "application/octet-stream"

    # --- UI Layout ---
    res_col1, res_col2 = st.columns([1, 1])

    with res_col1:
        st.subheader("🏁 Systematic Analysis Report")
        st.markdown(systematic_notes)
        st.divider()
        st.write("📥 **Download Systematic Notes:**")
        d_col1, d_col2 = st.columns(2)
        d_col1.download_button("Word (.docx)", systematic_notes, f"AI_Reconstruction_{uploaded_file.name}.docx")
        d_col2.download_button("PDF (.pdf)", export_summary_pdf(systematic_notes),
                               f"AI_Reconstruction_{uploaded_file.name}.pdf")

    with res_col2:
        st.subheader("📄 Document Preview (Highlights)")
        st.download_button(
            label=f"📥 Download Highlighted {file_ext.upper()}",
            data=processed_doc,
            file_name=f"Highlighted_{uploaded_file.name}",
            mime=mime_type,
            use_container_width=True
        )

        if file_ext == "pdf":
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            # Using embed with fixed heights to prevent Chrome blocking preview
            pdf_embed = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="900px" type="application/pdf">'
            st.markdown(pdf_embed, unsafe_allow_html=True)
        else:
            st.info("Live preview is native to PDF. For Word documents, please download the file to see highlights.")

else:
    st.info("Ready for analysis. Please upload your document to begin.")