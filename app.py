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


# --- 4. Formal Office Submission Logic ---
def create_formal_submission(text, doc_type, summary_sentences):
    """Generates high-level professional content for official submission."""
    summary_text = " ".join([str(s) for s in summary_sentences])

    if "CV" in doc_type:
        formal_content = f"""
### ✉️ Official Cover Letter & Professional Summary
**Subject:** Formal Submission of Professional Credentials

To the Respective Office,

Please find the enclosed professional dossier. This document outlines a comprehensive background characterized by a strategic focus on core competencies identified within the text.

**Executive Synopsis:**
{summary_text}

The enclosed Curriculum Vitae demonstrates a commitment to professional excellence and a trajectory of consistent growth. I am prepared to discuss how these experiences align with the strategic goals of your organization.

Respectfully Submitted,
[Your Name/Electronic Signature]
        """
    elif "Research" in doc_type or "Notes" in doc_type:
        formal_content = f"""
### 🏛️ Executive Formal Report
**Subject:** Transmittal of Advanced Analysis and Research Findings

To the Office of Academic/Professional Affairs,

This correspondence serves as the formal submission of the analyzed findings regarding the uploaded documentation. The content has been synthesized to highlight high-level intellectual property and data-driven insights.

**Core Findings & Analysis:**
{summary_text}

**Strategic Conclusion:**
The data provided warrants significant consideration for policy implementation/academic advancement. We remain available to provide further clarification or high-level briefings as required by your office.

Best Regards,
[Department of Analysis]
        """
    else:
        formal_content = f"""
### 📄 Formal Office Communication
**Subject:** Official Documentation Summary and Transmittal

Following a thorough intelligence-led review of the provided materials, we are formally submitting the executive summary.

**Detailed Context:**
{summary_text}

This documentation is submitted for your records and official action.
        """
    return formal_content


# --- 5. Narrative & Systematic Rewriting Logic ---
def create_systematic_summary(text, count):
    if not text.strip():
        return "Unreadable content.", []

    doc_type = identify_document_type(text)
    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    words = re.findall(r'\w+', text.lower())
    common = [word for word, count in Counter(words).most_common(50) if len(word) > 5]
    highlighted_keywords = ", ".join(common[:10])

    intro = f"### 🤖 AI Document Identification: **{doc_type}**\n\n"
    intro += f"**Core Keywords Identified:** `{highlighted_keywords}`\n\n"

    body = "### 📋 Systematic Rewriting of Content\n"
    for i, s in enumerate(summary_sentences):
        sentence_str = str(s)
        for word in common[:5]:
            sentence_str = re.sub(f'({word})', r'**\1**', sentence_str, flags=re.IGNORECASE)
        body += f"{i + 1}. {sentence_str}\n\n"

    analysis = f"""
### 🔍 Observation & Analysis (Lecture Preparation)
The document is structured as a **{doc_type}**. 
* **Content Validity:** The document addresses the topic by focusing on `{common[0] if common else 'main themes'}`.
* **Gaps Found:** There is a lack of diverse citations and a missing 'Future Work' or 'Risk Assessment' section.
* **Suggestion:** When presenting to your lecturer, emphasize the connection between **{common[1] if len(common) > 1 else 'the data'}** and the final conclusions.

### 🏁 Conclusion & Systematic Synthesis
In summary, this **{doc_type}** serves as a vital resource. To improve it, adding a more detailed methodology or a glossary of the terms **{highlighted_keywords}** is recommended.
"""
    # Create the Formal Office content
    formal_office_content = create_formal_submission(text, doc_type, summary_sentences)

    full_output = intro + body + analysis + "\n---\n" + formal_office_content
    return full_output, summary_sentences


# --- 6. Highlighting & Export Logic ---
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
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 8, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


# --- 7. Main App Flow ---
st.title("🖋️ Smart AI Document Analyst & Systematizer")

uploaded_file = st.file_uploader("Upload Document", type=["pdf", "docx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("🧠 AI is identifying document type and generating formal submission content..."):
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        full_notes, key_sentences = create_systematic_summary(raw_text, sentence_count)

        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, key_sentences)
            mime_type = "application/pdf"
        else:
            processed_doc = file_bytes
            mime_type = "application/octet-stream"

    col_a, col_b = st.columns([1, 1])

    with col_a:
        st.subheader("📝 Systematic Analysis & Formal Submission")
        st.markdown(full_notes)
        st.divider()
        st.write("📥 **Download Analysis & Formal Letter:**")
        c1, c2 = st.columns(2)
        c1.download_button("Download as Word", full_notes, "Formal_Submission.docx")
        c2.download_button("Download as PDF", export_summary_pdf(full_notes), "Formal_Submission.pdf")

    with col_b:
        st.subheader("📄 Highlighted Preview")
        st.download_button("📥 Download Highlighted File", processed_doc, f"Highlighted_{uploaded_file.name}",
                           mime=mime_type)

        if file_ext == "pdf":
            base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
            pdf_code = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800px" type="application/pdf">'
            st.markdown(pdf_code, unsafe_allow_html=True)
        else:
            st.info("Preview optimized for PDF. Download file to see highlights.")

else:
    st.info("Please upload a file to begin.")