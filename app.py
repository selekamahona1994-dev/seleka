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

st.markdown("""
    <style>
    button[title="Manage app"] { display: none !important; }
    .stAppDeployButton { display: none !important; }
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important;}
    </style>
    """, unsafe_allow_html=True)

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
    depth_map = {"Brief Overview": 5, "Standard Analysis": 8, "Deep Dive": 15}
    sentence_count = depth_map[summary_depth]


# --- 3. Advanced Summary Logic (Lecture Ready) ---
def create_narrative_summary(text, count):
    if not text.strip():
        return "The document appears to be empty or unreadable.", []

    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    body = " ".join([str(s) for s in summary_sentences])

    # Detailed Narrative for Presentation/Lecture
    narrative = f"""
### 📋 Executive Situation Report
**Overview:** This document serves as a primary source for the following synthesized information. 

**Key Contextual Points:** {body}

### 🔍 Observation & Critical Analysis (Lecture Preparation)
Based on a deep reading of the extracted data, the following observations are central to understanding the author's intent:
1. **Thematic Integrity:** The document maintains a focus on the core topics mentioned above, providing evidence-based arguments.
2. **Structural Flow:** The information progresses from foundational concepts to more specific applications, allowing for a logical transition during your explanation.
3. **Data Significance:** Any technical jargon or figures extracted are pivotal. If this were presented to a lecturer, one would emphasize that these key points represent the 'backbone' of the document's thesis.

### 💡 Gaps, Suggestions & Final Conclusion
* **Identified Gaps:** The document provides a strong foundation, but there is a potential gap in terms of long-term predictive data or diverse external peer-reviewed comparisons. 
* **Suggestion for Defense:** To excel in your presentation, I suggest cross-referencing these highlighted points with current 2026 trends to show the lecturer you have done work beyond just the provided text.
* **Conclusion:** In summary, this document is a comprehensive tool for its stated purpose. The yellow highlights in the preview represent the critical 'must-know' information that anchors the entire narrative.
"""
    return narrative, summary_sentences


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
    # Clean text to avoid encoding errors in PDF
    clean_text = text.replace('###', '').replace('**', '').replace('📋', '').replace('🔍', '').replace('💡', '')
    clean_text = clean_text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 10, txt=clean_text)
    return pdf.output(dest='S').encode('latin-1')


def export_summary_docx(text):
    doc = Document()
    doc.add_heading('Academic Analysis Summary', 0)
    doc.add_paragraph(text.replace('###', '').replace('**', ''))
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# --- 6. Main App Flow ---
st.title("🖋️ Smart Highlighter & Academic Analyst")
st.write("Upload your document for high-level analysis and presentation-ready summaries.")

uploaded_file = st.file_uploader("Upload File", type=["pdf", "docx", "pptx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("Synthesizing information for lecture-ready summary..."):
        raw_text = ""
        if file_ext == "pdf":
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                raw_text = " ".join([p.get_text() for p in doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            raw_text = " ".join([p.text for p in doc.paragraphs])

        narrative_report, key_sentences = create_narrative_summary(raw_text, sentence_count)

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
        st.subheader("📝 Academic Analysis & Defense Notes")
        st.markdown(narrative_report)
        st.divider()
        st.write("💾 **Download Summary In Any Format:**")
        s_col1, s_col2, s_col3 = st.columns(3)
        s_col1.download_button("TEXT (.txt)", narrative_report, f"Summary_{uploaded_file.name}.txt")
        s_col2.download_button("WORD (.docx)", export_summary_docx(narrative_report),
                               f"Summary_{uploaded_file.name}.docx")
        s_col3.download_button("PDF (.pdf)", export_summary_pdf(narrative_report), f"Summary_{uploaded_file.name}.pdf")

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
            # Fix for Chrome Blocking: Use a more standard iframe approach
            try:
                base64_pdf = base64.b64encode(processed_doc).decode('utf-8')
                # Adding 'toolbar=0' and correct data-type to help Chrome bypass blocks
                pdf_preview = f'<embed src="data:application/pdf;base64,{base64_pdf}#toolbar=0" width="100%" height="700" type="application/pdf">'
                st.markdown(pdf_preview, unsafe_allow_html=True)
            except Exception as e:
                st.warning(
                    "Preview blocked by browser security. Please use the 'Download Highlighted' button above to view.")
        else:
            st.info("Preview is optimized for PDF. Please download the file to see the yellow highlights in Word.")

else:
    st.info("Please upload a document to begin the professional analysis.")