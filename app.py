import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
import io
import os
import nltk
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer


# --- 1. Fix NLTK LookupError ---
def ensure_nltk_resources():
    resources = ['punkt', 'punkt_tab']
    for res in resources:
        try:
            nltk.data.find(f'tokenizers/{res}')
        except LookupError:
            nltk.download(res)


ensure_nltk_resources()

# --- 2. Page Configuration & UI ---
st.set_page_config(page_title="AI Document Highlighter", layout="wide")

# CSS to hide ONLY the 'Manage app' button and the top decoration
st.markdown("""
    <style>
    /* Hide the Manage App button */
    button[title="Manage app"] {
        display: none !important;
    }
    /* Keep sidebar and menu visible but clean up top bar */
    .stAppDeployButton {
        display: none !important;
    }
    footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# Sidebar with Logo
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.info("💡 Place 'logo.png' in your GitHub repo to see it here.")
    st.title("Settings")
    summary_length = st.slider("Summary Length (sentences)", 3, 10, 5)

st.title("🖋️ Smart Document Highlighter")
st.write("Upload a document. I will highlight key points in **yellow** and summarize the content.")


# --- 3. Processing Logic ---

def get_summary(text, count):
    if not text.strip(): return "No text found."
    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary = summarizer(parser.document, count)
    return [str(s) for s in summary]


def highlight_pdf(file_bytes, keypoints):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for point in keypoints:
            # We search for the first 60 chars to ensure a match
            text_instances = page.search_for(point[:60])
            for inst in text_instances:
                annot = page.add_highlight_annot(inst)
                annot.set_colors(stroke=(1, 1, 0))  # Yellow
                annot.update()

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def highlight_docx(file_bytes, keypoints):
    doc = Document(io.BytesIO(file_bytes))
    for para in doc.paragraphs:
        for point in keypoints:
            if point[:30] in para.text:
                for run in para.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# --- 4. Main App Flow ---
uploaded_file = st.file_uploader("Upload PDF, DOCX, or PPTX", type=["pdf", "docx", "pptx"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    file_ext = uploaded_file.name.split(".")[-1].lower()

    with st.spinner("Analyzing and Highlighting..."):
        # Extract Text for Summary
        text_content = ""
        if file_ext == "pdf":
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
            text_content = chr(12).join([p.get_text() for p in pdf_doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            text_content = "\n".join([p.text for p in doc.paragraphs])

        # Generate Summary & Keypoints
        keypoints = get_summary(text_content, summary_length)
        summary_text = " ".join(keypoints)

        # Process Document
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, keypoints)
            mime_type = "application/pdf"
        elif file_ext == "docx":
            processed_doc = highlight_docx(file_bytes, keypoints)
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            processed_doc = file_bytes  # Fallback
            mime_type = "application/octet-stream"

    # Display Results
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("✅ Summary")
        st.info(summary_text)
        st.download_button("📥 Download Summary (.txt)", summary_text, f"Summary_{uploaded_file.name}.txt")

    with col2:
        st.subheader("📄 Processed Document")
        st.success("Highlighting complete!")
        st.download_button(
            label=f"📥 Download Highlighted {file_ext.upper()}",
            data=processed_doc,
            file_name=f"Highlighted_{uploaded_file.name}",
            mime=mime_type
        )

        if file_ext == "pdf":
            st.write("---")
            st.write("Preview (First Page):")
            # PDF preview is possible in Streamlit
            st.download_button("Open Preview", processed_doc, "preview.pdf")