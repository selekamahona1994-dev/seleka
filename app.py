import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from pptx import Presentation
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
    /* Keep sidebar and menu visible */
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
        st.info("💡 Place 'logo.png' in your GitHub repo.")
    st.title("Settings")
    # Increased slider max to allow for a more descriptive summary
    summary_length = st.slider("Explanation Depth (sentences)", 3, 15, 7)

st.title("🖋️ Smart Document Highlighter & Analyst")
st.write("Upload your file. I will explain the **situation** and highlight key segments in **yellow**.")


# --- 3. Processing Logic ---

def get_meaningful_summary(text, count, file_type):
    if not text.strip():
        return "The document appears to be empty or contains no readable text."

    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary_sentences = summarizer(parser.document, count)

    # Building connectivity: Adding an introductory narrative line based on file type
    intro = f"Based on the analyzed {file_type.upper()} document, here is a brief explanation of the situation: \n\n"

    # Joining sentences to flow like a paragraph
    explanation = " ".join([str(s) for s in summary_sentences])

    # Ensure it's meaningful and cohesive
    full_narrative = f"{intro}{explanation}"
    return full_narrative, [str(s) for s in summary_sentences]


def highlight_pdf(file_bytes, keypoints):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        for point in keypoints:
            # Match shorter chunks to ensure the highlighter finds the text effectively
            text_instances = page.search_for(point[:50])
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

    with st.spinner("Analyzing document context..."):
        text_content = ""

        # Format-specific extraction
        if file_ext == "pdf":
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
            text_content = " ".join([p.get_text() for p in pdf_doc])
        elif file_ext == "docx":
            doc = Document(io.BytesIO(file_bytes))
            text_content = " ".join([p.text for p in doc.paragraphs])
        elif file_ext == "pptx":
            prs = Presentation(io.BytesIO(file_bytes))
            text_content = " ".join(
                [shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])

        # Generate the Meaningful Narrative Summary
        narrative_summary, key_sentences = get_meaningful_summary(text_content, summary_length, file_ext)

        # Process Highlighting
        if file_ext == "pdf":
            processed_doc = highlight_pdf(file_bytes, key_sentences)
            mime_type = "application/pdf"
        elif file_ext == "docx":
            processed_doc = highlight_docx(file_bytes, key_sentences)
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            # PPTX highlighting is complex locally; we provide the original for now
            processed_doc = file_bytes
            mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    # Display Results
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📑 Situation Summary")
        # Displaying the narrative in a clean box
        st.info(narrative_summary)
        st.download_button(
            "📥 Download Brief Summary (.txt)",
            narrative_summary,
            f"Summary_{uploaded_file.name}.txt"
        )

    with col2:
        st.subheader("📄 Keypoints Highlighted")
        st.success(f"Analysis of {uploaded_file.name} complete!")
        st.download_button(
            label=f"📥 Download Highlighted {file_ext.upper()}",
            data=processed_doc,
            file_name=f"Highlighted_{uploaded_file.name}",
            mime=mime_type
        )

        if file_ext == "pdf":
            st.write("---")
            st.caption("Preview available via Download button.")