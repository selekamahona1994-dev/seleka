import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from pptx import Presentation
from PIL import Image
import io
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
import nltk

# Download necessary data for summarization
nltk.download('punkt')

# 1. Page Configuration & UI Customization
st.set_page_config(page_title="AI Doc Highlighter", layout="wide")

# CSS to hide "Manage app" and the specific decoration line at the top
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: visible;}
            footer {visibility: hidden;}
            .stApp [data-testid="stToolbar"] {display: none;}
            /* This targets the 'Manage App' button specifically */
            button[title="Manage app"] {display: none !important;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# 2. Sidebar with Logo
with st.sidebar:
    try:
        st.image("logo.png", width=200)
    except:
        st.warning("logo.png not found. Please upload it to your repo.")
    st.title("Settings")
    st.info("Upload a document to automatically highlight key sentences and generate a summary.")

st.title("🖋️ Smart Document Highlighter & Summarizer")


# 3. Helper Functions
def summarize_text(text, sentences_count=5):
    """Generates a summary using LSA (Free, no API key)."""
    parser = PlaintextParser.from_string(text, Tokenizer("english"))
    summarizer = LsaSummarizer()
    summary = summarizer(parser.document, sentences_count)
    return " ".join([str(sentence) for sentence in summary])


def process_pdf(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    full_text = ""
    # Extract text for summary
    for page in doc:
        full_text += page.get_text()

    # Simple Logic: Highlight the first sentence of every paragraph as "Key Points"
    # In a real app, you'd use NLP to find 'important' sentences.
    summary_text = summarize_text(full_text, sentences_count=3)
    important_sentences = summary_text.split(". ")

    for page in doc:
        for sentence in important_sentences:
            text_instances = page.search_for(sentence[:50])  # Search first 50 chars
            for inst in text_instances:
                highlight = page.add_highlight_annot(inst)
                highlight.set_colors(stroke=(1, 1, 0))  # Yellow
                highlight.update()

    # Save to bytes
    out_pdf = io.BytesIO()
    doc.save(out_pdf)
    return out_pdf.getvalue(), full_text, summary_text


# 4. File Upload Logic
uploaded_file = st.file_uploader("Upload Document", type=["pdf", "docx", "pptx", "png", "jpg"])

if uploaded_file is not None:
    file_type = uploaded_file.type

    with st.status("Processing document...", expanded=True) as status:
        if "pdf" in file_type:
            highlighted_pdf, original_text, summary = process_pdf(uploaded_file)

            st.subheader("📄 Document Summary")
            st.write(summary)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Download Highlighted PDF",
                    data=highlighted_pdf,
                    file_name=f"highlighted_{uploaded_file.name}",
                    mime="application/pdf"
                )
            with col2:
                st.download_button(
                    label="Download Summary (.txt)",
                    data=summary,
                    file_name="summary.txt",
                    mime="text/plain"
                )

            # Preview (PDF only)
            st.divider()
            st.subheader("Preview")
            st.info("The highlighted version is ready for download above.")

        else:
            # Placeholder for other formats
            st.error(
                "Currently, automatic yellow highlighting is most stable for PDF. Word/PPT extraction is available, but visual highlighting is in progress.")
            # Simple text extraction for non-PDFs
            st.write("Extracting summary...")
            # (Logic for docx/pptx would go here similar to PDF)

else:
    st.info("Please upload a PDF to see the highlighting in action.")