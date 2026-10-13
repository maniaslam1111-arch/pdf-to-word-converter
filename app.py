import io
import os
import tempfile
import time
from pathlib import Path

import pymupdf
import streamlit as st
from pdf2docx import Converter

APP_NAME = "DocShift"
TAGLINE = "Seamless Document Conversion Across Any Device"
MAX_BYTES = 50 * 1024 * 1024
PAGE_ICON = "📄"

st.set_page_config(
    page_title=f"{APP_NAME} | PDF to Word",
    page_icon=PAGE_ICON,
    layout="wide",
    initial_sidebar_state="collapsed",
)

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Sora:wght@500;600;700&display=swap');

:root {
  --cyan: #6ee7ff;
  --violet: #8b7bff;
  --mint: #7dffc3;
  --glass: rgba(16, 28, 48, 0.55);
  --stroke: rgba(255, 255, 255, 0.12);
  --text: #e8f1ff;
  --muted: #9bb0c9;
}

html, body, [data-testid="stAppViewContainer"] {
  font-family: "Outfit", sans-serif;
  color: var(--text);
}

[data-testid="stAppViewContainer"] {
  background:
    radial-gradient(1200px 700px at 8% -10%, rgba(110, 231, 255, 0.18), transparent 55%),
    radial-gradient(900px 600px at 100% 0%, rgba(139, 123, 255, 0.22), transparent 50%),
    radial-gradient(700px 500px at 50% 110%, rgba(125, 255, 195, 0.12), transparent 45%),
    linear-gradient(160deg, #06101c 0%, #0a1628 45%, #07111f 100%);
}

[data-testid="stHeader"] { background: transparent; }
[data-testid="stToolbar"] { visibility: hidden; }
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stDecoration"] { display: none; }
[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"] { display: none !important; }
section.main { margin-left: 0 !important; }
.stApp { overflow-x: hidden; }

.block-container {
  padding-top: 1.4rem !important;
  padding-bottom: 3rem !important;
  padding-left: 1.1rem !important;
  padding-right: 1.1rem !important;
  max-width: 980px !important;
  overflow-x: hidden;
}

[data-testid="stHeaderActionElements"],
.stHeading a,
[data-testid="stMarkdownContainer"] h1 a,
[data-testid="stMarkdownContainer"] h4 a {
  display: none !important;
}

[data-testid="stHorizontalBlock"] {
  flex-wrap: wrap !important;
  gap: 1rem !important;
}

[data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
  flex: 1 1 300px !important;
  min-width: min(100%, 280px) !important;
}

div.stButton > button:disabled {
  opacity: 0.5;
  filter: grayscale(0.25);
  cursor: not-allowed;
  transform: none;
  box-shadow: none;
}

.hero {
  text-align: center;
  padding: 1.4rem 1rem 0.4rem;
}

.brand-mark {
  width: 64px;
  height: 64px;
  margin: 0 auto 0.85rem;
  border-radius: 18px;
  display: grid;
  place-items: center;
  font-family: "Sora", sans-serif;
  font-weight: 700;
  font-size: 1.35rem;
  letter-spacing: -0.04em;
  color: #041018;
  background: linear-gradient(135deg, #6ee7ff, #8b7bff 55%, #7dffc3);
  box-shadow: 0 10px 40px rgba(110, 231, 255, 0.28);
}

.brand-title {
  font-family: "Sora", sans-serif;
  font-size: clamp(2rem, 5vw, 3.4rem);
  font-weight: 700;
  letter-spacing: -0.04em;
  margin: 0;
  background: linear-gradient(90deg, #ffffff, #bfefff 40%, #c9c2ff);
  -webkit-background-clip: text;
  background-clip: text;
  color: transparent;
}

.brand-tagline {
  margin: 0.55rem auto 0;
  max-width: 540px;
  color: var(--muted);
  font-size: clamp(0.95rem, 2.2vw, 1.12rem);
  font-weight: 400;
}

.glass-card {
  background: var(--glass);
  border: 1px solid var(--stroke);
  border-radius: 22px;
  padding: 1.25rem 1.3rem;
  backdrop-filter: blur(18px);
  -webkit-backdrop-filter: blur(18px);
  box-shadow: 0 18px 50px rgba(0, 0, 0, 0.28);
}

.section-label {
  font-size: 0.78rem;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: var(--cyan);
  font-weight: 600;
  margin-bottom: 0.35rem;
}

.feature-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 0.9rem;
  margin: 0.4rem 0 1.2rem;
}

.feature-item {
  background: rgba(255,255,255,0.04);
  border: 1px solid var(--stroke);
  border-radius: 16px;
  padding: 0.95rem 1rem;
}

.feature-item h4 {
  margin: 0 0 0.3rem;
  font-family: "Sora", sans-serif;
  font-size: 0.95rem;
}

.feature-item p {
  margin: 0;
  color: var(--muted);
  font-size: 0.86rem;
  line-height: 1.45;
}

.meta-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
  gap: 0.8rem;
}

.meta-chip {
  background: rgba(110, 231, 255, 0.08);
  border: 1px solid rgba(110, 231, 255, 0.16);
  border-radius: 14px;
  padding: 0.85rem 0.9rem;
}

.meta-chip span {
  display: block;
  color: var(--muted);
  font-size: 0.72rem;
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.meta-chip strong {
  display: block;
  margin-top: 0.25rem;
  font-size: 0.98rem;
  word-break: break-word;
}

.status-ok { color: #7dffc3; }
.status-warn { color: #ffd27a; }
.status-err { color: #ff8b9a; }

[data-testid="stFileUploader"] {
  background: rgba(255,255,255,0.03);
  border: 1.5px dashed rgba(110, 231, 255, 0.35);
  border-radius: 18px;
  padding: 0.4rem 0.8rem 0.9rem;
}

[data-testid="stFileUploader"] section {
  background: transparent !important;
}

div.stButton > button {
  width: 100%;
  border: 0;
  border-radius: 14px;
  padding: 0.72rem 1rem;
  font-family: "Sora", sans-serif;
  font-weight: 600;
  letter-spacing: 0.01em;
  color: #041018;
  background: linear-gradient(135deg, #6ee7ff, #8b7bff);
  box-shadow: 0 10px 28px rgba(110, 231, 255, 0.22);
  transition: transform 0.15s ease, filter 0.15s ease;
}

div.stButton > button:hover {
  filter: brightness(1.06);
  transform: translateY(-1px);
}

div.stDownloadButton > button {
  width: 100%;
  border: 1px solid rgba(125, 255, 195, 0.45);
  border-radius: 14px;
  padding: 0.72rem 1rem;
  font-family: "Sora", sans-serif;
  font-weight: 600;
  color: #041018;
  background: linear-gradient(135deg, #7dffc3, #6ee7ff);
}

.footer-note {
  text-align: center;
  color: var(--muted);
  font-size: 0.82rem;
  margin-top: 1.6rem;
}

@media (max-width: 900px) {
  .block-container {
    padding-left: 0.85rem !important;
    padding-right: 0.85rem !important;
  }
}
</style>
"""


def format_size(num_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    size = float(num_bytes)
    for unit in units:
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{num_bytes} B"


def inspect_pdf(file_bytes: bytes) -> dict:
    try:
        doc = pymupdf.open(stream=file_bytes, filetype="pdf")
    except Exception as exc:
        raise ValueError("This file could not be opened as a PDF. It may be corrupted.") from exc

    try:
        if doc.is_encrypted:
            if not doc.authenticate(""):
                raise PermissionError(
                    "This PDF is password-protected. Remove the password and try again."
                )

        page_count = doc.page_count
        if page_count < 1:
            raise ValueError("This PDF has no readable pages.")

        return {
            "page_count": page_count,
            "metadata": doc.metadata or {},
        }
    finally:
        doc.close()


def convert_pdf_to_docx(file_bytes: bytes, progress_cb=None) -> bytes:
    pdf_path = None
    docx_path = None
    converter = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pdf_file:
            pdf_file.write(file_bytes)
            pdf_path = pdf_file.name

        fd, docx_path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)

        if progress_cb:
            progress_cb(15, "Preparing conversion engine…")

        converter = Converter(pdf_path)
        if progress_cb:
            progress_cb(45, "Extracting text, tables, and images…")

        converter.convert(docx_path, start=0, end=None)
        if progress_cb:
            progress_cb(85, "Assembling Word document…")

        output = Path(docx_path).read_bytes()
        if not output:
            raise RuntimeError("Conversion produced an empty Word file.")
        return output
    except PermissionError:
        raise
    except ValueError:
        raise
    except Exception as exc:
        message = str(exc).lower()
        if "password" in message or "encrypt" in message:
            raise PermissionError(
                "This PDF is password-protected. Remove the password and try again."
            ) from exc
        if "broken" in message or "corrupt" in message or "invalid" in message:
            raise ValueError("This PDF appears corrupted and could not be converted.") from exc
        raise RuntimeError(f"Conversion failed: {exc}") from exc
    finally:
        if converter is not None:
            try:
                converter.close()
            except Exception:
                pass
        for path in (pdf_path, docx_path):
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass


def init_state():
    defaults = {
        "docx_bytes": None,
        "docx_name": None,
        "error": None,
        "converted": False,
        "last_file_id": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_hero():
    st.markdown(
        f"""
        <div class="hero">
          <div class="brand-mark">DS</div>
          <h1 class="brand-title">{APP_NAME}</h1>
          <p class="brand-tagline">{TAGLINE}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_features():
    st.markdown(
        """
        <div class="feature-grid">
          <div class="feature-item">
            <h4>High-fidelity conversion</h4>
            <p>Preserves text styles, tables, and embedded images with pdf2docx.</p>
          </div>
          <div class="feature-item">
            <h4>Any screen size</h4>
            <p>Layout stacks cleanly on phones, tablets, laptops, and desktops.</p>
          </div>
          <div class="feature-item">
            <h4>Private by design</h4>
            <p>Files stay on this machine and temp copies are deleted after conversion.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main():
    init_state()
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    render_hero()
    render_features()

    left, right = st.columns([1.15, 0.85], gap="large")

    with left:
        st.markdown('<div class="section-label">Upload</div>', unsafe_allow_html=True)
        st.markdown("#### Drop your PDF here")
        uploaded = st.file_uploader(
            "Drag and drop a PDF, or browse from your device",
            type=["pdf"],
            accept_multiple_files=False,
            help="Maximum file size: 50 MB",
            label_visibility="collapsed",
        )

        if uploaded is None:
            st.session_state.docx_bytes = None
            st.session_state.converted = False
            st.session_state.error = None
            st.caption("Accepted format: PDF · Max size: 50 MB")
        else:
            file_id = f"{uploaded.name}-{uploaded.size}"
            if st.session_state.last_file_id != file_id:
                st.session_state.docx_bytes = None
                st.session_state.converted = False
                st.session_state.error = None
                st.session_state.last_file_id = file_id

    with right:
        st.markdown('<div class="section-label">Preview</div>', unsafe_allow_html=True)
        st.markdown("#### File details")

        if uploaded is None:
            st.info("Upload a PDF to see name, size, and page count.")
        elif uploaded.size > MAX_BYTES:
            st.error("File exceeds the 50 MB limit. Please choose a smaller PDF.")
        else:
            file_bytes = uploaded.getvalue()
            try:
                info = inspect_pdf(file_bytes)
                st.markdown(
                    f"""
                    <div class="glass-card">
                      <div class="meta-grid">
                        <div class="meta-chip">
                          <span>File name</span>
                          <strong>{uploaded.name}</strong>
                        </div>
                        <div class="meta-chip">
                          <span>Size</span>
                          <strong>{format_size(uploaded.size)}</strong>
                        </div>
                        <div class="meta-chip">
                          <span>Pages</span>
                          <strong>{info['page_count']}</strong>
                        </div>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            except PermissionError as exc:
                st.error(str(exc))
                uploaded = None
            except ValueError as exc:
                st.error(str(exc))
                uploaded = None
            except Exception:
                st.error("Unable to read this PDF. It may be damaged.")
                uploaded = None

    st.markdown("---")
    action_col, status_col = st.columns([0.42, 0.58], gap="large")

    convert_clicked = False
    with action_col:
        convert_clicked = st.button(
            "Convert to Word",
            type="primary",
            disabled=uploaded is None,
            use_container_width=True,
        )

    with status_col:
        status_slot = st.empty()
        progress_slot = st.empty()

        if uploaded is None:
            status_slot.caption("Waiting for a PDF upload.")
        elif st.session_state.converted:
            status_slot.success("Conversion complete. Your Word file is ready.")
        else:
            status_slot.info("Ready to convert. Formatting, tables, and images will be preserved.")

    if convert_clicked and uploaded is not None:
        file_bytes = uploaded.getvalue()
        status_slot.info("Starting conversion…")
        bar = progress_slot.progress(5, text="Reading PDF…")

        def on_progress(value: int, label: str):
            bar.progress(value, text=label)

        spinner_box = st.empty()
        with spinner_box.container():
            with st.spinner("Converting your document. This can take a few seconds…"):
                try:
                    time.sleep(0.15)
                    inspect_pdf(file_bytes)
                    docx_bytes = convert_pdf_to_docx(file_bytes, on_progress)
                    bar.progress(100, text="Done")
                    st.session_state.docx_bytes = docx_bytes
                    stem = Path(uploaded.name).stem or "document"
                    st.session_state.docx_name = f"{stem}.docx"
                    st.session_state.converted = True
                    st.session_state.error = None
                    status_slot.success("Conversion complete. Your Word file is ready.")
                except PermissionError as exc:
                    st.session_state.error = str(exc)
                    st.session_state.converted = False
                    status_slot.error(str(exc))
                    progress_slot.empty()
                except ValueError as exc:
                    st.session_state.error = str(exc)
                    st.session_state.converted = False
                    status_slot.error(str(exc))
                    progress_slot.empty()
                except Exception as exc:
                    st.session_state.error = f"Conversion failed: {exc}"
                    st.session_state.converted = False
                    status_slot.error(st.session_state.error)
                    progress_slot.empty()
        spinner_box.empty()

    if st.session_state.converted and st.session_state.docx_bytes:
        st.download_button(
            label="Download converted .docx",
            data=st.session_state.docx_bytes,
            file_name=st.session_state.docx_name or "converted.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    st.markdown(
        f'<p class="footer-note">{APP_NAME} keeps conversions local. Temporary files are removed as soon as conversion finishes.</p>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
