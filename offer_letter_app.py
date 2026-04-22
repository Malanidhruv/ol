import streamlit as st
import zipfile
import io
import os
import subprocess
import tempfile
from datetime import date, timedelta

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Harion Research – Offer Letter Generator",
    page_icon="📄",
    layout="centered",
)

# ── Branding / CSS ────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background: #f7f9fc; }
    .main-card {
        background: white;
        border-radius: 14px;
        padding: 2.5rem 2.5rem 2rem;
        box-shadow: 0 2px 18px rgba(0,0,0,0.07);
        margin-top: 1.5rem;
    }
    .brand-header {
        display:flex; align-items:center; gap:14px;
        border-bottom: 2px solid #1a3a5c; padding-bottom:1rem; margin-bottom:1.8rem;
    }
    .brand-title { font-size:1.55rem; font-weight:700; color:#1a3a5c; margin:0; }
    .brand-sub   { font-size:.85rem; color:#5a7fa8; margin:0; }
    .section-label { font-size:.78rem; font-weight:600; color:#5a7fa8;
                     text-transform:uppercase; letter-spacing:.08em; margin-bottom:.3rem; }
    div[data-testid="stButton"] button {
        background:#1a3a5c; color:white; border:none;
        border-radius:8px; padding:.65rem 2rem; font-size:1rem; font-weight:600;
        width:100%; margin-top:.5rem; cursor:pointer;
    }
    div[data-testid="stDownloadButton"] button {
        background:#0e6e3e; color:white; border:none;
        border-radius:8px; padding:.65rem 2rem; font-size:1rem; font-weight:600;
        width:100%; margin-top:.5rem;
    }
    .preview-box {
        background:#f0f4f8; border-radius:10px; padding:1.5rem 2rem;
        border-left: 4px solid #1a3a5c; margin-top:1.2rem;
        font-size:.92rem; line-height:1.7; color:#2d2d2d;
    }
    .success-badge {
        background:#e6f4ee; color:#0e6e3e; border-radius:6px;
        padding:.45rem 1rem; font-weight:600; font-size:.9rem;
        display:inline-block; margin-bottom:1rem;
    }
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "offer_letter_temp.docx")


# ── Helpers ───────────────────────────────────────────────────────────────────

def fill_docx_template(name: str, start_date: str) -> bytes:
    """Replace {} placeholders in the docx XML and return new docx bytes."""
    with open(TEMPLATE_PATH, "rb") as f:
        docx_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    xml = files["word/document.xml"].decode("utf-8")
    first  = xml.index("{}")
    xml    = xml[:first] + name + xml[first + 2:]
    second = xml.index("{}")
    xml    = xml[:second] + start_date + xml[second + 2:]
    files["word/document.xml"] = xml.encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in files.items():
            zout.writestr(n, data)
    return out.getvalue()


def docx_to_pdf(docx_bytes: bytes) -> bytes | None:
    """
    Convert docx → PDF via LibreOffice.
    Sets HOME to a writable temp dir — required on Streamlit Cloud where
    the default HOME is read-only and LibreOffice needs to write its profile.
    """
    with tempfile.TemporaryDirectory() as tmp:
        # Give LibreOffice a writable home for its user profile
        lo_home = os.path.join(tmp, "lo_home")
        os.makedirs(lo_home, exist_ok=True)

        docx_path = os.path.join(tmp, "offer_letter.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        env = os.environ.copy()
        env["HOME"]   = lo_home   # <-- the key fix for Streamlit Cloud
        env["TMPDIR"] = tmp

        for cmd in [
            ["soffice",     "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
        ]:
            try:
                subprocess.run(cmd, capture_output=True, text=True,
                               timeout=120, env=env)
                pdf_path = docx_path.replace(".docx", ".pdf")
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        return f.read()
            except (FileNotFoundError, subprocess.TimeoutExpired):
                continue
    return None


def ordinal(n: int) -> str:
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(
        n % 10 if n % 100 not in (11, 12, 13) else 0, "th"
    )
    return f"{n}{suffix}"


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="main-card">
  <div class="brand-header">
    <div>
      <p class="brand-title">📄 Harion Research</p>
      <p class="brand-sub">Internship Offer Letter Generator · HR Portal</p>
    </div>
  </div>
""", unsafe_allow_html=True)

st.markdown("Fill in the candidate details below to generate a personalised offer letter as a PDF.")

col1, col2 = st.columns([3, 2])

with col1:
    st.markdown('<p class="section-label">Candidate Full Name</p>', unsafe_allow_html=True)
    candidate_name = st.text_input(
        "Candidate Full Name",
        placeholder="e.g. Ananya Sharma",
        label_visibility="collapsed"
    )

with col2:
    st.markdown('<p class="section-label">Internship Start Date</p>', unsafe_allow_html=True)
    start_date = st.date_input(
        "Start Date",
        value=date.today() + timedelta(days=7),
        min_value=date.today(),
        label_visibility="collapsed",
        format="DD/MM/YYYY"
    )

formatted_date = f"{ordinal(start_date.day)} {start_date.strftime('%B %Y')}"

# Live preview
if candidate_name.strip():
    st.markdown(f"""
    <div class="preview-box"><strong>Preview snippet:</strong><br><br>
Dear <strong>{candidate_name}</strong>,<br><br>
On behalf of <strong>Harion Research</strong>, I am pleased to offer you the position of
<strong>Equity Research Analyst Intern</strong> for a duration of <strong>2 months</strong>,
starting from <strong>{formatted_date}</strong>.
    </div>""", unsafe_allow_html=True)
else:
    st.info("👆 Enter the candidate's name above to see a live preview.")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("⚡ Generate Offer Letter PDF", type="primary"):
    if not candidate_name.strip():
        st.error("Please enter the candidate's name before generating.")
    else:
        with st.spinner("Filling template and converting to PDF…"):
            docx_bytes = fill_docx_template(candidate_name.strip(), formatted_date)
            pdf_bytes  = docx_to_pdf(docx_bytes)

        safe_name = candidate_name.strip().replace(" ", "_")

        if pdf_bytes:
            st.markdown('<p class="success-badge">✅ Offer letter ready for download!</p>',
                        unsafe_allow_html=True)
            st.download_button(
                label="⬇️  Download PDF",
                data=pdf_bytes,
                file_name=f"Offer_Letter_{safe_name}.pdf",
                mime="application/pdf",
            )
        else:
            st.warning("PDF conversion failed. Downloading as Word document instead.")
            st.download_button(
                label="⬇️  Download DOCX",
                data=docx_bytes,
                file_name=f"Offer_Letter_{safe_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

st.markdown("</div>", unsafe_allow_html=True)

st.markdown("---")
st.markdown(
    "<center style='color:#aaa;font-size:.78rem;'>© 2025 Harion Research · HR Internal Tool</center>",
    unsafe_allow_html=True
)
