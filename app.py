import base64
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

import streamlit as st
import streamlit.components.v1 as components

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MAIN_PY = os.path.join(BASE_DIR, "main.py")
TEMPLATE_PDF = os.path.join(BASE_DIR, "temp.pdf")
OUTPUT_NAME = "All_Students_Report.pdf"  # name main.py writes


def expected_excel_name():
    """Read the excel filename main.py loads, so main.py never needs editing."""
    with open(MAIN_PY, encoding="utf-8", errors="ignore") as f:
        m = re.search(r'pd\.read_excel\(\s*["\'](.+?)["\']', f.read())
    return m.group(1) if m else "data.xlsx"


def sample_sheets():
    files = sorted(glob.glob(os.path.join(BASE_DIR, "*.xlsx")))
    return [f for f in files if not os.path.basename(f).startswith("~$")]


def generate_report(excel_bytes):
    """Run the unchanged main.py in an isolated temp folder. Returns (pdf_bytes, log)."""
    with tempfile.TemporaryDirectory() as tmp:
        shutil.copy(MAIN_PY, os.path.join(tmp, "main.py"))
        shutil.copy(TEMPLATE_PDF, os.path.join(tmp, "temp.pdf"))
        with open(os.path.join(tmp, expected_excel_name()), "wb") as f:
            f.write(excel_bytes)

        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        proc = subprocess.run(
            [sys.executable, "main.py"],
            cwd=tmp, capture_output=True, text=True, encoding="utf-8",
            errors="replace", env=env,
        )
        log = proc.stdout + ("\n" + proc.stderr if proc.stderr else "")
        out = os.path.join(tmp, OUTPUT_NAME)
        if proc.returncode != 0 or not os.path.exists(out):
            return None, log
        with open(out, "rb") as f:
            return f.read(), log


def auto_download(data, filename, mime):
    b64 = base64.b64encode(data).decode()
    components.html(
        f"""<a id="dl" href="data:{mime};base64,{b64}" download="{filename}"></a>
        <script>document.getElementById('dl').click();</script>""",
        height=0,
    )


st.set_page_config(page_title="Report Card Generator", page_icon="📄")
st.title("📄 Report Card Generator")
st.caption("Upload the marks Excel sheet and get the combined PDF report.")

with st.expander("⬇️ Download example Excel sheets"):
    sheets = sample_sheets()
    if not sheets:
        st.info("No example sheets found next to app.py.")
    for path in sheets:
        name = os.path.basename(path)
        with open(path, "rb") as f:
            st.download_button(
                f"Download {name}", f.read(), file_name=name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"sample_{name}",
            )

uploaded = st.file_uploader("Upload marks Excel sheet", type=["xlsx"])

if uploaded and st.button("Generate report", type="primary"):
    with st.spinner("Generating report cards... this can take a while for big sheets"):
        pdf, log = generate_report(uploaded.getvalue())
    if pdf is None:
        st.error("Generation failed. Check the sheet's columns/format. Log below.")
        st.code(log or "(no output)")
    else:
        st.session_state["pdf"] = pdf
        st.session_state["log"] = log
        st.session_state["autodl"] = True

if "pdf" in st.session_state:
    st.success("Report ready!")
    if st.session_state.pop("autodl", False):
        auto_download(st.session_state["pdf"], OUTPUT_NAME, "application/pdf")
    st.download_button(
        "⬇️ Download combined report (PDF)", st.session_state["pdf"],
        file_name=OUTPUT_NAME, mime="application/pdf",
    )
    with st.expander("Generation log"):
        st.code(st.session_state["log"])
