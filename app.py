import os, shutil
from pathlib import Path

import io
import pathlib
import streamlit as st
import pandas as pd

HERE = pathlib.Path(__file__).parent

@st.cache_data(show_spinner=False)
def load_excel_df(src, *, sheet_name=0, header=0):
    """src can be an UploadedFile, bytes, path-like, or file-like."""
    if hasattr(src, "read"):           # UploadedFile / file-like
        return pd.read_excel(src, engine="openpyxl", sheet_name=sheet_name, header=header)
    src = pathlib.Path(src)
    with src.open("rb") as f:
        return pd.read_excel(f, engine="openpyxl", sheet_name=sheet_name, header=header)

def resolve_support_file(uploaded, fallback_filename):
    """Return a file-like object for pandas: uploaded if present, else repo file."""
    if uploaded is not None:
        # Make an in-memory copy so cached readers can re-use it
        return io.BytesIO(uploaded.getbuffer())
    # Fall back to file shipped in the repo
    fallback_path = HERE / fallback_filename
    if not fallback_path.exists():
        raise FileNotFoundError(f"Missing required support file: {fallback_filename}")
    return fallback_path



















REQUIRED_SUPPORT = ["Parts.xlsx", "revenue_type.xlsx", "lookup_gl.xlsx"]

def find_support_dir(team_subpath: str) -> Path | None:
    """
    Try to locate a OneDrive or SharePoint 'Team Sites' folder for the current user,
    then append team_subpath. Returns a Path or None.
    """
    candidates: list[Path] = []

    # Env vars set by OneDrive
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.environ.get(var)
        if p:
            candidates += [Path(p), Path(p) / team_subpath]

    # Common roots under the current profile (handles 'OneDrive - Company' and 'Team Sites')
    up = Path(os.environ.get("USERPROFILE", str(Path.home())))
    candidates += [p for p in up.glob("OneDrive*")]
    candidates += [p for p in up.glob("*Team Sites*/*")]

    # De-dup
    seen = set(); uniq = []
    for c in candidates:
        if c not in seen:
            seen.add(c); uniq.append(c)

    # Test each candidate + subpath
    for root in uniq:
        probe = root / team_subpath if not str(root).endswith(team_subpath) else root
        if probe.exists() and all((probe / f).exists() for f in REQUIRED_SUPPORT):
            return probe

    return None




import streamlit as st
import subprocess, sys, os, shutil
import pandas as pd
from datetime import datetime
from pathlib import Path

st.title("Merged Revenue Analysis")

choice = st.sidebar.radio(
    "Choose which notebook code to run:",
    ["NC 100% Revenue", "NC Partial Revenue"]
)

KEEP_COLS = [
    "project manager","account","install id (task)","sales order number",
    "total % completed","total so extended $ amount","total rec'd"
]

def timestamped_filename(base_name: str) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    name, ext = os.path.splitext(base_name)
    return f"{name}_{ts}{ext}"

if choice == "NC 100% Revenue":
    st.header("NC 100% Revenue (Cloud-safe)")

    # Required uploads
    ab_file = st.file_uploader("Upload AB File (ab.xlsx)", type=["xlsx"], key="abfile")
    ml_file = st.file_uploader("Upload Mavenlink File (ML.xlsx)", type=["xlsx"], key="mlfile")
    cost_file = st.file_uploader("Upload Costs File (Costs.xlsx)", type=["xlsx"], key="costfile")

    st.markdown("**Optional Support Files (if not bundled in repo):**")
    parts_up   = st.file_uploader("Parts.xlsx (optional)", type=["xlsx"], key="parts100")
    revtype_up = st.file_uploader("revenue_type.xlsx (optional)", type=["xlsx"], key="revtype100")
    gl_up      = st.file_uploader("lookup_gl.xlsx (optional)", type=["xlsx"], key="gl100")

    # Require core uploads
    if not (ab_file and ml_file and cost_file):
        st.info("Please upload AB, ML and Costs files to enable processing.")
        st.stop()

    # Resolve support files
    try:
        parts_src   = resolve_support_file(parts_up,   "Parts.xlsx")
        revtype_src = resolve_support_file(revtype_up, "revenue_type.xlsx")
        gl_src      = resolve_support_file(gl_up,      "lookup_gl.xlsx")
    except FileNotFoundError as e:
        st.error(f"{e}\n\nAdd the file to your GitHub repo (same folder as app.py) or upload it here.")
        st.stop()

    # Quick preview of ML sheet (Summary tab)
    try:
        df_ml = load_excel_df(ml_file, sheet_name="Summary", header=2)
        st.write("ML columns detected:", list(df_ml.columns)[:12], "…")
    except Exception as e:
        st.warning(f"Could not preview ML file: {e}")

    import os, sys, subprocess, shutil, tempfile, datetime

    def timestamped_filename(base):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        stem, ext = os.path.splitext(base)
        return f"{stem}_{ts}{ext}"

    if st.button("Run 100% Revenue Process"):
        with st.spinner("Preparing inputs …"):
            workdir = tempfile.mkdtemp(prefix="full_")

            # Save core inputs with expected names
            for src, name in [(ab_file,"ab.xlsx"),(ml_file,"ML.xlsx"),(cost_file,"Costs.xlsx")]:
                with open(os.path.join(workdir, name), "wb") as f:
                    f.write(src.getbuffer())

            # Save/Copy support files
            def _dump(src, dstname):
                dst = os.path.join(workdir, dstname)
                if isinstance(src, (pathlib.Path, str)):
                    shutil.copy2(src, dst)
                else:
                    with open(dst, "wb") as f:
                        f.write(src.getbuffer() if hasattr(src, "getbuffer") else src.read())
            _dump(parts_src,   "Parts.xlsx")
            _dump(revtype_src, "revenue_type.xlsx")
            _dump(gl_src,      "lookup_gl.xlsx")

        with st.spinner("Running nc100.py …"):
            script_path = str(HERE / "nc100.py")
            result = subprocess.run(
                [sys.executable, script_path],
                cwd=workdir, capture_output=True, text=True
            )

        st.subheader("Process logs")
        st.code(result.stdout or "(no stdout)")
        if result.stderr:
            st.error(result.stderr)

        produced = []
        for name in ("Updated Revenue Report.xlsx",):
            p = os.path.join(workdir, name)
            if os.path.exists(p):
                produced.append(p)

        if not produced:
            st.warning("No output file produced — check the logs above.")
        else:
            for p in produced:
                ts_name = timestamped_filename(os.path.basename(p))
                new_p = os.path.join(workdir, ts_name)
                os.replace(p, new_p)
                with open(new_p, "rb") as f:
                    st.download_button(
                        label=f"Download {ts_name}",
                        data=f.read(),
                        file_name=ts_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )



elif choice == "NC Partial Revenue":
    st.header("NC Partial Revenue (Cloud-safe)")

    # Only the two operational uploads are required for Partial
    pb_partial = st.file_uploader("Power BI Partial (powerbipartial.xlsx)", type=["xlsx"], key="pbpartial")
    ml_partial = st.file_uploader("Mavenlink Partial (MLPARTIAL.xlsx)", type=["xlsx"], key="mlpartial")

    st.markdown("**Optional (only if not bundled in repo):** Upload support files or the app will use the copies in the repo.")
    parts_up   = st.file_uploader("Parts.xlsx (optional)", type=["xlsx"], key="parts")
    revtype_up = st.file_uploader("revenue_type.xlsx (optional)", type=["xlsx"], key="revtype")
    gl_up      = st.file_uploader("lookup_gl.xlsx (optional)", type=["xlsx"], key="gl")

    # Gate: need the two operational files
    if not (pb_partial and ml_partial):
        st.info("Upload both Power BI Partial and MLPARTIAL to enable processing.")
        st.stop()

    # --- Resolve support files (upload > fallback to repo) ---
    try:
        parts_src   = resolve_support_file(parts_up,   "Parts.xlsx")
        revtype_src = resolve_support_file(revtype_up, "revenue_type.xlsx")
        gl_src      = resolve_support_file(gl_up,      "lookup_gl.xlsx")
    except FileNotFoundError as e:
        st.error(f"{e}\n\nAdd the file to your GitHub repo (same folder as app.py) or upload it here.")
        st.stop()

    # --- Preview (optional) ---
    with st.expander("Preview detected inputs", expanded=False):
        try:
            st.write("Power BI columns:", list(load_excel_df(pb_partial).columns)[:12], "…")
        except Exception as e:
            st.warning(f"Could not preview Power BI file: {e}")
        try:
            st.write("ML columns:", list(load_excel_df(ml_partial, sheet_name="Summary", header=2).columns)[:12], "…")
        except Exception as e:
            st.warning(f"Could not preview ML file: {e}")

    # --- Run ncpartial.py using local temp files in the cloud container ---
    import os, sys, subprocess, shutil, tempfile, datetime

    def timestamped_filename(base):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        stem, ext = os.path.splitext(base)
        return f"{stem}_{ts}{ext}"

    if st.button("Run Partial Revenue Process"):
        with st.spinner("Preparing inputs…"):
            # Work in a temp dir to avoid collisions between sessions
            workdir = tempfile.mkdtemp(prefix="partial_")
            # Save the two required inputs with the exact names ncpartial.py expects
            with open(os.path.join(workdir, "powerbipartial.xlsx"), "wb") as f:
                f.write(pb_partial.getbuffer())
            with open(os.path.join(workdir, "MLPARTIAL.xlsx"), "wb") as f:
                f.write(ml_partial.getbuffer())
            # Save/Copy the support files with expected names
            def _dump(src, dstname):
                dst = os.path.join(workdir, dstname)
                if isinstance(src, (pathlib.Path, str)):
                    shutil.copy2(src, dst)
                else:
                    with open(dst, "wb") as f:
                        f.write(src.getbuffer() if hasattr(src, "getbuffer") else src.read())
                return dst
            _dump(parts_src,   "Parts.xlsx")
            _dump(revtype_src, "revenue_type.xlsx")
            _dump(gl_src,      "lookup_gl.xlsx")

        with st.spinner("Running ncpartial.py… this may take a moment"):
            # Run ncpartial.py located next to app.py; set cwd to the temp workdir
            script_path = str(HERE / "ncpartial.py")
            result = subprocess.run(
                [sys.executable, script_path],
                cwd=workdir, capture_output=True, text=True
            )

        st.subheader("Process logs")
        st.code(result.stdout or "(no stdout)")
        if result.stderr:
            st.error(result.stderr)

        # Offer any outputs the script produced
        produced = []
        for name in ("partialrevenue.xlsx", "Updated_Partial_Revenue.xlsx"):
            p = os.path.join(workdir, name)
            if os.path.exists(p):
                produced.append(p)

        if not produced:
            st.warning("No output files were produced. Check the logs above for errors.")
        else:
            for p in produced:
                ts_name = timestamped_filename(os.path.basename(p))
                # Rename within temp dir (purely cosmetic)
                new_p = os.path.join(workdir, ts_name)
                os.replace(p, new_p)
                with open(new_p, "rb") as f:
                    st.download_button(
                        label=f"Download {ts_name}",
                        data=f.read(),
                        file_name=ts_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
