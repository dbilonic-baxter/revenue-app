import os, shutil
from pathlib import Path

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

# =========================
# NC 100% Revenue (UNCHANGED)
# =========================
if choice == "NC 100% Revenue":
    ab_file = st.file_uploader("Upload AB File (ab.xlsx)", type=["xlsx"], key="ab100")
    ml_file = st.file_uploader("Upload ML File (ML.xlsx)", type=["xlsx"], key="ml100")
    costs_file = st.file_uploader("Upload Costs File (Costs.xlsx)", type=["xlsx"], key="costs100")

    if ab_file and ml_file and costs_file:
        with open("ab.xlsx", "wb") as f: f.write(ab_file.getbuffer())

        df_ml = pd.read_excel(ml_file, sheet_name="Summary", header=2)
        df_ml.columns = df_ml.columns.str.strip().str.lower()
        st.write("📋 ML columns detected:", list(df_ml.columns))

        special_action_col = next((c for c in df_ml.columns if "special" in c and "action" in c), None)
        if special_action_col is None:
            st.error("❌ Could not find a 'Special Action' column in the ML file.")
        else:
            df_ml[special_action_col] = df_ml[special_action_col].astype(str).str.strip()
            cols_to_keep = KEEP_COLS + [special_action_col]
            df_ml = df_ml[[c for c in cols_to_keep if c in df_ml.columns]]
            df_ml = df_ml[df_ml[special_action_col].str.lower() == "full rec"]
            st.subheader("Filtered ML Preview (first 20 rows)")
            st.dataframe(df_ml.head(20))
            df_ml.to_excel("ML.xlsx", index=False)

        with open("Costs.xlsx", "wb") as f: f.write(costs_file.getbuffer())

        st.header("NC 100% Revenue")
        if st.button("Run 100% Revenue Process"):
            with st.spinner("Running NC 100% Revenue... please wait"):
                result = subprocess.run([sys.executable, "nc100.py"], capture_output=True, text=True)
            st.code(result.stdout)
            if result.stderr:
                st.error(result.stderr)
            else:
                original_file = "Updated Revenue Report.xlsx"
                if os.path.exists(original_file):
                    ts_file = timestamped_filename(original_file)
                    os.rename(original_file, ts_file)
                    st.success(f"✅ Process completed! File saved: {ts_file}")
                    with open(ts_file, "rb") as f:
                        st.download_button("Download Final Report", f, file_name=ts_file,
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Please upload AB, ML, and Costs files to enable processing.")

# =========================
# NC Partial Revenue (INDEPENDENT, ONLY 2 UPLOADS)
# =========================
elif choice == "NC Partial Revenue":
    st.header("NC Partial Revenue (Independent – 2 uploads)")

    # Upload only these two
    pb_partial = st.file_uploader("Power BI Partial (powerbipartial.xlsx)", type=["xlsx"], key="pbpartial")
    ml_partial = st.file_uploader("Mavenlink Partial (MLPARTIAL.xlsx)", type=["xlsx"], key="mlpartial")

    # Your shared subpath *relative to* each user’s OneDrive/Team Sites root
    TEAM_SUBPATH = r"Documents\HRC- CC-Revenue\General\Month End Close\Nurse Call\python_app"
    support_dir = find_support_dir(TEAM_SUBPATH)

    # Manual override if auto-detect fails or you want to point elsewhere
    manual = st.text_input(
        "Support files folder (auto-detected if possible)",
        value=str(support_dir) if support_dir else ""
    ).strip()
    if manual:
        p = Path(manual)
        if p.exists():
            support_dir = p

    # Save the two uploads to the exact filenames ncpartial.py expects
    ready_uploads = True
    if pb_partial:
        with open("powerbipartial.xlsx", "wb") as f: f.write(pb_partial.getbuffer())
    else:
        ready_uploads = False
    if ml_partial:
        with open("MLPARTIAL.xlsx", "wb") as f: f.write(ml_partial.getbuffer())
    else:
        ready_uploads = False

    # Copy support files from the detected/override folder
    missing_support = []
    if support_dir and support_dir.exists():
        for fname in REQUIRED_SUPPORT:
            src = support_dir / fname
            dst = Path.cwd() / fname
            if not src.exists():
                missing_support.append(fname); continue
            try:
                if (not dst.exists()) or (src.stat().st_mtime > dst.stat().st_mtime):
                    shutil.copy2(src, dst)
            except Exception as e:
                st.error(f"Copy failed for {fname}: {e}")
    else:
        st.warning("Support folder not found. Enter a valid path above.")

    if missing_support:
        st.warning("Missing in support folder: " + ", ".join(missing_support))

    if not ready_uploads:
        st.info("Upload both Excel files to enable the Partial process.")
    elif missing_support:
        st.info("Place the missing support files in the folder before running.")
    else:
        if st.button("Run Partial Revenue Process"):
            with st.spinner("Running ncpartial.py... please wait"):
                result = subprocess.run([sys.executable, "ncpartial.py"], capture_output=True, text=True)
            st.code(result.stdout)
            if result.stderr:
                st.error(result.stderr)
            for produced in ["partialrevenue.xlsx", "Updated_Partial_Revenue.xlsx"]:
                if os.path.exists(produced):
                    ts_file = timestamped_filename(produced)
                    os.rename(produced, ts_file)
                    st.success(f"✅ Created: {ts_file}")
                    with open(ts_file, "rb") as f:
                        st.download_button(f"Download {ts_file}", f, file_name=ts_file,
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
