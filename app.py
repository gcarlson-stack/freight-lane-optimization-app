import io
from pathlib import Path
import pandas as pd
import streamlit as st
import math
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _safe_filename(name: str) -> str:
    return "".join(c if c.isalnum() or c in (" ", "_", "-") else "_" for c in name).strip().replace(" ", "_")

def _format_money(x):
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return ""

def ensure_origin_dest(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure df has origin_city, origin_state, dest_city, dest_state.
    If they are missing and df is not empty, derive them from lane_detail / Lane_Detail / lane_key
    using split_lane_detail().
    If df is empty, just add empty string columns.
    """
    needed = ["origin_city", "origin_state", "dest_city", "dest_state"]

    # If all columns already exist, nothing to do
    if all(c in df.columns for c in needed):
        return df

    # If the DF is empty, just add blank columns and return
    if df.empty:
        for c in needed:
            if c not in df.columns:
                df[c] = ""
        return df

    # Choose source text: lane_detail (or Lane_Detail) if present, else lane_key
    if "lane_detail" in df.columns:
        src = df["lane_detail"]
    elif "Lane_Detail" in df.columns:
        src = df["Lane_Detail"]
    else:
        src = df["_lane"]

    # Apply split_lane_detail row-by-row, naming the 4 outputs explicitly
    od = src.apply(
        lambda x: pd.Series(
            split_lane_detail(x),
            index=["origin_city", "origin_state", "dest_city", "dest_state"],
        )
    )

    # Now od **always** has these 4 columns, even if df has only 1 row
    for c in needed:
        if c not in df.columns:
            df[c] = od[c]

    return df

def _clean_cs(x):
    """Clean city/state component for lane_key_compact: uppercase, no spaces."""
    if pd.isna(x):
        return ""
    return str(x).strip().upper().replace(" ", "")

def build_letter_docx(
    carrier: str,
    df_carrier: pd.DataFrame,
    sender_company: str,
    sender_name: str,
    sender_title: str,
    reply_by: str,
    body_template: str,
) -> bytes:
    """
    Create a single DOCX letter for a specific carrier with a table of NEGOTIATE lanes.
    Includes % over benchmark in the table.
    """

    # Make sure origin/dest columns exist
    df_carrier = ensure_origin_dest(df_carrier.copy())

    doc = Document()

    # ---- Title ----
    title = doc.add_paragraph(f"Rate Review Request — {carrier}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.size = Pt(16)
        run.bold = True

    # ---- Body text with placeholders ----
    num_lanes = len(df_carrier)
    if "delta_pct" in df_carrier.columns:
        avg_over_pct = df_carrier["delta_pct"].mean(skipna=True)
        avg_over_pct_text = f"{avg_over_pct:.0f}%" if pd.notna(avg_over_pct) else "N/A"
    else:
        avg_over_pct_text = "N/A"

    body_text = body_template.format(
        num_lanes=num_lanes,
        avg_over_pct=avg_over_pct_text,
        reply_by=reply_by,
    )
    doc.add_paragraph(body_text)

    # ---- Summary line ----
    total_delta = df_carrier["delta"].sum(skipna=True) if "delta" in df_carrier.columns else 0
    doc.add_paragraph(
        f"Lanes to negotiate: {num_lanes} • Current total variance vs. benchmark: {_format_money(total_delta)}"
    )

    # ---- Table: Origin/Dest + Costs + % over benchmark ----
    # ---- Aggregate to unique lanes with frequency ----
    # We group by origin/dest and compute:
    # - Frequency = number of rows (shipments) on that lane
    # - Company Cost = average per-move company cost
    # - Benchmark Cost = average per-move benchmark cost
    # - % over Benchmark = average delta_pct
    lane_group = (
        df_carrier
        .groupby(
            ["origin_city", "origin_state", "dest_city", "dest_state"],
            dropna=False
        )
        .agg(
            Frequency=("_lane", "size"),
            Company_Cost=("company_cost", "mean"),
            Benchmark_Cost=("benchmark_cost", "mean"),
            Over_Pct=("delta_pct", "mean"),
        )
        .reset_index()
    )

    # ---- Table: Unique Lanes + Frequency + Costs + % over benchmark ----
    table_cols = [
        "origin_city",
        "origin_state",
        "dest_city",
        "dest_state",
        "Frequency",
        "Company_Cost",
        "Benchmark_Cost",
        "Over_Pct",
    ]
    pretty = lane_group[table_cols].copy()

    # Rename for nicer headers
    pretty.columns = [
        "Origin City",
        "Origin State",
        "Dest City",
        "Dest State",
        "Frequency",
        "Company Cost",
        "Benchmark Cost",
        "% over Benchmark",
    ]

    table = doc.add_table(rows=1, cols=len(pretty.columns))
    hdr = table.rows[0].cells
    for i, c in enumerate(pretty.columns):
        p = hdr[i].paragraphs[0]
        run = p.add_run(c)
        run.bold = True
        run.font.size = Pt(8)       # decrease header font size
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # optional: center headers

    for _, row in pretty.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(row["Origin City"])
        cells[1].text = str(row["Origin State"])
        cells[2].text = str(row["Dest City"])
        cells[3].text = str(row["Dest State"])
        cells[4].text = str(int(row["Frequency"])) if not pd.isna(row["Frequency"]) else ""
        cells[5].text = _format_money(row["Company Cost"])
        cells[6].text = _format_money(row["Benchmark Cost"])
        if not pd.isna(row["% over Benchmark"]):
            cells[7].text = f"{row['% over Benchmark']:.0f}%"
        else:
            cells[7].text = ""
   
    # Small note under table explaining benchmark usage
    note = doc.add_paragraph(
        " Benchmark cost reflects market-based rates for comparable lanes and serves as a reference point for rate review."
    )
    note.paragraph_format.space_before = Pt(4)
    note.paragraph_format.space_after = Pt(12)
    note.runs[0].italic = True
    note.runs[0].font.size = Pt(9)

    doc.add_paragraph("")  # spacer

    # ---- Closing/signature ----
    doc.add_paragraph("We appreciate your partnership and look forward to your response.")
    doc.add_paragraph(f"\n{sender_name}\n{sender_title}\n{sender_company}")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

def build_rfp_letter_docx(
    carrier: str,
    df_carrier: pd.DataFrame,
    sender_company: str,
    sender_name: str,
    sender_title: str,
    reply_by: str,
    rfp_body_template: str,
    total_rfp_lanes: int,
) -> bytes:
    """
    Create a DOCX RFP letter for a specific carrier.
    Shows unique lanes with Frequency, top 10% by delta, a remaining summary row,
    and a total row. Company/benchmark costs are per-trip (averages).
    """
    # Ensure origin/dest exists
    df_carrier = ensure_origin_dest(df_carrier.copy())

    # --- Aggregate to unique lanes (per-trip metrics + frequency) ---
    lane_group = (
        df_carrier
        .groupby(
            ["origin_city", "origin_state", "dest_city", "dest_state"],
            dropna=False
        )
        .agg(
            Frequency=("_lane", "size"),
            Company_Cost=("company_cost", "mean"),
            Benchmark_Cost=("benchmark_cost", "mean"),
            Delta=("delta", "sum"),          # total $ variance per lane
            Over_Pct=("delta_pct", "mean"),  # avg % over benchmark
        )
        .reset_index()
    )

    unique_lane_count = len(lane_group)
    if unique_lane_count == 0:
        # Just return a minimal letter if no lanes (shouldn't normally happen)
        doc = Document()
        p = doc.add_paragraph(f"RFP Letter – {carrier}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        bio = BytesIO()
        doc.save(bio); bio.seek(0)
        return bio.read()

    # Sort lanes by greatest total delta
    lane_group = lane_group.sort_values("Delta", ascending=False, na_position="last")
    
    # --- Top 10% lanes by Delta ---
    top_n = max(1, math.ceil(0.10 * unique_lane_count))
    top_lanes = lane_group.head(top_n)
    remaining_lanes = lane_group.iloc[top_n:]

    remaining_lane_count = len(remaining_lanes)
    remaining_delta = remaining_lanes["Delta"].sum(skipna=True) if remaining_lane_count > 0 else 0.0

    total_delta = lane_group["Delta"].sum(skipna=True)
    # For body text, use average % over benchmark across this carrier's RFP lanes
    avg_over_pct = df_carrier["delta_pct"].mean(skipna=True)
    avg_over_pct_text = f"{avg_over_pct:.0f}%" if pd.notna(avg_over_pct) else "N/A"

    # lane_group: each row is a unique lane
    total_lane_count = unique_lane_count
    total_volume = lane_group["Frequency"].sum()
    
    remaining_lane_count = len(remaining_lanes)
    remaining_volume = remaining_lanes["Frequency"].sum() if remaining_lane_count > 0 else 0

    # --- Build DOCX ---
    doc = Document()

    # Title
    title = doc.add_paragraph(f"RFP Participation – {carrier}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.size = Pt(16)
        run.bold = True

    # Body with template placeholders
    # {total_rfp_lanes} = total lanes in the overall RFP (across all carriers)
    # {carrier_lane_count} = this carrier's lane_count
    # {avg_over_pct} = average % over benchmark for this carrier
    body_text = rfp_body_template.format(
        total_rfp_lanes=total_rfp_lanes,
        carrier_lane_count=unique_lane_count,
        avg_over_pct=avg_over_pct_text,
        reply_by=reply_by,
    )
    doc.add_paragraph(body_text)

    # --- Build table: top 10% lanes + remaining summary + total row ---
    table_cols = [
        "Origin City", "Origin State", "Dest City", "Dest State",
        "Frequency", "Company Cost", "Benchmark Cost",
        "% over Benchmark", "Total Delta",
    ]

    table = doc.add_table(rows=1, cols=len(table_cols))
    hdr = table.rows[0].cells
    for i, c in enumerate(table_cols):
        p = hdr[i].paragraphs[0]
        run = p.add_run(c)
        run.bold = True
        run.font.size = Pt(8)       # decrease header font size
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # optional: center headers

    def set_small_font(cell, size=8):
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size)

    # Helper to add a row safely
    def add_lane_row(row_data):
        row_cells = table.add_row().cells
        row_cells[0].text = str(row_data.get("Origin City", ""))
        row_cells[1].text = str(row_data.get("Origin State", ""))
        row_cells[2].text = str(row_data.get("Dest City", ""))
        row_cells[3].text = str(row_data.get("Dest State", ""))
        
        freq = row_data.get("Frequency", "")
        row_cells[4].text = str(int(freq)) if freq not in ("", None) and not pd.isna(freq) else ""
    
        row_cells[5].text = _format_money(row_data.get("Company Cost"))
        row_cells[6].text = _format_money(row_data.get("Benchmark Cost"))
    
        pct = row_data.get("% over Benchmark", None)
        row_cells[7].text = f"{pct:.0f}%" if pd.notna(pct) else ""
    
        row_cells[8].text = _format_money(row_data.get("Total Delta"))

        # ---- APPLY SMALL FONT TO ALL BODY CELLS ----
        for cell in row_cells:
            set_small_font(cell, size=8)

    # Top 10% lanes rows
    for _, r in top_lanes.iterrows():
        add_lane_row({
            "Origin City": r["origin_city"],
            "Origin State": r["origin_state"],
            "Dest City": r["dest_city"],
            "Dest State": r["dest_state"],
            "Frequency": r["Frequency"],
            "Company Cost": r["Company_Cost"],
            "Benchmark Cost": r["Benchmark_Cost"],
            "% over Benchmark": r["Over_Pct"],
            "Total Delta": r["Delta"],
        })

    # Remaining lanes summary row
    if remaining_lane_count > 0:
        add_lane_row({
            # label now includes lane count
            "Origin City": f"Remaining lanes ({remaining_lane_count})",
            "Origin State": "",
            "Dest City": "",
            "Dest State": "",
            # Frequency column now shows total shipment volume for remaining lanes
            "Frequency": remaining_volume,
            "Company Cost": None,
            "Benchmark Cost": None,
            "% over Benchmark": None,
            "Total Delta": remaining_delta,
        })
    
    # Total row
    add_lane_row({
        # label now includes total lane count
        "Origin City": f"Total (all {total_lane_count} lanes)",
        "Origin State": "",
        "Dest City": "",
        "Dest State": "",
        # Frequency column shows total shipment volume across all lanes
        "Frequency": total_volume,
        "Company Cost": None,
        "Benchmark Cost": None,
        "% over Benchmark": None,
        "Total Delta": total_delta,
    })

    doc.add_paragraph("")  # spacer
    doc.add_paragraph("We appreciate your partnership and look forward to your participation in the RFP.")
    doc.add_paragraph(f"\n{sender_name}\n{sender_title}\n{sender_company}")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()


def build_letters_zip(df_all, include_privfleet: bool, sender_company: str, sender_name: str,
                      sender_title: str, reply_by: str, body_template: str) -> bytes:
    if not include_privfleet:
        df_all = df_all[df_all["carrier_name"].astype(str).str.upper() != "GREIF PRIVATE FLEET"]

    df_neg = df_all[df_all["action"] == "NEGOTIATE"].copy()
    if df_neg.empty:
        buff = BytesIO()
        with ZipFile(buff, "w", ZIP_DEFLATED):
            pass
        buff.seek(0)
        return buff.read()

    buff = BytesIO()
    with ZipFile(buff, "w", ZIP_DEFLATED) as zf:
        for carrier, dfc in df_neg.groupby("carrier_name", dropna=False):
            carrier_name = "UNKNOWN" if pd.isna(carrier) or str(carrier).strip() == "" else str(carrier)
            dfc = dfc.sort_values("delta", ascending=False, na_position="last")

            doc_bytes = build_letter_docx(
                carrier=carrier_name,
                df_carrier=dfc,
                sender_company=sender_company,
                sender_name=sender_name,
                sender_title=sender_title,
                reply_by=reply_by,
                body_template=body_template,
            )
            fname = f"Negotiation_{_safe_filename(carrier_name)}.docx"
            zf.writestr(fname, doc_bytes)

    buff.seek(0)
    return buff.read()

def build_rfp_letters_zip(
    rfp_df: pd.DataFrame,
    sender_company: str,
    sender_name: str,
    sender_title: str,
    reply_by: str,
    rfp_body_template: str,
    include_privfleet: bool = False,
) -> bytes:
    """
    Build a ZIP with one RFP letter per carrier present in rfp_df.
    """
    df_all = rfp_df.copy()

    if not include_privfleet:
        df_all = df_all[df_all["carrier_name"].astype(str).str.upper() != "GREIF PRIVATE FLEET"]

    if df_all.empty:
        buff = BytesIO()
        with ZipFile(buff, "w", ZIP_DEFLATED):
            pass
        buff.seek(0)
        return buff.read()

    total_rfp_lanes = int(df_all["_lane"].nunique())

    buff = BytesIO()
    with ZipFile(buff, "w", ZIP_DEFLATED) as zf:
        for carrier, dfc in df_all.groupby("carrier_name", dropna=False):
            carrier_name = "UNKNOWN" if pd.isna(carrier) or str(carrier).strip() == "" else str(carrier)

            # Only consider lanes that are above benchmark (NEGOTIATE) and in RFP
            dfc_use = dfc[dfc["action"] == "NEGOTIATE"].copy()
            if dfc_use.empty:
                continue

            doc_bytes = build_rfp_letter_docx(
                carrier=carrier_name,
                df_carrier=dfc_use,
                sender_company=sender_company,
                sender_name=sender_name,
                sender_title=sender_title,
                reply_by=reply_by,
                rfp_body_template=rfp_body_template,
                total_rfp_lanes=total_rfp_lanes,
            )

            fname = f"RFP_Letter_{_safe_filename(carrier_name)}.docx"
            zf.writestr(fname, doc_bytes)

    buff.seek(0)
    return buff.read()

def split_lane_detail(text: str):
    """
    Split a lane detail like 'TAYLORS,SC to HOUSTON,TX' into:
    origin_city, origin_state, dest_city, dest_state

    If it can't parse, returns empty strings.
    """
    if not isinstance(text, str):
        return "", "", "", ""
    s = text.strip()
    if not s:
        return "", "", "", ""

    # Expect 'ORIG_CITY,ST to DEST_CITY,ST'
    parts = s.split(" to ")
    if len(parts) != 2:
        return "", "", "", ""

    orig, dest = parts[0].strip(), parts[1].strip()

    def split_part(p):
        if "," not in p:
            return p.strip(), ""
        city, st = p.rsplit(",", 1)
        return city.strip(), st.strip()

    oc, os = split_part(orig)
    dc, ds = split_part(dest)
    return oc, os, dc, ds

st.set_page_config(page_title="Freight Lane Comparison", layout="wide")

# Fixed location exclusions
FIXED_EXCLUDE_LOCATIONS = [
    "GLADSTONEVA", "FITCHBURGMA", "MORGAN HILLCA", "MASSILLONOH", "MERCEDCA",
    "LOUISVILLEKY", "MASONMI", "GREENSBORONC", "CONCORDNC", "PALMYRAPA", "DALLASTX"
]

def norm_lane(x):
    if pd.isna(x):
        return ""
    return str(x).strip().upper()

def build_compact_key(df, city_col, state_col, new_col_name):
    """
    Build a compact 'CITYSTATE' lane key from separate city/state columns.
    Example: 'DETROIT' + 'MI' -> 'DETROITMI'

    - df:          DataFrame to modify in-place
    - city_col:    column with city names
    - state_col:   column with state codes
    - new_col_name: name of the resulting lane-key column
    """
    # If either column is missing, do nothing
    if city_col not in df.columns or state_col not in df.columns:
        return df

    city = df[city_col].astype(str).str.strip().str.upper()
    state = df[state_col].astype(str).str.strip().str.upper()

    # Remove spaces so 'SAINT LOUIS' + 'MO' -> 'SAINTLOUISMO'
    city = city.str.replace(" ", "", regex=False)
    state = state.str.replace(" ", "", regex=False)

    df[new_col_name] = city + state
    return df

def parse_percent_col(series: pd.Series) -> pd.Series:
    """
    Convert a column that may be:
      - '30%'   -> 0.30
      - '30'    -> 0.30
      - 30      -> 0.30
      - 0.30    -> 0.30  (already decimal)
    into a decimal fraction (0.30).
    """
    s = series.astype(str).str.strip().str.replace("%", "", regex=False)
    vals = pd.to_numeric(s, errors="coerce")

    # If the largest non-null value is <= 1, it is already a fraction (e.g. 0.30)
    max_val = vals.max(skipna=True)
    if max_val is not None and max_val <= 1.0:
        return vals

    # Otherwise we assume it's in whole percent (30 -> 0.30)
    return vals / 100.0

def read_any(upload, sheet=None):
    name = upload.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(upload)
    elif name.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(upload, sheet_name=sheet)
    elif name.endswith(".xlsb"):
        return pd.read_excel(upload, engine="pyxlsb", sheet_name=sheet)
    else:
        raise ValueError("Unsupported file type. Please upload CSV/XLSX/XLS/XLSB.")

def infer_sheets(upload):
    if upload is None: 
        return []
    name = upload.name.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls", ".xlsb")):
        try:
            xl = pd.ExcelFile(upload)
            return xl.sheet_names
        except Exception:
            return []
    return []
@st.cache_data
def read_any_cached(upload, sheet=None):
    """Cached wrapper around read_any to avoid re-reading the same file."""
    if upload is None:
        return None
    return read_any(upload, sheet=sheet)


@st.cache_data
def infer_sheets(upload):
    if upload is None: 
        return []
    name = upload.name.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls", ".xlsb")):
        try:
            xl = pd.ExcelFile(upload)
            return xl.sheet_names
        except Exception:
            return []
    return []
st.title("🚚 FLO.ai")

st.markdown("""
### What this tool does

This app compares your **company freight costs** to **benchmark rates**, then classifies each lane into:
- **RFP lanes** – lanes that will be included in a broader bid event
- **Negotiation lanes** – "non-vanilla" lanes or lanes excluded from RFP for specific reasons that 
    will be handled via targeted vendor letters. Letters can also be used for monthly rate review, 
    e.g., month-to-month negotiations with carriers to monitor variance to benchmark
- **Excluded lanes** – filtered out due to location, mode, or carrier exclusions

From this, you can:
- Download a **full comparison workbook**
- Generate an **RFP template** (Overview, Timeline, Bid Lanes, Locations)
- Generate **RFP letters** to carriers participating in the bid
- Generate **negotiation letters** for "non-vanilla" lanes flagged for direct rate review

**Note: when the program is running or loading, the screen will "gray out" and/or icons of a person doing different sports will appear in the top righthand corner. Please do not refresh the page 
or make changes to any of the inputs while the page is loading.**

""")

with st.sidebar:
    st.header("🧭 How to use this tool")

    st.markdown("""
**Please note: when the app is running, the website will "gray out." Please do not refresh the page 
or make changes to any of the inputs while the page is loading.**

**Step 1 – Upload data**
1. Upload **Company** and **Benchmark** files.
2. Map the correct columns (lane, cost, carrier, mode).
3. Add any **location, carrier, or mode exclusions**.

**Step 2 – Run comparison**
- Click **Run comparison** to:
  - Match lanes to benchmark
  - Compute $ and % deltas
  - Classify lanes into RFP vs Negotiation vs Excluded.

**When to generate RFP Template + RFP Letters**
- Use when you want to run a **formal bid event** across multiple lanes/carriers.
- A formal bid event should be run no more than twice in a 12-month period, i.e., a formal bid 
    should be run once every 6 months.
- The **RFP Template** gives you:
  - Tab 1: Overview & Instructions
  - Tab 2: Bid Timeline
  - Tab 3: Bid Lanes (with Market Rate)
  - Tab 4: Location List
- The **RFP Letters**:
  - Summarize the lanes above benchmark for each carrier
  - Show **top 10% lanes by variance** plus a summary for the rest.

**When to generate Negotiation Letters**
- Use when you **do not** want a full RFP, but instead:
  - Target specific lanes for **direct rate negotiation**
  - Send **custom letters** to carriers with lane-level detail and % over benchmark.
- Negotiation letters can be used as a monthly check for lanes above benchmark. 
    Letters can be sent once a month to vendors notifying vendors that lanes are 
    still not within compliance of benchmark values.
    """)

    st.caption("Tip: Run the comparison first, review the RFP and Negotiation tabs, then generate the templates and letters.")

# ---- Session state init ----
if "results_ready" not in st.session_state:
    st.session_state["results_ready"] = False

# ============ Uploads ============
colL, colR = st.columns(2)
with colL:
    st.subheader("Company file")
    st.markdown("Upload a data export from your TMS system that contains lanes (origin and destination), carriers, base rate charged, transportation mode (e.g., TL). \n"
                "**Once the file is uploaded, select the sheet from the dropdown options that contains the relevant data.**")
    client_file = st.file_uploader("Upload Company data (CSV/XLSX/XLS/XLSB)", type=["csv","xlsx","xls","xlsb"], key="client")
    client_sheets = infer_sheets(client_file)
    client_sheet = st.selectbox("Company sheet (optional)", options=["<first sheet>"] + client_sheets if client_sheets else ["<first sheet>"])
    
    st.caption("If you don’t have a formatted company file with carrier names and modes, origin, destination, and base rate charged, download the template below, fill it in, and re-upload it as the company data.")

    # -------- Company template download --------
    template_cols = [
        "Origin City",
        "Origin State",
        "Dest City",
        "Dest State",
        "Total Base Charges",  # map this to company_cost_col
        "Carrier Name",      # map this to client_carrier_col
        "Carrier Mode",      # map this to mode_col
    ]

    template_df = pd.DataFrame(columns=template_cols)

    tmpl_buf = io.BytesIO()
    with pd.ExcelWriter(tmpl_buf, engine="openpyxl") as writer:
        template_df.to_excel(
            writer,
            index=False,
            sheet_name="Company Data Template",
        )
    tmpl_buf.seek(0)

    st.download_button(
        label="⬇️ Download Company Data Template (Excel)",
        data=tmpl_buf,
        file_name="company_data_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Build list of company columns for dropdown (if file is uploaded)
client_mode_columns = ["<None>"]
df_client_preview = pd.DataFrame()
if client_file is not None:
    try:
        sheet_c_preview = None if client_sheet == "<first sheet>" else client_sheet
        df_client_preview = read_any_cached(client_file, sheet_c_preview)
        client_mode_columns = ["<None>"] + list(df_client_preview.columns)
    except Exception as e:
        st.warning(f"Could not read client file to detect columns: {e}")
        
with colR:
    st.subheader("Benchmark file")
    st.markdown("Upload a data export from your benchmark source (e.g., DAT) that contains lanes (origin and destination) and base rate charged. \n"
                "**Once the file is uploaded, select the sheet from the dropdown options that contains the relevant data.**")
    bench_file = st.file_uploader("Upload Benchmark (CSV/XLSX/XLS/XLSB)", type=["csv","xlsx","xls","xlsb"], key="bench")
    bench_sheets = infer_sheets(bench_file)
    bench_sheet = st.selectbox("Benchmark sheet (optional)", options=["<first sheet>"] + bench_sheets if bench_sheets else ["<first sheet>"])
bench_mode_columns = ["<None>"]
if bench_file is not None:
    try:
        sheet_b_preview = None if bench_sheet == "<first sheet>" else bench_sheet
        df_bench_preview = read_any_cached(bench_file, sheet_b_preview)
        bench_mode_columns = ["<None>"] + list(df_bench_preview.columns)
    except Exception as e:
        st.warning(f"Could not read benchmark file to detect columns: {e}")
        
st.markdown("STOP - loading both files will likely take a few minutes. Please wait for the upload to complete before moving forward.")
st.markdown("---")

# ============ Columns & Options ============
st.subheader("Columns & Options")
st.markdown("Please enter the names of the columns where each data can be found in the company and benchmark files.")
c1, c2, c3, c4 = st.columns(4)

with c1:
    client_lane_col = st.selectbox(
        "Company lane column (key)", 
        options=client_mode_columns,
        index=0,
        help="Choose the column in the company data file that has the lane names.")
    company_cost_col = st.selectbox(
        "Company cost column", 
        options=client_mode_columns,
        index=0,
        help="Choose the column in the company data file that has the base charges (e.g., linehaul rates).")
    company_fuel_col = st.selectbox(
        "Company fuel surcharge $ column (optional)", 
        options=client_mode_columns,
        index=0,
        help="Choose the column in the company data file that has the fuel surcharges.")
    client_carrier_col = st.selectbox(
        "Company carrier column", 
        options=client_mode_columns,
        index=0,
        help="Choose the column in the company data file that has the carrier names.")
    lane_detail_col = st.text_input(
        "Company lane detail column (e.g., 'Lane_Detail' or '_lane')",
        value="Lane_Detail"
    )
    mode_col = st.selectbox(
        "Company mode column (TL / LTL)",
        options=client_mode_columns,
        index=0,
        help="Choose the column in the client file that indicates TL vs LTL (or other modes)."
    )
    # Benchmark mode column
    bench_mode_col = st.selectbox(
        "Benchmark mode column (TL / LTL)",
        options=bench_mode_columns,
        index=0,
        help="Choose the column in the benchmark file that indicates TL vs LTL (or other modes)."
    )

with c2:
    # New controls to map city/state columns
    city_state_cols = ["<None>"] + (list(df_client_preview.columns) if client_file is not None else [])

    st.markdown("**Optional, use only if lane names are not a concatenated string and are listed across separate columns in the data file**: Map separate Origin / Destination city & state columns")
    origin_city_col = st.selectbox("Origin city column", options=city_state_cols, index=0)
    origin_state_col = st.selectbox("Origin state column", options=city_state_cols, index=0)
    dest_city_col   = st.selectbox("Dest city column",   options=city_state_cols, index=0)
    dest_state_col  = st.selectbox("Dest state column",  options=city_state_cols, index=0)

with c3:
    bench_lane_col = st.selectbox(
        "Benchmark lane column (key)", 
        options=bench_mode_columns,
        index=0,
        help="Choose the column in the benchmark data file that has the lane names.")
    bench_cost_col = st.selectbox(
        "Benchmark cost column", 
        options=bench_mode_columns,
        index=0,
        help="Choose the column in the benchmark data file that has the base charges (e.g., linehaul rates).")
    bench_fuel_col = st.selectbox(
        "Benchmark fuel surcharge % column (optional)", 
        options=bench_mode_columns,
        index=0,
        help="Choose the column in the benchmark data file that has the fuel surcharges.")
    bench_agg = st.selectbox("Benchmark duplicate lanes aggregation", options=["mean","median"], index=0)

with c4:
    apply_fixed_exclusions = st.checkbox(
        "Apply default location exclusions",
        value=True,  # default ON, same behavior as before
        help="If checked, the default list of locations will be excluded."
    )

    if apply_fixed_exclusions:
        st.caption("Default locations currently excluded:")
        st.code(", ".join(FIXED_EXCLUDE_LOCATIONS), language=None)

    extra_locations = st.text_area(
        "Extra locations to exclude (ALL CAPS, comma-separated)",
        placeholder="ATLANTAGA, CHICAGOIL, BOSTONMA"
    )
    extra_carriers = st.text_area(
        "Carriers to exclude (comma-separated, case-insensitive)",
        placeholder="CARRIER A, CARRIER B"
    )

st.subheader("Mode Filter (exclude LTL, etc.)")
st.markdown("Select the column from the dropdown that contains the transportation mode in the company data file. Then fill in the mode that is to be excluded.") 
m1, m2 = st.columns(2)
with m1:
    mode_col = st.selectbox(
        "Company mode column (LTL/TL)",
        options=client_mode_columns,
        index=0,  # "<None>" by default
        help="Choose the column that contains TL / LTL (or other mode) information."
    )
with m2:
    exclude_modes_raw = st.text_input(
        "Modes to exclude (comma-separated, case-insensitive)",
        value="LTL",
        help="Example: LTL, PARCEL"
    )

st.subheader("RFP Overrides")
st.markdown("Provide lane names that are to be excluded from the RFP. Ensure lanes are entered in the format provided in the example.")
letter_override_raw = st.text_area(
    "Lane keys to treat as Vendor Letters instead of RFP (comma or newline separated)",
    placeholder="Example: ATLANTAGADALLASTX, CHICAGOILHOUSTONTX",
    help="Provide _lane values exactly as they appear in the output (e.g., ORIGINSTATEDESTSTATE). "
         "Any lane listed here will be handled via vendor letters instead of RFP."
)

# Normalize override lane list (upper-case, stripped)
override_letter_lanes = set()
if letter_override_raw.strip():
    # accept commas OR newlines
    pieces = []
    for chunk in letter_override_raw.replace("\n", ",").split(","):
        c = chunk.strip()
        if c:
            pieces.append(c)
    override_letter_lanes = {p.upper() for p in pieces}

run = st.button("Run comparison")

# If user clicked "Run comparison", recompute and store results
if run:
    # ============ Load data ============
    if client_file is None or bench_file is None:
        st.error("Please upload both Company and Benchmark files.")
        st.stop()

    sheet_c = None if client_sheet == "<first sheet>" else client_sheet
    sheet_b = None if bench_sheet == "<first sheet>" else bench_sheet

    try:
        df_client = read_any_cached(client_file, sheet_c)
        df_bench = read_any_cached(bench_file, sheet_b)
    except Exception as e:
        st.error(f"Error loading files: {e}")
        st.stop()
    
    # --- Determine whether both datasets have mode columns mapped ---
    client_has_mode = (mode_col not in ("<None>", None, "") 
                       and mode_col in df_client.columns)
    bench_has_mode = (bench_mode_col not in ("<None>", None, "") 
                      and bench_mode_col in df_bench.columns)
    
    use_mode_matching = client_has_mode and bench_has_mode
    
    # ---------- Optionally build _lane from city/state BEFORE validation ----------
    # Only try to build a compact key if the user chose city/state columns
    using_city_state = (
        origin_city_col  not in ("<None>", None, "") and
        origin_state_col not in ("<None>", None, "") and
        dest_city_col    not in ("<None>", None, "") and
        dest_state_col   not in ("<None>", None, "")
    )

    # If the user mapped city/state columns and the lane key column doesn't exist yet,
    # build it as ORIGCITY+ORIGSTATE+DESTCITY+DESTSTATE (no spaces, all caps).
    if using_city_state and client_lane_col not in df_client.columns:
        missing_any = [
            c for c in [origin_city_col, origin_state_col, dest_city_col, dest_state_col]
            if c not in df_client.columns
        ]
        if missing_any:
            st.error(
                f"Client lane column '{client_lane_col}' is missing and the mapped "
                f"city/state columns {missing_any} are not found in the client data."
            )
            st.write("Company columns:", list(df_client.columns))
            st.stop()

        st.info(
            f"Client lane column '{client_lane_col}' not found. "
            "Building it from origin/dest city & state columns "
            f"('{origin_city_col}', '{origin_state_col}', '{dest_city_col}', '{dest_state_col}')."
        )

        df_client[client_lane_col] = (
            df_client[origin_city_col].apply(_clean_cs)
            + df_client[origin_state_col].apply(_clean_cs)
            + df_client[dest_city_col].apply(_clean_cs)
            + df_client[dest_state_col].apply(_clean_cs)
        )

    # ---------- Now validate columns (after any auto-build) ----------
    missing_client = [
        c for c in [client_lane_col, company_cost_col, client_carrier_col]
        if c not in df_client.columns
    ]
    missing_bench = [
        c for c in [bench_lane_col, bench_cost_col]
        if c not in df_bench.columns
    ]

    if missing_client:
        st.error(f"Company file missing columns: {missing_client}")
        st.write("Company columns:", list(df_client.columns))
        st.stop()
    if missing_bench:
        st.error(f"Benchmark file missing columns: {missing_bench}")
        st.write("Benchmark columns:", list(df_bench.columns))
        st.stop()

    # ============ Normalize & select ============

    df_client = df_client.copy()
    df_bench = df_bench.copy()
    # Ensure origin/dest columns exist on both datasets
    df_client = ensure_origin_dest(df_client)
    df_bench = ensure_origin_dest(df_bench)

    df_client["_lane"] = df_client[client_lane_col].map(norm_lane)
    df_bench["_lane"]  = df_bench[bench_lane_col].map(norm_lane)

    def norm_mode(series):
        return (
            series.astype(str)
                  .str.strip()
                  .str.upper()
                  .replace({
                      "TRUCKLOAD": "TL",
                      "TL": "TL",
                      "LTL": "LTL",
                      "LESS THAN TRUCKLOAD": "LTL",
                  })
        )

    if use_mode_matching:
        df_client["_mode"] = norm_mode(df_client[mode_col])
        df_bench["_mode"]  = norm_mode(df_bench[bench_mode_col])
    else:
        # default placeholder for mode so merge still works safely
        df_client["_mode"] = "DEFAULT"
        df_bench["_mode"]  = "DEFAULT"
   
    df_client["company_linehaul"] = pd.to_numeric(
        df_client[company_cost_col], errors="coerce"
    )
    
    company_has_fuel = (
        company_fuel_col not in (None, "", "<None>")
        and company_fuel_col in df_client.columns
    )
    
    if company_has_fuel:
        # company fuel surcharge is already a $ amount per lane
        df_client["company_fuel_cost"] = pd.to_numeric(
            df_client[company_fuel_col], errors="coerce"
        )
    else:
        df_client["company_fuel_cost"] = 0.0
    
    # total company cost = linehaul + fuel
    df_client["company_cost"] = df_client["company_linehaul"] + df_client["company_fuel_cost"]
    
    # ----- BENCHMARK: linehaul + fuel + total -----
    
    # Make sure "mode" exists on benchmark side using the selected mode column
    df_bench["mode"] = df_bench[bench_mode_col]
    
    # Linehaul ($)
    df_bench["benchmark_linehaul"] = pd.to_numeric(
        df_bench[bench_cost_col],
        errors="coerce",
    )
    
    # Fuel surcharge (%) -> decimal fraction, if we have a column selected
    if bench_fuel_col != "<None>" and bench_fuel_col in df_bench.columns:
        bench_fuel_pct = parse_percent_col(df_bench[bench_fuel_col])
        df_bench["benchmark_fuel_cost"] = df_bench["benchmark_linehaul"] * bench_fuel_pct
    else:
        df_bench["benchmark_fuel_cost"] = 0.0
    
    # Total benchmark cost
    df_bench["benchmark_cost"] = (
        df_bench["benchmark_linehaul"] + df_bench["benchmark_fuel_cost"]
    )

    # --- select client columns, including lane detail if it exists ---
    # Build the list of columns we want to keep from the company file
    client_cols_to_keep = [
        client_lane_col,
        client_carrier_col,
        "company_linehaul",
        "company_fuel_cost",
        "company_cost",
        "_lane",
        "_mode",
    ]
    
    if lane_detail_col in df_client.columns:
        client_cols_to_keep.append(lane_detail_col)
    
    client_keep = df_client[client_cols_to_keep].rename(
        columns={
            client_lane_col: "_lane",
            client_carrier_col: "carrier_name",
            "_mode": "mode",
        }
    )

    # Linehaul (benchmark) as numeric
    df_bench["benchmark_linehaul"] = pd.to_numeric(
        df_bench[bench_linehaul_col], errors="coerce"
    )

   # Linehaul numeric
    df_bench["benchmark_linehaul"] = pd.to_numeric(
        df_bench[bench_linehaul_col], errors="coerce"
    )
    
    # Fuel percentage -> fraction using the robust parser
    bench_fuel_pct = parse_percent_col(df_bench[bench_fuel_pct_col])
    
    df_bench["benchmark_fuel_cost"] = df_bench["benchmark_linehaul"] * bench_fuel_pct
    df_bench["benchmark_cost"] = (
        df_bench["benchmark_linehaul"] + df_bench["benchmark_fuel_cost"]
    )
    
    # ============ Apply exclusions ============
    # Locations
    exclude_locations = FIXED_EXCLUDE_LOCATIONS.copy()
    if apply_fixed_exclusions:
        exclude_locations = FIXED_EXCLUDE_LOCATIONS.copy()
    else:
        exclude_locations = []
    
    if extra_locations:
        extra_locs_list = [x.strip().upper() for x in extra_locations.split(",") if x.strip()]
        exclude_locations.extend(extra_locs_list)

    client_keep["lane_key_upper"] = client_keep["_lane"].astype(str).str.upper()

    def match_locs(text):
        return [loc for loc in exclude_locations if loc in text]

    client_keep["excluded_matches"] = client_keep["lane_key_upper"].apply(match_locs)
    mask_loc_excl = client_keep["excluded_matches"].str.len().fillna(0).gt(0)
    excluded_detail_df = client_keep.loc[mask_loc_excl, ["_lane", "carrier_name", "excluded_matches"]].copy()
    excluded_exploded = excluded_detail_df.explode("excluded_matches")
    if len(excluded_exploded):
        excluded_summary_df = (
            excluded_exploded["excluded_matches"]
            .value_counts(dropna=False)
            .rename_axis("Excluded_Location")
            .reset_index(name="Count")
        )
    else:
        excluded_summary_df = pd.DataFrame(columns=["Excluded_Location","Count"])

    client_keep = client_keep.loc[~mask_loc_excl].drop(columns=["excluded_matches"], errors="ignore")

    # Carriers
    if extra_carriers:
        carriers_list = [c.strip().upper() for c in extra_carriers.split(",") if c.strip()]
        before = len(client_keep)
        client_keep = client_keep[~client_keep["carrier_name"].astype(str).str.upper().isin(carriers_list)]
        st.info(f"Carrier exclusions removed {before - len(client_keep)} rows.")

    # ============ Mode exclusions (e.g., exclude LTL) ============
    # exclude_modes_raw is your text input: "LTL, PARCEL", etc.
    
    exclude_modes = [
        m.strip().upper()
        for m in exclude_modes_raw.split(",")
        if m.strip()
    ]
    
    if exclude_modes:
        if use_mode_matching and "mode" in client_keep.columns:
            # We have a real mode field (TL/LTL/etc.) and user wants exclusions
            before = len(client_keep)
            client_keep = client_keep[
                ~client_keep["mode"].astype(str).str.upper().isin(exclude_modes)
            ]
            removed = before - len(client_keep)
            if removed > 0:
                st.info(
                    f"Mode exclusions removed {removed} rows "
                    f"(modes excluded: {', '.join(exclude_modes)})."
                )
        else:
            # Either the user turned mode matching off, or the mode columns
            # weren’t mapped correctly when building _mode earlier.
            st.warning(
                "Mode-based exclusions were requested, but mode matching is not enabled "
                "or no valid mode column is available. No mode-based exclusions applied."
            )
    # If exclude_modes is empty, quietly do nothing
    # ---------- Build benchmark frame used for aggregation ----------
    # This assumes df_bench already has:
    #   _lane, _mode, benchmark_linehaul, benchmark_fuel_cost, benchmark_cost
    # Keep only the columns we need on the benchmark side
    df_bench["_lane"] = df_bench[bench_lane_col]
    df_bench["_mode"] = df_bench[bench_mode_col]
    bench_cols_to_keep = [
    "_lane",
    "_mode",
    "benchmark_linehaul",
    "benchmark_fuel_cost",
    "benchmark_cost",
    ]
    
    bench_keep = df_bench[bench_cols_to_keep].rename(
        columns={"_mode": "mode"}
    )

    # ============ Build benchmark aggregate (one row per lane) ============
    group_cols = ["_lane"] + (["mode"] if use_mode_matching else [])
    value_cols = ["benchmark_linehaul", "benchmark_fuel_cost", "benchmark_cost"]
    
    agg_func = "median" if bench_agg == "median" else "mean"
    
    bench_agg_df = (
        bench_keep
        .groupby(group_cols, as_index=False, dropna=False)[value_cols]
        .agg(agg_func)
    )
    if use_mode_matching:
        merged = client_keep.merge(
            bench_agg_df,
            how="left",
            on=["_lane", "mode"],
        )
    else:
        merged = client_keep.merge(
            bench_agg_df,
            how="left",
            on="_lane",
        )

    # If there is truly no benchmark match for a lane, keep 0 on the benchmark side
    merged["benchmark_linehaul"] = merged["benchmark_linehaul"].fillna(0.0)
    merged["benchmark_fuel_cost"] = merged["benchmark_fuel_cost"].fillna(0.0)
    merged["benchmark_cost"] = merged["benchmark_cost"].fillna(0.0)
    
    # Deltas
    merged["delta_linehaul"] = merged["company_linehaul"] - merged["benchmark_linehaul"]
    merged["delta_fuel"] = merged["company_fuel_cost"] - merged["benchmark_fuel_cost"]
    merged["delta"] = merged["company_cost"] - merged["benchmark_cost"]
    
    mask = merged["benchmark_cost"] != 0
    merged["delta_pct"] = None
    merged.loc[mask, "delta_pct"] = (
        merged.loc[mask, "delta"] / merged.loc[mask, "benchmark_cost"] * 100.0
    )
    
    merged["action"] = merged["delta"].apply(
        lambda d: "NEGOTIATE" if pd.notna(d) and d > 0 else "None"
    )

    out = merged[
        [
            "_lane",
            "carrier_name",
            "mode",
            "company_linehaul",
            "company_fuel_cost",
            "company_cost",          # total
            "benchmark_linehaul",
            "benchmark_fuel_cost",
            "benchmark_cost",        # total
            "delta_linehaul",
            "delta_fuel",
            "delta",
            "delta_pct",
            "action",
            "origin_city",
            "origin_state",
            "dest_city",
            "dest_state",
        ]
    ].sort_values("delta", ascending=False, na_position="last")

    # Remove GREIF from main results (they get a separate tab/sheet)
    out = out[out["carrier_name"].astype(str).str.upper() != "GREIF PRIVATE FLEET"]
    out = ensure_origin_dest(out)
    # ============ GREIF (post-exclusion) ============
    gpf_rows = client_keep[client_keep["carrier_name"].astype(str).str.upper() == "GREIF PRIVATE FLEET"].copy()
    if gpf_rows.empty:
        gpf_export = pd.DataFrame(columns=["_lane","carrier_name","company_cost","benchmark_cost", "delta","action"])
        gpf_count = 0
        gpf_negotiate_count = 0
        gpf_total_delta = 0.0
    else:
        gpf_merged = gpf_rows.merge(bench_agg_df, how="left", on="_lane")
        gpf_merged["delta"] = gpf_merged["company_cost"] - gpf_merged["benchmark_cost"]
        gpf_merged["action"] = gpf_merged["delta"].apply(lambda d: "NEGOTIATE" if pd.notna(d) and d > 0 else "None")
        gpf_export = gpf_merged[["_lane","carrier_name","company_cost","benchmark_cost","delta","action"]] \
                            .sort_values("delta", ascending=False, na_position="last")
        gpf_count = int(gpf_export.shape[0])
        gpf_negotiate_count = int((gpf_export["action"] == "NEGOTIATE").sum())
        gpf_total_delta = float(gpf_export.loc[gpf_export["action"] == "NEGOTIATE", "delta"].sum(skipna=True))
        # Add origin/dest columns for Private Fleet lanes
    gpf_export = ensure_origin_dest(gpf_export)

    # ============ Summary ============
    neg_mask_out = out["action"] == "NEGOTIATE"
    overall_count = int(neg_mask_out.sum())
    
    # Scenario 1: only linehaul is negotiated; fuel surcharges left as-is
    linehaul_savings = float(
        out.loc[neg_mask_out, "delta_linehaul"].clip(lower=0).sum(skipna=True)
    )
    
    # Scenario 2: both linehaul and fuel negotiated
    fuel_savings = float(
        out.loc[neg_mask_out, "delta_fuel"].clip(lower=0).sum(skipna=True)
    )
    overall_total = linehaul_savings + fuel_savings  # total savings if both are negotiated

    summary_df = pd.DataFrame([
        {
            "Segment": "OVERALL",
            "Negotiate_Lanes": overall_count,
            "Total_Delta": overall_total,
            "Summary_Text": f"SUMMARY (OVERALL): {overall_count} lanes marked as NEGOTIATE with total delta ${overall_total:,.2f}."
        },
        {
            "Segment": "PRIVATE FLEET",
            "Negotiate_Lanes": gpf_negotiate_count,
            "Total_Delta": gpf_total_delta,
            "Summary_Text": f"SUMMARY: {gpf_count} lanes total; NEGOTIATE lanes delta ${gpf_total_delta:,.2f}."
        }
    ])

    # After you’ve built gpf_export
    if not gpf_export.empty:
        gpf_export = ensure_origin_dest(gpf_export)

    # ---- store everything in session_state ----
    st.session_state["out"] = out
    st.session_state["gpf_export"] = gpf_export
    st.session_state["summary_df"] = summary_df
    st.session_state["excluded_summary_df"] = excluded_summary_df
    st.session_state["excluded_detail_df"] = excluded_detail_df
    st.session_state["results_ready"] = True

# If we still don't have results, show a hint and stop
if not st.session_state["results_ready"]:
    st.info("Click 'Run comparison' to generate results.")
    st.stop()

# From here on, ALWAYS use the stored versions
out = st.session_state["out"]
gpf_export = st.session_state["gpf_export"]
summary_df = st.session_state["summary_df"]
excluded_summary_df = st.session_state["excluded_summary_df"]
excluded_detail_df = st.session_state["excluded_detail_df"]

# Derive Private Fleet stats from stored gpf_export
gpf_count = int(gpf_export.shape[0])
gpf_negotiate_count = int((gpf_export["action"] == "NEGOTIATE").sum()) if not gpf_export.empty else 0
gpf_total_delta = float(
    gpf_export.loc[gpf_export["action"] == "NEGOTIATE", "delta"].sum(skipna=True)
) if not gpf_export.empty else 0.0

# ============ RFP vs Letter vs No-Action Classification ============

# lane_key normalized for matching overrides
out["lane_key_norm"] = out["_lane"].astype(str).str.strip().str.upper()

def classify_row(row):
    """
    Logic:
      - Only consider lanes that are above benchmark (action == 'NEGOTIATE')
        and have a valid benchmark_cost.
      - By default, those lanes go to RFP.
      - If the lane_key is in override_letter_lanes, treat as Letter instead.
      - All other lanes => None (no action).
    """
    # must be above benchmark and have a usable benchmark
    if row["action"] != "NEGOTIATE":
        return "None"
    if pd.isna(row["benchmark_cost"]) or row["benchmark_cost"] <= 0:
        return "None"

    # user override: force to Letter
    if row["lane_key_norm"] in override_letter_lanes:
        return "Letter"

    # default: send out to bid
    return "RFP"

out["lane_treatment"] = out.apply(classify_row, axis=1)

# Split into views
rfp_df = out[out["lane_treatment"] == "RFP"].copy()
letter_df = out[out["lane_treatment"] == "Letter"].copy()
no_action_df = out[out["lane_treatment"] == "None"].copy()

# --- derive lane counts & savings from 'out', 'letter_df', 'rfp_df' ---

neg_mask_out = out["action"] == "NEGOTIATE"
overall_count = int(neg_mask_out.sum())
overall_total = float(out.loc[neg_mask_out, "delta"].sum(skipna=True))

letter_neg_mask = letter_df["action"] == "NEGOTIATE"
letter_lane_count = int(letter_neg_mask.sum())
letter_savings = float(letter_df.loc[letter_neg_mask, "delta"].sum(skipna=True))

rfp_neg_mask = rfp_df["action"] == "NEGOTIATE"
rfp_lane_count = int(rfp_neg_mask.sum())
rfp_savings = float(rfp_df.loc[rfp_neg_mask, "delta"].sum(skipna=True))

combined_savings = letter_savings + rfp_savings
savings_diff = overall_total - combined_savings

# ============ UI Output ============
st.markdown("---")
st.subheader("Output (Results of FLO.ai rate comparison)")
st.markdown(
    "This section contains 6 tabs for the output from the comparison exercise, including: \n"
    "1) 🧾 Summary: Total savings opportunity across lanes to be included in RFP and in negotiation letters (including how many lanes are included in Private Fleet and excluded from any bid exercise).\n"
    "2) 📦 RFP Candidates: Total number of lanes and value opportunity to be sent out to bid (included in RFP).\n"
    "3) 📊 Letter Candidates: Total number of lanes and value opportunity to receive negotiation letter according to lanes specified in \"RFP overrides\".\n"
    "4) 🚛 Private Fleet: Total number of lanes and value opportunity in Private Fleet lanes that are excluded from the comparison and analysis.\n"
    "5) 🚫 Excluded (Summary): Total number of lanes excluded from comparison and analysis according to user-set exclusions (location, carrier, lane).\n"
    "6) 🚫 Excluded (Detail): Lane by lane detail of exclusions."
)

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🧾 Summary",
    "📦 RFP Candidates",
    "📊 Letter Candidates",
    "🚛 Private Fleet",
    "🚫 Excluded (Summary)",
    "🚫 Excluded (Detail)"
])

linehaul_savings = float(
        out.loc[neg_mask_out, "delta_linehaul"].clip(lower=0).sum(skipna=True)
    )
fuel_savings = float(
    out.loc[neg_mask_out, "delta_fuel"].clip(lower=0).sum(skipna=True)
)
overall_total = linehaul_savings + fuel_savings
with tab1:
    st.markdown("### Savings scenarios")
    st.markdown(
        f"- **Scenario 1 – Linehaul only:** ${linehaul_savings:,.2f}  \n"
        f"- **Scenario 2 – Linehaul + fuel:** {overall_total:,.2f} \\(includes {fuel_savings:,.2f} from fuel surcharges\\)"
        f"(includes ${fuel_savings:,.2f} from fuel surcharges)"
    )
    st.markdown(
        f"""
        <h4 style="text-align:center;">
            <b>Total lanes to be negotiated (Non-Private Fleet):</b> {overall_count:,} &nbsp;&nbsp; | &nbsp;&nbsp;
            <b>Total savings potential (Non-Private Fleet):</b> ${overall_total:,.2f}
        </h4>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(summary_df, width='stretch')

    st.markdown("---")
    st.markdown("### Savings Consistency Check (Non-Private Fleet)")
    st.markdown("Savings check ensures that savings identified for RFP lanes and negotiation letter lanes are equal to the total savings identified in the original rate comparison. In other words, this check ensures no lanes are double counted or unaccounted for in the output.")
    st.markdown(
        f"- Total savings potential (Summary, non-Private Fleet): **${overall_total:,.2f}**  \n"
        f"- Savings from Vendor Letters (non-Private Fleet lanes): **${letter_savings:,.2f}**  \n"
        f"- Savings from RFP Lanes (non-Private Fleet lanes): **${rfp_savings:,.2f}**  \n"
        f"- Letters + RFP = **${combined_savings:,.2f}**"
    )

    if abs(savings_diff) > 1e-6:
        st.error(
            f"⚠️ Savings mismatch detected: Summary total minus (Letters + RFP) = "
            f"${savings_diff:,.2f}. Check classification logic or filters."
        )
    else:
        st.success("✅ Savings check passed: Letters + RFP savings equals the Summary total (non-Private Fleet).")

with tab2:  # your RFP tab
    if rfp_df.empty:
        st.info("No lanes are currently flagged for RFP under the configured rules.")
    else:
        # ---- Summary of savings potential for RFP lanes ----
        rfp_negotiate_mask = rfp_df["action"] == "NEGOTIATE"
        rfp_lane_count = int(rfp_negotiate_mask.sum())
        rfp_total_savings = float(rfp_df.loc[rfp_negotiate_mask, "delta"].sum(skipna=True))

        st.markdown(
            f"""
            <h4 style="text-align:center;">
                📦 <b>RFP lanes to be negotiated:</b> {rfp_lane_count:,} &nbsp;&nbsp; | &nbsp;&nbsp;
                <b>Total savings potential (company - benchmark):</b> ${rfp_total_savings:,.2f}
            </h4>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("**RFP Candidate Lanes (showing first 1,000 rows):**")
        base_display_cols = [
            "_lane",
            "carrier_name",
            "company_linehaul",
            "company_fuel_cost",
            "company_cost",
            "benchmark_linehaul",
            "benchmark_fuel_cost",
            "benchmark_cost",
            "delta_linehaul",
            "delta_fuel",
            "delta",
            "delta_pct",
            "carrier_count",
            "lane_treatment",
            "action",
        ]
        od_cols = ["origin_city", "origin_state", "dest_city", "dest_state"]
        rfp_display_cols = base_display_cols.copy()
        # Check whether ALL origin/dest values are blank / NaN
        show_od = False
        if all(col in rfp_df.columns for col in od_cols):
            # Treat empty strings as NaN for this check
            od_block = rfp_df[od_cols].replace("", pd.NA)
            if not od_block.isna().all().all():
                show_od = True

        if show_od:
            rfp_display_cols = od_cols + base_display_cols
        else:
            rfp_display_cols = base_display_cols
        st.dataframe(
            rfp_df[[c for c in rfp_display_cols if c in rfp_df.columns]].head(1000),
            width='stretch'
        )

        # ---- Download button for just RFP lanes ----
        st.markdown("### Download RFP Lanes")

        # Optionally subset columns for RFP template
        rfp_export_cols = [
            "origin_city", "origin_state", "dest_city", "dest_state",
            "_lane", "carrier_name",
            "benchmark_cost", "company_cost",
            "delta", "delta_pct",
            "carrier_count", "lane_treatment", "action",
        ]
        rfp_export_view = rfp_df[[c for c in rfp_export_cols if c in rfp_df.columns]].copy()

        rfp_csv = rfp_export_view.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="⬇️ Download RFP lanes as CSV",
            data=rfp_csv,
            file_name="rfp_lanes.csv",
            mime="text/csv",
        )

with tab3:

    st.markdown(
        f"""
        <h4 style="text-align:center;">
            💼 <b>Total Lanes to be Negotiated (Letters only, non-Private Fleet) {letter_lane_count:,}
            &nbsp;&nbsp; | &nbsp;&nbsp;
            <b>Total savings potential (Letters only, non-Private Fleet):</b> ${letter_savings:,.2f}
        </h4>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("**Lanes to address via vendor letters (excluding Private Fleet):**")
    st.dataframe(letter_df, width='stretch')

with tab4:
    if gpf_export.empty:
        st.info("No PRIVATE FLEET lanes found after exclusions.")
    else:
        st.caption(f"{gpf_count} Private Fleet lanes | NEGOTIATE delta: ${gpf_total_delta:,.2f}")
        st.dataframe(gpf_export, width='stretch')

with tab5:
    st.dataframe(excluded_summary_df, width='stretch')

with tab6:
    st.dataframe(
        excluded_detail_df if not excluded_detail_df.empty
        else pd.DataFrame(columns=["_lane", "carrier_name", "excluded_matches"]),
        width='stretch'
    )

# ============ Download Excel ============
def build_comparison_workbook():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="All_NonPrivateFleet", index=False)
        letter_df.to_excel(writer, sheet_name="Letter_Lanes", index=False)
        rfp_df.to_excel(writer, sheet_name="RFP_Lanes", index=False)
        gpf_export.to_excel(writer, sheet_name="Private_Fleet", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        excluded_summary_df.to_excel(writer, sheet_name="Excluded_Summary", index=False)
        (excluded_detail_df if not excluded_detail_df.empty else
            pd.DataFrame(columns=["_lane","carrier_name","excluded_matches"])) \
            .to_excel(writer, sheet_name="Excluded_Detail", index=False)
    buf.seek(0)
    return buf.getvalue()

if st.button("Prepare Excel comparison workbook"):
    st.session_state["comparison_xlsx"] = build_comparison_workbook()

if "comparison_xlsx" in st.session_state:
    st.download_button(
        label="⬇️ Download Excel (Comparison + Private Fleet + Summary + Exclusions)",
        data=st.session_state["comparison_xlsx"],
        file_name="lane_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ------------------ RFP Template Controls ------------------
st.markdown("---")
st.markdown("""
#### 📝 RFP Template & Letters

Use this section when you are planning an **RFP event**:

The **RFP Template (Excel)** contains:
  - A cover sheet with overview, instructions, and contact info
  - A bid timeline
  - A tab of RFP lanes with **Market Rate** and fields for carrier input
  - A location list tab

The **RFP letters (by carrier)**:
  - Reference the total number of lanes in the RFP
  - Highlight the lanes that carrier currently services above benchmark
  - Show the **top 10% lanes with the largest variance**, plus summaries. Please note that the top 10% threshold cannot be adjusted. If there are certainly lanes to be excluded from the summary, please ensure they are entered in the 'RFP overrides' section above.
""")
st.subheader("📦 Build RFP Template (Excel)")

# Overview text (Tab 1)
overview_text = st.text_area(
    "RFP overview text (Tab 1): Please **revise the text below** to be specific to your company. Any edits you make here will be reflected in the RFP template.",
    value="Greif, Inc. is a global manufacturer of industrial packaging products" 
        "and services, supplying fiber, steel, and plastic drums; intermediate " 
        "bulk containers; corrugated and paper-based packaging; and containerboard " 
        "to customers worldwide. The company also offers packaging accessories, " 
        "transit protection products, and a range of reconditioning, recycling, " 
        "and supply-chain services that support Circular Economy and sustainability initiatives. "
        "Greif serves major industries including chemicals, pharmaceuticals, agriculture, "
        "food and beverage, petroleum, and general manufacturing. With a broad production "
        "footprint and a strong commitment to operational excellence, product quality, and " 
        "environmental stewardship, Greif is a trusted partner for safe, reliable, and "
        "efficient packaging solutions. Founded in 1877, the company is headquartered in Delaware, Ohio.",
    help="This text will populate the Overview tab in the RFP."
)

instructions_text = st.text_area(
    "RFP Instructions (Tab 1): Please **revise the text below** to be specific to your company. Any edits you make here will be reflected in the RFP template.",
    value="Greif is continously looking to improve our ways of working and unlock efficiencies across our Supply Chain. "
        'Therefore, we are leveraging our volume and knowledge of market rates to complete a bid for several of our lanes and '
        'request that your company participate. \n'
        'In this document you will find three tabs:\n(1) Bid timeline - expected due dates and responses dates\n'
        '(2) All lanes that are open to bid - please enter bid cost as base rate charge and fuel surcharge as percentage of base rate'
        '\n(3) Locations covered by lanes included in this RFP - for reference only, no action required\n '
        'We request you complete and return this RFP by no later than {reply_by}.',
    height=200
)

instruction_lines = [line for line in instructions_text.split("\n") if line.strip()]

# Timeline text (Tab 2)
timeline_text = st.text_area(
    "Bid timeline (Tab 2) Please **revise the text below** to be specific to your company's RFP timeline. Any edits you make here will be reflected in the RFP template. Please enter as one line per milestone, format: Date, Milestone",
    value="4/22, Bid Release\n5/1, Round 1 Carrier Offers Due\n5/15, Round 1 "
    "[Company Name] Feedback to Carriers\n5/29, Round 2 Carrier Offers Due\n6/7, "
    "Round 2 [Company Name] Feedback to Carriers & Final Negotiations\n6/13, Final Awards",
    help="Each line should have a date and a milestone, separated by a comma."
)

# Location list upload (Tab 4)
rfp_loc_file = st.file_uploader(
    "Upload location list for RFP (optional – becomes last tab)",
    type=["xlsx", "xls", "csv"],
    key="rfp_locs"
)

# ---------- Build dataframes used in the export ----------


# Tab 2: Timeline -> Date + Milestone
timeline_rows = []
for raw in timeline_text.split("\n"):
    if not raw.strip():
        continue
    if "," in raw:
        date_str, milestone = raw.split(",", 1)
        timeline_rows.append({"Date": date_str.strip(), "Milestone": milestone.strip()})
    else:
        timeline_rows.append({"Date": raw.strip(), "Milestone": ""})

if timeline_rows:
    timeline_df = pd.DataFrame(timeline_rows)
else:
    timeline_df = pd.DataFrame(columns=["Date", "Milestone"])

# Tab 3: Bid Lanes -> group rfp_df by lane and compute count + market rate
lane_cols = ["origin_city", "origin_state", "dest_city", "dest_state"]

if not rfp_df.empty:
    rfp_group = (
        rfp_df
        .groupby(lane_cols, dropna=False)
        .agg(
            Shipment_Count=("_lane", "size"),
            Market_Rate=("benchmark_cost", "mean")
        )
        .reset_index()
    )

    bid_lanes_export = pd.DataFrame({
        "Zip ID": range(1, len(rfp_group) + 1),
        "Move Type": "US Domestic",
        "Origin City": rfp_group["origin_city"],
        "Origin State": rfp_group["origin_state"],
        "Origin Country": "US",
        "Destination City": rfp_group["dest_city"],
        "Destination State": rfp_group["dest_state"],
        "Destination Country": "US",
        "Shipment Count": rfp_group["Shipment_Count"],
        "Market Rate": rfp_group["Market_Rate"],  # benchmark_cost
        "Flat rate (before fuel surcharge)": pd.NA,
        "Fuel surcharge (as percentage of flat rate)": pd.NA,
    })
else:
    bid_lanes_export = pd.DataFrame(
        columns=[
            "Zip ID", "Move Type",
            "Origin City", "Origin State", "Origin Country",
            "Destination City", "Destination State", "Destination Country",
            "Shipment Count", "Market Rate",
            "Flat rate (before fuel surcharge)",
            "Fuel surcharge (as percentage of flat rate)",
        ]
    )

# ------------------ Export button + writer ------------------
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

if st.button("⬇️ Build RFP Template (Excel)"):

    # Location list
    if rfp_loc_file is not None:
        loc_df = read_any_cached(rfp_loc_file)
    else:
        loc_df = pd.DataFrame()

    # Prepare text lines for sections
    overview_lines = [line for line in overview_text.split("\n") if line.strip()]
    instruction_lines = [line for line in instructions_text.split("\n") if line.strip()]

    # Simple contact template (user will overwrite in Excel)
    contact_lines = [
        "Name: _______________________________",
        "Title: ______________________________",
        "Company: ____________________________",
        "Email: _____________________________",
        "Phone: _____________________________",
    ]

    rfp_buf = io.BytesIO()
    with pd.ExcelWriter(rfp_buf, engine="openpyxl") as writer:

        # Create an empty sheet for Tab 1
        pd.DataFrame().to_excel(writer, sheet_name="RFP Overview", index=False)
        ws = writer.sheets["RFP Overview"]

        # ---- Helper functions for sections ----

        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

        def apply_border(ws, start_row, end_row, start_col="B", end_col="H"):
            for row in ws[f"{start_col}{start_row}:{end_col}{end_row}"]:
                for cell in row:
                    cell.border = thin_border

        def add_text_section(ws, title, lines, start_row, fill_color="EEECE1"):
            """
            Adds a titled, shaded, bordered section starting at row `start_row`.
            Returns the next free row after the section (with a blank row spacer).
            """
            # Section title in column B
            title_cell = ws[f"B{start_row}"]
            title_cell.value = title
            title_cell.font = Font(size=14, bold=True)
            title_cell.alignment = Alignment(horizontal="left", vertical="bottom")

            # Body starts one row below title
            body_start = start_row + 1

            # Choose a reasonable body height based on number of lines
            # at least 5 rows high
            body_height = max(5, len(lines) + 2)
            body_end = body_start + body_height - 1

            # Merge body area B..H
            merge_range = f"B{body_start}:H{body_end}"
            ws.merge_cells(merge_range)

            # Put text into top-left of merged region
            cell = ws[f"B{body_start}"]
            cell.value = "\n".join(lines) if lines else ""
            cell.alignment = Alignment(
                wrap_text=True,
                vertical="top",
                horizontal="left",
            )
            cell.font = Font(size=12)

            # Shading
            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

            # Apply border to the entire merged block
            apply_border(ws, body_start, body_end)

            # Return next free row (with one blank row for spacing)
            return body_end + 2

        # ---- Column widths for readability ----
        ws.column_dimensions["A"].width = 2
        ws.column_dimensions["B"].width = 60
        for col in ["C", "D", "E", "F", "G", "H"]:
            ws.column_dimensions[col].width = 20

        # ===========================
        # SECTION 1 — OVERVIEW
        # ===========================
        current_row = 2
        current_row = add_text_section(
            ws,
            title="1. Overview",
            lines=overview_lines,
            start_row=current_row,
            fill_color="FFF2CC",  # light yellow
        )

        # ===========================
        # SECTION 2 — INSTRUCTIONS
        # ===========================
        current_row = add_text_section(
            ws,
            title="2. Instructions",
            lines=instruction_lines,
            start_row=current_row,
            fill_color="E2EFDA",  # light green
        )

        # ===========================
        # SECTION 3 — CONTACT INFORMATION
        # ===========================
        current_row = add_text_section(
            ws,
            title="3. Contact Information",
            lines=contact_lines,
            start_row=current_row,
            fill_color="D9E1F2",  # light blue
        )

        # ---- TAB 2: Bid Timeline ----
        sheet_name_timeline = "Bid Timeline"
        
        # Write timeline data starting at row 4 to leave space for title
        timeline_df.to_excel(writer, sheet_name=sheet_name_timeline, index=False, startrow=3, startcol=1)
        ws_tl = writer.sheets[sheet_name_timeline]
        
        # Thin border style
        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )
        
        # ---- Title at top ----
        title_cell = ws_tl["B2"]
        title_cell.value = "Bid Timeline"
        title_cell.font = Font(size=14, bold=True)
        title_cell.alignment = Alignment(horizontal="left", vertical="bottom")
        
        # Determine data range
        start_row = 4  # where we wrote the header
        end_row = start_row + len(timeline_df)  # header + rows
        start_col = "B"
        end_col = chr(ord(start_col) + len(timeline_df.columns) - 1)  # e.g. B + 1 -> C if 2 columns
        
        # Header row styling (light shading, bold, centered)
        header_row = ws_tl[start_row]
        for cell in header_row:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        # Data row styling (wrapped, left-aligned for Milestone)
        for row in ws_tl.iter_rows(min_row=start_row+1, max_row=end_row):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        
        # Apply border & shading around the full table block (header + data)
        for row in ws_tl[f"{start_col}{start_row}:{end_col}{end_row}"]:
            for cell in row:
                cell.border = thin_border
        
        # Light background for the whole table area (optional)
        for row in ws_tl[f"{start_col}{start_row}:{end_col}{end_row}"]:
            for cell in row:
                if cell.row == start_row:
                    continue  # header already shaded
                cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        
        # ---- Autosize columns so text fits ----
        for col_cells in ws_tl.iter_cols(min_row=start_row, max_row=end_row, min_col=2, max_col=1+len(timeline_df.columns)):
            max_length = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                try:
                    val = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(val))
                except Exception:
                    pass
            # Add some padding
            ws_tl.column_dimensions[col_letter].width = max_length + 4

        # ---- REMOVE ANY FORMATTING FROM COLUMN A ----
        for cell in ws_tl["A:A"]:
            cell.border = Border()  # remove border
            cell.fill = PatternFill(fill_type=None)  # remove background
            cell.font = Font(size=11)  # standard font
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
        ws_tl.column_dimensions["A"].width = 2  # narrow blank column

        # ---- TAB 3: Bid Lanes ----
        sheet_name_lanes = "Bid Lanes"

        # Write the DataFrame starting at row 4 (leaving room for title)
        bid_lanes_export.to_excel(writer, sheet_name=sheet_name_lanes, index=False, startrow=3, startcol=1)
        ws_bl = writer.sheets[sheet_name_lanes]
        
        # Borders
        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )
        
        # ---- Title ----
        title_cell_bl = ws_bl["B2"]
        title_cell_bl.value = "Bid Lanes"
        title_cell_bl.font = Font(size=14, bold=True)
        title_cell_bl.alignment = Alignment(horizontal="left", vertical="bottom")
        
        # Determine data range
        start_row_bl = 4  # header row from to_excel(startrow=3)
        end_row_bl = start_row_bl + len(bid_lanes_export)  # header + data rows
        start_col_idx = 2  # column B
        num_cols = len(bid_lanes_export.columns)
        end_col_idx = start_col_idx + num_cols - 1
        
        # Map column index -> header text
        header_cells = ws_bl[start_row_bl]
        col_idx_to_name = {}
        for cell in header_cells:
            col_idx_to_name[cell.column] = str(cell.value)
        
        # Define which columns are carrier-fill
        carrier_fill_cols = {
            "Flat rate (before fuel surcharge)",
            "Fuel surcharge (as percentage of flat rate)",
        }
        
        # Some colors
        header_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # light yellow
        data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")    # white
        carrier_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # light blue for carrier inputs
        
        # ---- Style header row ----
        for cell in header_cells:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill
            cell.border = thin_border
        
        # ---- Style data rows ----
        for row in ws_bl.iter_rows(min_row=start_row_bl+1, max_row=end_row_bl,
                                   min_col=start_col_idx, max_col=end_col_idx):
            for cell in row:
                header_name = col_idx_to_name.get(cell.column, "")
                # Carrier-fill columns get special shading
                if header_name in carrier_fill_cols:
                    cell.fill = carrier_fill
                else:
                    cell.fill = data_fill
        
                cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                cell.border = thin_border
        
        # ---- Autosize columns so text fits ----
        for col_idx in range(start_col_idx, end_col_idx + 1):
            col_letter = ws_bl.cell(row=start_row_bl, column=col_idx).column_letter
            max_length = 0
            for row in range(start_row_bl, end_row_bl + 1):
                cell = ws_bl.cell(row=row, column=col_idx)
                try:
                    val = str(cell.value) if cell.value is not None else ""
                    max_length = max(max_length, len(val))
                except Exception:
                    pass
        
            # Add padding; wider for descriptive columns
            if col_idx == start_col_idx:  # e.g., Zip ID
                ws_bl.column_dimensions[col_letter].width = max(10, max_length + 2)
            else:
                ws_bl.column_dimensions[col_letter].width = max(15, max_length + 4)
        
        # ---- REMOVE ANY FORMATTING FROM COLUMN A ----
        for cell in ws_bl["A:A"]:
            cell.border = Border()
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font(size=11)
            cell.alignment = Alignment(horizontal="left", vertical="center")
        
        ws_bl.column_dimensions["A"].width = 2

        # ---- TAB 4: Location List ----
        if not loc_df.empty:
            loc_df.to_excel(writer, sheet_name="Company Location List", index=False)
        else:
            pd.DataFrame(
                columns=[
                    "Location Name", "Street Address", "City", "State", "Zip",
                    "Location Type", "Ledger Flag", "Platform"
                ]
            ).to_excel(writer, sheet_name="Company Location List", index=False)

    rfp_buf.seek(0)

    st.download_button(
        label="📥 Download RFP Template (Overview + Instructions + Contacts + Timeline + Bid Lanes + Locations)",
        data=rfp_buf,
        file_name="RFP_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


st.markdown("---")
st.subheader("✉️ Generate RFP Letters (by Carrier)")
st.markdown("""
Use negotiation letters when you **do not want a full RFP** for these lanes.

This section will:
  - Use lanes marked as **Letter** in the comparison logic
  - Group lanes by carrier
  - Show **unique lanes**, their **frequency**, **company vs benchmark cost**, and **% over benchmark**
Letters are best for:
  - Smaller numbers of lanes
  - One-off rate reviews
  - Strategic/relationship-based negotiations
  - Month-to-month rate comparisons to ensure carriers moving towards compliance with benchmark rates
""")
# Sender info can reuse the same as negotiation section, or separate:
col_a, col_b, col_c = st.columns(3)
with col_a:
    rfp_sender_company = st.text_input("Your company name: Please enter the name of your company", value="Greif")
with col_b:
    rfp_sender_name = st.text_input("Your name: Please enter your name to be included in signature line", value="Your Name")
with col_c:
    rfp_sender_title = st.text_input("Your title: Please enter your title to be included in signature line", value="Procurement Manager")

rfp_reply_by = st.text_input("RFP reply-by date: Please enter the deadline for vendors to respond", value="")

# Default body text from your prompt
rfp_default_body = (
    "Good afternoon,\n\n"
    "We are reaching out as we are completing an RFP for our Freight spend. "
    "The RFP that has been shared with you has a total {total_rfp_lanes} lanes. "
    "We appreciate your partnership and would like to continue leveraging your company as a strategic carrier.\n\n "
    "Of the lanes that are included in this RFP, please note there are {carrier_lane_count} "
    "lanes you service that are currently above benchmark rates at "
    "an average of {avg_over_pct} above benchmark. Our goal is to close the gap to benchmark "
    "rates as much as possible and would appreciate your careful consideration as you complete this RFP\n\n"
    "For your reference and ease, the top 10% lanes with the greatest delta to benchmark are listed below with a summary of remaining lanes. "
    "If you would like further detail on the remaining lanes, please reach out.\n\n"
    "Please submit your responses by {reply_by}."
)

rfp_body_template = st.text_area(
    "RFP letter body: Please **revise the text below** to be specific to your company. Any edits you make here will be reflected in the RFP letters.",
    value=rfp_default_body,
    height=220,
)

include_privfleet_in_rfp_letters = st.checkbox(
    "Include PRIVATE FLEET in RFP letters",
    value=False
)

if st.button("Build RFP letters (ZIP)"):
    if rfp_df.empty:
        st.warning("No RFP lanes available to create letters.")
    else:
        zip_bytes = build_rfp_letters_zip(
            rfp_df=rfp_df,
            sender_company=rfp_sender_company,
            sender_name=rfp_sender_name,
            sender_title=rfp_sender_title,
            reply_by=rfp_reply_by if rfp_reply_by else "the requested RFP due date",
            rfp_body_template=rfp_body_template,
            include_privfleet=include_privfleet_in_rfp_letters,
        )
        st.download_button(
            label="⬇️ Download RFP letters (ZIP)",
            data=zip_bytes,
            file_name="rfp_letters.zip",
            mime="application/zip"
        )

st.markdown("---")
st.subheader("✉️ Generate Negotiation Letters")
# ---- determine which lanes are actually in-scope for negotiation letters ----
neg_letter_df = letter_df[letter_df["action"] == "NEGOTIATE"].copy()

if neg_letter_df.empty:
    st.info(
        "No lanes are currently flagged for negotiation letters. "
        "Under the current rules and overrides, all lanes above benchmark "
        "are either routed to the RFP or excluded."
    )
else:
    st.success(
        f"{neg_letter_df['carrier_name'].nunique()} carriers and "
        f"{neg_letter_df['_lane'].nunique()} unique lanes "
        "are in scope for negotiation letters."
    )
# Combine only Letter lanes (non-Private Fleet) + all other lanes lanes (if included)
# Combine only Letter lanes (non-Private Fleet) + optional GREIF lanes
combined_for_letters = neg_letter_df.assign(source="NON_GREIF")
# GREIF lanes will be added later inside the button when (and if) user checks the box

col_a, col_b, col_c = st.columns(3)
with col_a:
    sender_company = st.text_input("Your company name: Please enter your company name", value="Greif")
with col_b:
    sender_name = st.text_input("Your name: Please enter your name which will be included in signature line", value="Your Name")
with col_c:
    sender_title = st.text_input("Your title: Please enter your title which will be included in signature line", value="Procurement Manager")

col_d, col_e = st.columns(2)
with col_d:
    reply_by = st.text_input("Reply-by date (e.g., 2025-01-15):  Please enter the deadline for vendors to respond", value="")
with col_e:
    include_privfleet_in_letters = st.checkbox("Include PRIVATE FLEET in letters", value=False)

# --- NEW: editable intro template ---
default_intro = (
    "We appreciate your service and would like to continue our partnership. "
    "We have identified {num_lanes} lanes where the rates we are charged are "
    "above market by an average of {avg_over_pct} and would like to review rates. "
    "Please review and provide your best offer by {reply_by}."
)

st.markdown("**Negotiation letter body**")
st.caption(
    "Available placeholders: "
    "`{num_lanes}` = number of lanes in the letter, "
    "`{avg_over_pct}` = average %% above market, "
    "`{reply_by}` = reply-by date."
)

letter_body_template = st.text_area(
    "Negotiation letter body: Please **revise the text below** to be specific to your company. Any edits you make here will be reflected in the RFP template",
    value=default_intro, 
    height=120,
)

if st.button("Build negotiation letters (ZIP)"):

    # Start from non-GREIF negotiation lanes
    combined_for_letters = neg_letter_df.assign(source="NON_GREIF")

    # Optionally add GREIF lanes that are NEGOTIATE
    if include_greif_in_letters and not gpf_export.empty:
        greif_neg = gpf_export[gpf_export["action"] == "NEGOTIATE"].copy()
        greif_neg["source"] = "GREIF"
        combined_for_letters = pd.concat(
            [combined_for_letters, greif_neg],
            ignore_index=True,
            sort=False,
        )

    if combined_for_letters.empty:
        st.warning(
            "No negotiation letters were generated because no lanes are "
            "currently classified as Letter + NEGOTIATE."
        )
    else:
        zip_bytes = build_letters_zip(
            df_all=combined_for_letters,
            include_greif=include_greif_in_letters,
            sender_company=sender_company,
            sender_name=sender_name,
            sender_title=sender_title,
            reply_by=reply_by if reply_by else "7 business days from receipt",
            body_template=letter_body_template,
        )
        st.download_button(
            label="⬇️ Download negotiation letters (ZIP)",
            data=zip_bytes,
            file_name="negotiation_letters.zip",
            mime="application/zip"
        )
