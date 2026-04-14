import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Third-Party Risk Assessment",
    page_icon="🔐",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .metric-card {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 1rem 1.25rem;
        text-align: center;
    }
    .stAlert { border-radius: 8px; }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #dee2e6;
        border-radius: 12px;
        padding: 1.5rem;
    }
    .posture-label {
        font-size: 0.8rem;
        color: #6c757d;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
SEVERITY_COLORS  = {"High": "#E24B4A", "Medium": "#EF9F27", "Low": "#639922"}
SEVERITY_ORDER   = ["High", "Medium", "Low"]
BADGE_STYLES     = {
    "High":   "background:#FCEBEB;color:#A32D2D;padding:2px 10px;border-radius:6px;font-size:12px;font-weight:500",
    "Medium": "background:#FAEEDA;color:#854F0B;padding:2px 10px;border-radius:6px;font-size:12px;font-weight:500",
    "Low":    "background:#EAF3DE;color:#3B6D11;padding:2px 10px;border-radius:6px;font-size:12px;font-weight:500",
}

# ── Column name aliases (flexible detection) ──────────────────────────────────
CAT_ALIASES  = ["category", "domain", "area", "control area", "control domain"]
RISK_ALIASES = ["risk", "risk description", "description", "finding", "issue", "observation"]
SEV_ALIASES  = ["severity", "level", "risk level", "rating", "priority", "criticality"]
VDR_ALIASES  = ["vendor", "vendor name", "company", "supplier"]
QST_ALIASES  = ["questionnaire", "assessment", "framework", "standard"]

def find_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    lower_map = {c.lower().strip(): c for c in df.columns}
    for a in aliases:
        if a in lower_map:
            return lower_map[a]
    return None

# ── Excel parser ──────────────────────────────────────────────────────────────
def parse_excel(file) -> tuple[pd.DataFrame, dict]:
    """Read uploaded Excel, auto-detect columns, return cleaned DataFrame + metadata."""
    try:
        xl = pd.ExcelFile(file)
        sheet = xl.sheet_names[0]
        raw = xl.parse(sheet)
        raw.columns = raw.columns.str.strip()
    except Exception as e:
        st.error(f"Could not read file: {e}")
        return None, {}

    cat_col  = find_col(raw, CAT_ALIASES)
    risk_col = find_col(raw, RISK_ALIASES)
    sev_col  = find_col(raw, SEV_ALIASES)

    missing = []
    if not cat_col:  missing.append("Category")
    if not risk_col: missing.append("Risk / Description")
    if not sev_col:  missing.append("Severity / Level")

    if missing:
        st.error(
            f"**Missing required columns:** {', '.join(missing)}\n\n"
            f"Columns found in file: `{'`, `'.join(raw.columns.tolist())}`\n\n"
            "Rename your columns to **Category**, **Risk**, and **Severity** and re-upload."
        )
        return None, {}

    df = pd.DataFrame({
        "Category": raw[cat_col].astype(str).str.strip(),
        "Risk":     raw[risk_col].astype(str).str.strip(),
        "Severity": raw[sev_col].astype(str).str.strip().str.capitalize(),
    })
    df = df[df["Risk"].notna() & (df["Risk"] != "") & (df["Risk"] != "Nan")]
    df = df[df["Severity"].isin(SEVERITY_ORDER)]
    df = df.reset_index(drop=True)

    meta = {}
    vdr_col = find_col(raw, VDR_ALIASES)
    qst_col = find_col(raw, QST_ALIASES)
    if vdr_col: meta["vendor"] = str(raw[vdr_col].dropna().iloc[0]) if not raw[vdr_col].dropna().empty else ""
    if qst_col: meta["questionnaire"] = str(raw[qst_col].dropna().iloc[0]) if not raw[qst_col].dropna().empty else ""
    meta["sheet"] = sheet
    meta["filename"] = getattr(file, "name", "uploaded file")

    return df, meta

# ── Posture score ─────────────────────────────────────────────────────────────
def posture_score(df: pd.DataFrame) -> int:
    if df.empty: return 0
    weights = {"High": 0, "Medium": 0.5, "Low": 1.0}
    total   = len(df)
    earned  = sum(weights.get(s, 0) for s in df["Severity"])
    return round((earned / total) * 100)

# ── Sample Excel download ─────────────────────────────────────────────────────
def make_sample_excel() -> bytes:
    sample = pd.DataFrame([
        {"Category": "Application Security", "Risk": "No SAST/DAST tools in CI/CD pipeline.",                        "Severity": "Medium", "Vendor": "Acme Corp", "Questionnaire": "TPCRA Part 2"},
        {"Category": "Application Security", "Risk": "No SBOM or open-source vulnerability tracking (e.g. Snyk).",   "Severity": "Medium"},
        {"Category": "Identity & Access Management", "Risk": "No dedicated PAM solution (CyberArk, BeyondTrust).",    "Severity": "Medium"},
        {"Category": "Identity & Access Management", "Risk": "MFA uses TOTP only — no FIDO2/WebAuthn for admins.",    "Severity": "Low"},
        {"Category": "Data Protection",       "Risk": "No dedicated DLP tool — relies on layered controls.",          "Severity": "Medium"},
        {"Category": "Data Protection",       "Risk": "TLS 1.2 in transit — TLS 1.3 not confirmed.",                  "Severity": "Low"},
        {"Category": "Incident Response",     "Risk": "48-hr notification SLA may conflict with GDPR 72-hr rule.",    "Severity": "Medium"},
        {"Category": "Incident Response",     "Risk": "No tabletop exercises or red-team simulation on record.",       "Severity": "Low"},
        {"Category": "Business Continuity",   "Risk": "No RTO/RPO targets defined in BCP.",                           "Severity": "Medium"},
        {"Category": "Business Continuity",   "Risk": "BCP test results not shared — effectiveness unverifiable.",    "Severity": "Low"},
        {"Category": "AI Controls",           "Risk": "No AI Controls section — AI/ML usage completely unassessed.",  "Severity": "High"},
        {"Category": "Infrastructure Security","Risk": "No SLA for non-critical patch windows.",                       "Severity": "Low"},
        {"Category": "Cloud Services",        "Risk": "Shared responsibility model with AWS not documented.",          "Severity": "Low"},
    ])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Risk Assessment")
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

# ── Header ────────────────────────────────────────────────────────────────────
st.title("🔐 Third-Party Risk Assessment")
st.caption("Upload a vendor questionnaire spreadsheet to generate the risk dashboard.")
st.divider()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Upload")
    uploaded = st.file_uploader(
        "Questionnaire Excel file",
        type=["xlsx", "xls"],
        help="First sheet is used. Required columns: Category, Risk, Severity.",
    )
    st.divider()
    st.subheader("Required columns")
    st.markdown("""
| Column | Accepted names |
|--------|---------------|
| Category | `Category`, `Domain`, `Area` |
| Risk | `Risk`, `Description`, `Finding` |
| Severity | `Severity`, `Level`, `Rating` |
| Vendor *(opt.)* | `Vendor`, `Company` |
    """)
    st.divider()
    st.download_button(
        "⬇ Download sample Excel",
        data=make_sample_excel(),
        file_name="sample_risk_assessment.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Use the sample to see the expected format.")

# ── Empty state ───────────────────────────────────────────────────────────────
if not uploaded:
    st.info("👈 Upload an Excel file from the sidebar to get started.", icon="📂")
    st.markdown("#### What this dashboard shows")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Risk metrics**\n\nTotal, High, Medium, and Low risk counts with an overall security posture score.")
    with c2:
        st.markdown("**Visual breakdown**\n\nStacked bar chart by category and severity distribution donut chart.")
    with c3:
        st.markdown("**Filterable table**\n\nAll risks with severity badges — filterable by severity and category.")
    st.stop()

# ── Parse ─────────────────────────────────────────────────────────────────────
df, meta = parse_excel(uploaded)
if df is None or df.empty:
    st.warning("No valid risk rows found. Check that Severity values are High / Medium / Low.")
    st.stop()

# ── Metadata banner ───────────────────────────────────────────────────────────
vendor = meta.get("vendor", "")
qst    = meta.get("questionnaire", "")
fname  = meta.get("filename", "")

banner_parts = [f"**{vendor}**" if vendor else None, qst if qst else None, f"`{fname}`"]
st.success("  ·  ".join(p for p in banner_parts if p))

# ── Summary ───────────────────────────────────────────────────────────────────
counts  = df["Severity"].value_counts()
n_high  = int(counts.get("High",   0))
n_med   = int(counts.get("Medium", 0))
n_low   = int(counts.get("Low",    0))
n_total = len(df)
score   = posture_score(df)

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total risks",  n_total)
col2.metric("🔴 High",      n_high)
col3.metric("🟠 Medium",    n_med)
col4.metric("🟢 Low",       n_low)
col5.metric("Posture score", f"{score}%",
    delta="Strong" if score >= 70 else ("Moderate" if score >= 40 else "Weak"),
    delta_color="normal" if score >= 70 else ("off" if score >= 40 else "inverse"),
)
st.divider()

# ── Charts ─────────────────────────────────────────────────────────────────────
chart_l, chart_r = st.columns([3, 2])

with chart_l:
    st.subheader("Risks by category")
    cat_df = (
        df.groupby(["Category", "Severity"])
        .size()
        .reset_index(name="Count")
    )
    fig_bar = px.bar(
        cat_df,
        x="Count", y="Category", color="Severity",
        orientation="h",
        color_discrete_map=SEVERITY_COLORS,
        category_orders={"Severity": SEVERITY_ORDER},
        labels={"Count": "Number of risks", "Category": ""},
        height=max(320, len(df["Category"].unique()) * 44),
    )
    fig_bar.update_layout(
        margin=dict(l=0, r=0, t=10, b=0),
        legend_title_text="Severity",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        yaxis={"categoryorder": "total ascending"},
        font=dict(size=12),
    )
    fig_bar.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)")
    st.plotly_chart(fig_bar, use_container_width=True)

with chart_r:
    st.subheader("Severity distribution")
    sev_counts = (
        df["Severity"].value_counts()
        .reindex(SEVERITY_ORDER)
        .fillna(0)
        .reset_index()
    )
    sev_counts.columns = ["Severity", "Count"]
    fig_donut = go.Figure(go.Pie(
        labels=sev_counts["Severity"],
        values=sev_counts["Count"],
        hole=0.65,
        marker_colors=[SEVERITY_COLORS[s] for s in sev_counts["Severity"]],
        textinfo="label+percent",
        hovertemplate="%{label}: %{value}<extra></extra>",
    ))
    fig_donut.update_layout(
        margin=dict(l=0, r=0, t=10, b=0),
        showlegend=False,
        height=320,
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(size=12),
    )
    st.plotly_chart(fig_donut, use_container_width=True)

st.divider()

# ── Risk table ────────────────────────────────────────────────────────────────
st.subheader("All risks")

f1, f2 = st.columns([2, 3])
with f1:
    sev_filter = st.multiselect(
        "Filter by severity",
        options=SEVERITY_ORDER,
        default=SEVERITY_ORDER,
    )
with f2:
    cat_filter = st.multiselect(
        "Filter by category",
        options=sorted(df["Category"].unique()),
        default=sorted(df["Category"].unique()),
    )

filtered = df[
    df["Severity"].isin(sev_filter) &
    df["Category"].isin(cat_filter)
].reset_index(drop=True)

st.caption(f"Showing {len(filtered)} of {n_total} risks")

# Styled HTML table
rows_html = ""
for _, row in filtered.iterrows():
    badge = f'<span style="{BADGE_STYLES.get(row["Severity"], "")}">{row["Severity"]}</span>'
    rows_html += f"""
    <tr>
      <td style="padding:10px 12px;border-bottom:1px solid #f0f0f0;color:#333;font-size:13px;width:20%;vertical-align:top">{row['Category']}</td>
      <td style="padding:10px 12px;border-bottom:1px solid #f0f0f0;color:#333;font-size:13px;width:65%;vertical-align:top;line-height:1.5">{row['Risk']}</td>
      <td style="padding:10px 12px;border-bottom:1px solid #f0f0f0;font-size:13px;width:15%;vertical-align:top">{badge}</td>
    </tr>"""

table_html = f"""
<div style="border:1px solid #e9ecef;border-radius:10px;overflow:hidden;margin-top:8px">
  <table style="width:100%;border-collapse:collapse;table-layout:fixed">
    <thead>
      <tr style="background:#f8f9fa">
        <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef;width:20%">Category</th>
        <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef;width:65%">Risk description</th>
        <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef;width:15%">Severity</th>
      </tr>
    </thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>"""

st.markdown(table_html, unsafe_allow_html=True)

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()
exp_col1, exp_col2, _ = st.columns([1, 1, 3])

with exp_col1:
    csv = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇ Export filtered CSV",
        data=csv,
        file_name="risk_assessment_filtered.csv",
        mime="text/csv",
        use_container_width=True,
    )

with exp_col2:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        filtered.to_excel(writer, index=False, sheet_name="Filtered Risks")
    st.download_button(
        "⬇ Export filtered Excel",
        data=buf.getvalue(),
        file_name="risk_assessment_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.caption("Third-Party Cyber Risk Assessment Dashboard · TPCRA")
