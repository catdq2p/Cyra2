import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import datetime
from openpyxl import load_workbook

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TPCRA Dashboard",
    page_icon="🔐",
    layout="wide",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
.section-header {
    font-size: 11px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.07em; color: #6c757d; margin: 0 0 8px 0;
}
.response-badge {
    display: inline-block; padding: 2px 9px; border-radius: 6px;
    font-size: 12px; font-weight: 500;
}
.stTabs [data-baseweb="tab"] { font-size: 13px; }
div[data-testid="metric-container"] { background: #f8f9fa; border-radius: 10px; padding: 12px; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
# Section letter → readable name
SECTION_MAP = {
    "A": "Organizational Management",
    "B": "Human Resource Management",
    "C": "Infrastructure Security",
    "D": "Data Protection",
    "E": "Access Management",
    "F": "Application Security",
    "G": "System Security",
    "H": "Email Security",
    "I": "Mobile Devices",
    "J": "Incident Response",
    "K": "Cloud Services",
    "L": "Business Continuity",
}

# Compliance responses
YES_VALS  = {"yes", "y"}
NO_VALS   = {"no", "n"}
NA_VALS   = {"n/a", "na", "not applicable"}
PART_VALS = {"partial", "partly", "partially"}

BADGE_CSS = {
    "Yes":     "background:#EAF3DE;color:#3B6D11",
    "No":      "background:#FCEBEB;color:#A32D2D",
    "Partial": "background:#FAEEDA;color:#854F0B",
    "N/A":     "background:#F1EFE8;color:#5F5E5A",
    "—":       "background:#F1EFE8;color:#5F5E5A",
}

COLORS = {
    "Yes":     "#639922",
    "No":      "#E24B4A",
    "Partial": "#EF9F27",
    "N/A":     "#B4B2A9",
}


def normalize_response(val) -> str:
    if val is None:
        return "—"
    if isinstance(val, datetime.datetime):
        return val.strftime("%-m/%-d/%Y")
    s = str(val).strip().lower()
    if s in YES_VALS:  return "Yes"
    if s in NO_VALS:   return "No"
    if s in NA_VALS:   return "N/A"
    if s in PART_VALS: return "Partial"
    return str(val).strip()


def get_section(key) -> str:
    """Extract section letter from a key like 'A.1', 'B', 'C.2.1'."""
    if key is None:
        return ""
    s = str(key).strip()
    if s and s[0].isalpha():
        return s[0].upper()
    return ""


# ── Parser ────────────────────────────────────────────────────────────────────
def parse_tpcra(file) -> dict:
    """
    Parse the TPCRA questionnaire Excel.
    Returns a dict with:
      - title, vendor, rep, email
      - sections: { letter: { name, questions: [{key, question, response, norm}] } }
      - all_items: flat list of all question rows
    """
    try:
        wb = load_workbook(file, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        st.error(f"Could not open file: {e}")
        return None

    result = {
        "title": "",
        "vendor": "",
        "rep": "",
        "email": "",
        "sections": {},
        "all_items": [],
    }

    current_section = None

    for row in rows:
        key, question, response = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)

        # Title row
        if isinstance(key, str) and "TPCRA" in key and result["title"] == "":
            result["title"] = key.strip()
            continue

        # Contact info rows
        if isinstance(question, str):
            ql = question.lower()
            if "company name" in ql and response:
                result["vendor"] = str(response).strip()
            elif "authorized representative" in ql and "email" not in ql and response:
                result["rep"] = str(response).strip()
            elif "email" in ql and response:
                result["email"] = str(response).strip()

        # Section header row (key is a single letter like 'A', 'B', ...)
        if isinstance(key, str) and len(key.strip()) == 1 and key.strip().isalpha():
            letter = key.strip().upper()
            name = SECTION_MAP.get(letter, question or letter)
            current_section = letter
            if letter not in result["sections"]:
                result["sections"][letter] = {"name": name, "questions": []}
            continue

        # Sub-section title rows (no response, question looks like a heading in CAPS)
        if key is not None and question is not None and response is None:
            continue

        # Question row
        if question and str(question).strip():
            sec = get_section(key) or current_section or "?"
            if sec and sec not in result["sections"]:
                result["sections"][sec] = {
                    "name": SECTION_MAP.get(sec, sec),
                    "questions": []
                }

            norm = normalize_response(response)
            item = {
                "key":      str(key).strip() if key else "",
                "section":  sec,
                "section_name": SECTION_MAP.get(sec, sec) if sec else "Other",
                "question": str(question).strip(),
                "response": str(response).strip() if response else "",
                "norm":     norm,
            }

            if sec and sec in result["sections"]:
                result["sections"][sec]["questions"].append(item)
            result["all_items"].append(item)

    return result


# ── Compliance helpers ────────────────────────────────────────────────────────
def compliance_counts(items: list[dict]) -> dict:
    c = {"Yes": 0, "No": 0, "Partial": 0, "N/A": 0, "Other": 0}
    for it in items:
        n = it["norm"]
        if n in c:
            c[n] += 1
        elif n not in ("—",):
            c["Other"] += 1
    return c


def compliance_score(counts: dict) -> int:
    """Weighted score: Yes=100, Partial=50, No/Other=0, N/A excluded."""
    scored = {k: v for k, v in counts.items() if k not in ("N/A", "Other", "—")}
    total = sum(scored.values())
    if total == 0:
        return 0
    earned = scored.get("Yes", 0) * 100 + scored.get("Partial", 0) * 50
    return round(earned / total)


def badge_html(norm: str) -> str:
    css = BADGE_CSS.get(norm, BADGE_CSS["—"])
    return f'<span class="response-badge" style="{css}">{norm}</span>'


# ── Sample download ───────────────────────────────────────────────────────────
def make_sample_excel() -> bytes:
    rows = [
        ("TPCRA Questionnaire - Part 2", None, None),
        (None, None, None),
        (None, "Question", "Response"),
        (1, "CONTACT INFORMATION", None),
        (1.1, "Company Name", "Acme Corp"),
        (1.2, "Name of Authorized Representative / Position", "Jane Doe"),
        (1.3, "Email Address of Authorized Representative", "jane@acme.com"),
        ("A", "ORGANIZATIONAL MANAGEMENT", None),
        ("A.1", "IT Security policies and procedures are established & documented.", "Yes"),
        ("A.2", "IT Security policies and procedures are reviewed at least annually.", "Yes"),
        ("B", "HUMAN RESOURCE MANAGEMENT", None),
        ("B.1", "Describe security awareness program/trainings conducted for employees.", "Annual training provided to all staff."),
        ("B.2", "IT security awareness training is provided to all employees.", "Yes"),
        ("C", "INFRASTRUCTURE SECURITY", None),
        ("C.1", "Patch management procedure is established.", "Yes"),
        ("C.2", "Anti-malware solution is deployed.", "Partial"),
        ("D", "DATA PROTECTION", None),
        ("D.1", "Encryption is used for data in transit.", "Yes"),
        ("D.2", "Encryption is used for data at rest.", "Yes"),
        ("E", "ACCESS MANAGEMENT", None),
        ("E.1", "Access to critical systems is granted on a need basis.", "Yes"),
        ("E.2", "Multi-factor authentication is enforced.", "NO"),
        ("F", "APPLICATION SECURITY", None),
        ("F.1", "Secure coding techniques are used.", "Yes"),
        ("F.2", "Penetration testing is conducted periodically.", "Yes"),
        ("G", "SYSTEM SECURITY", None),
        ("G.1", "Systems are hardened to established baseline standards.", "Yes"),
        ("H", "EMAIL SECURITY", None),
        ("H.1", "Email gateway scans for malware and phishing.", "Yes"),
        ("I", "MOBILE DEVICES", None),
        ("I.1", "Mobile devices are equipped with security software.", "Yes"),
        ("J", "INCIDENT RESPONSE", None),
        ("J.1", "Incident Response Plan is established.", "Yes"),
        ("J.2", "IRP is reviewed and tested annually.", "Yes"),
        ("K", "CLOUD SERVICES", None),
        ("K.1", "Cloud services are in scope.", "Yes"),
        ("K.2", "Cloud SP holds relevant security certifications.", "Yes"),
        ("L", "BUSINESS CONTINUITY", None),
        ("L.1", "A business continuity plan is established.", "Yes"),
        ("L.2", "BCP is updated at least annually.", "N/A"),
        ("L.3", "BCP is tested at least annually.", "Yes"),
    ]
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=False, sheet_name="Questionnaire")
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Upload questionnaire")
    uploaded = st.file_uploader(
        "TPCRA Excel file (.xlsx)",
        type=["xlsx", "xls"],
        help="Upload a completed TPCRA questionnaire in the standard format.",
    )
    st.divider()
    st.markdown("**Expected format**")
    st.markdown("""
- Single sheet named `Questionnaire`  
- Row 1: Title (`TPCRA Questionnaire - Part 2`)  
- Section headers: single letter keys (A, B, C …)  
- Question rows: keys like `A.1`, `B.2`, `C.1.1`  
- Response column: `Yes` / `No` / `Partial` / `N/A`
    """)
    st.divider()
    st.download_button(
        "⬇ Download sample Excel",
        data=make_sample_excel(),
        file_name="sample_TPCRA_questionnaire.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Empty state ───────────────────────────────────────────────────────────────
if not uploaded:
    st.title("🔐 TPCRA Risk Assessment Dashboard")
    st.markdown("Upload a completed questionnaire from the sidebar to generate the dashboard.")
    st.divider()
    c1, c2, c3, c4 = st.columns(4)
    c1.info("**Overview**\nCompliance score and response distribution across all sections.")
    c2.info("**By section**\nDrill into each domain — Org Mgmt, Data Protection, IAM, etc.")
    c3.info("**Gap analysis**\nAll `No` and `Partial` responses flagged for follow-up.")
    c4.info("**Full responses**\nBrowse every question and answer with filters.")
    st.stop()

# ── Parse ─────────────────────────────────────────────────────────────────────
data = parse_tpcra(uploaded)
if not data or not data["all_items"]:
    st.error("No questions found. Please check your file matches the TPCRA format.")
    st.stop()

all_items = data["all_items"]
sections  = data["sections"]

# ── Header ────────────────────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title(f"🔐 {data['vendor'] or 'Vendor'} — TPCRA Dashboard")
    if data["rep"] or data["email"]:
        st.caption(f"{data['rep']}  ·  {data['email']}")
with col_h2:
    st.caption(data["title"])

st.divider()

# ── Summary metrics ───────────────────────────────────────────────────────────
counts = compliance_counts(all_items)
score  = compliance_score(counts)
total  = sum(counts.values())

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Total questions", total)
m2.metric("✅ Yes",     counts["Yes"])
m3.metric("❌ No",      counts["No"])
m4.metric("⚠️ Partial", counts["Partial"])
m5.metric("➖ N/A",     counts["N/A"])
m6.metric("Compliance score", f"{score}%",
    delta="Strong" if score >= 80 else ("Moderate" if score >= 60 else "Needs attention"),
    delta_color="normal" if score >= 80 else ("off" if score >= 60 else "inverse"),
)
st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_overview, tab_sections, tab_gaps, tab_all = st.tabs([
    "Overview", "By section", "Gap analysis", "All responses"
])

# ════════════════════════════════
# TAB 1 — OVERVIEW
# ════════════════════════════════
with tab_overview:
    ch1, ch2 = st.columns([2, 1])

    with ch1:
        st.subheader("Compliance by section")
        sec_rows = []
        for letter, sec in sections.items():
            qs = sec["questions"]
            if not qs:
                continue
            sc = compliance_counts(qs)
            sec_rows.append({
                "Section": sec["name"],
                "Yes":     sc["Yes"],
                "No":      sc["No"],
                "Partial": sc["Partial"],
                "N/A":     sc["N/A"],
                "Score":   compliance_score(sc),
            })
        sec_df = pd.DataFrame(sec_rows)

        if not sec_df.empty:
            melt = sec_df.melt(
                id_vars="Section",
                value_vars=["Yes", "No", "Partial", "N/A"],
                var_name="Response", value_name="Count"
            )
            melt = melt[melt["Count"] > 0]
            fig = px.bar(
                melt, x="Count", y="Section", color="Response",
                orientation="h",
                color_discrete_map=COLORS,
                category_orders={"Response": ["Yes", "Partial", "No", "N/A"]},
                labels={"Count": "Questions", "Section": ""},
                height=max(300, len(sec_df) * 46),
            )
            fig.update_layout(
                margin=dict(l=0, r=0, t=10, b=0),
                legend_title_text="Response",
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                yaxis={"categoryorder": "total ascending"},
                font=dict(size=12),
            )
            fig.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.07)")
            st.plotly_chart(fig, use_container_width=True)

    with ch2:
        st.subheader("Overall distribution")
        labels = [k for k, v in counts.items() if v > 0 and k != "Other"]
        values = [counts[k] for k in labels]
        colors = [COLORS.get(k, "#888") for k in labels]
        fig2 = go.Figure(go.Pie(
            labels=labels, values=values, hole=0.62,
            marker_colors=colors, textinfo="label+percent",
            hovertemplate="%{label}: %{value}<extra></extra>",
        ))
        fig2.update_layout(
            margin=dict(l=0, r=0, t=10, b=0),
            showlegend=False, height=300,
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(size=12),
        )
        st.plotly_chart(fig2, use_container_width=True)

        st.subheader("Section scores")
        if not sec_df.empty:
            for _, row in sec_df.sort_values("Score").iterrows():
                color = "#639922" if row["Score"] >= 80 else ("#EF9F27" if row["Score"] >= 50 else "#E24B4A")
                st.markdown(
                    f'<div style="margin-bottom:8px">'
                    f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px">'
                    f'<span style="color:#555">{row["Section"]}</span>'
                    f'<span style="font-weight:500;color:{color}">{row["Score"]}%</span></div>'
                    f'<div style="background:#f0f0f0;border-radius:4px;height:6px;overflow:hidden">'
                    f'<div style="width:{row["Score"]}%;height:100%;background:{color};border-radius:4px"></div>'
                    f'</div></div>',
                    unsafe_allow_html=True,
                )

# ════════════════════════════════
# TAB 2 — BY SECTION
# ════════════════════════════════
with tab_sections:
    sec_letters = list(sections.keys())
    sec_names   = [f"{l} — {sections[l]['name']}" for l in sec_letters]
    choice      = st.selectbox("Select section", sec_names)
    chosen_letter = choice.split(" — ")[0]
    sec_data    = sections.get(chosen_letter, {})
    qs          = sec_data.get("questions", [])

    if not qs:
        st.info("No questions in this section.")
    else:
        sc = compliance_counts(qs)
        s_score = compliance_score(sc)

        sm1, sm2, sm3, sm4, sm5 = st.columns(5)
        sm1.metric("Questions", len(qs))
        sm2.metric("Yes",     sc["Yes"])
        sm3.metric("No",      sc["No"])
        sm4.metric("Partial", sc["Partial"])
        sm5.metric("Score",   f"{s_score}%")

        st.markdown("---")

        for item in qs:
            norm = item["norm"]
            badge = badge_html(norm)
            is_finding = norm in ("No", "Partial")
            bg = "#fffaf5" if norm == "Partial" else ("#fff5f5" if norm == "No" else "transparent")
            with st.container():
                st.markdown(
                    f'<div style="padding:10px 14px;border-radius:8px;margin-bottom:8px;'
                    f'background:{bg};border:0.5px solid #e9ecef">'
                    f'<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px">'
                    f'<div style="flex:1">'
                    f'<div style="font-size:11px;color:#999;margin-bottom:3px">{item["key"]}</div>'
                    f'<div style="font-size:13px;color:#333;line-height:1.5">{item["question"]}</div>'
                    f'</div>'
                    f'<div style="flex-shrink:0">{badge}</div>'
                    f'</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
                # Show long text responses
                if item["response"] and len(item["response"]) > 30 and norm not in ("Yes","No","Partial","N/A","—"):
                    with st.expander("View full response"):
                        st.markdown(f'<div style="font-size:13px;color:#555;line-height:1.6">{item["response"]}</div>', unsafe_allow_html=True)

# ════════════════════════════════
# TAB 3 — GAP ANALYSIS
# ════════════════════════════════
with tab_gaps:
    gaps = [it for it in all_items if it["norm"] in ("No", "Partial")]

    if not gaps:
        st.success("No gaps found — all answered questions are compliant.")
    else:
        g_no      = [g for g in gaps if g["norm"] == "No"]
        g_partial = [g for g in gaps if g["norm"] == "Partial"]

        st.markdown(f"**{len(gaps)} gaps identified** — {len(g_no)} non-compliant, {len(g_partial)} partial")
        st.divider()

        # Summary by section
        gap_sec = {}
        for g in gaps:
            s = g["section_name"]
            gap_sec.setdefault(s, {"No": 0, "Partial": 0})
            gap_sec[s][g["norm"]] += 1

        gap_sec_df = pd.DataFrame([
            {"Section": k, "No": v["No"], "Partial": v["Partial"]}
            for k, v in gap_sec.items()
        ]).sort_values("No", ascending=False)

        if not gap_sec_df.empty:
            fig_gap = px.bar(
                gap_sec_df.melt(id_vars="Section", var_name="Type", value_name="Count"),
                x="Section", y="Count", color="Type",
                color_discrete_map={"No": "#E24B4A", "Partial": "#EF9F27"},
                labels={"Count": "Gaps", "Section": ""},
                height=260,
                barmode="stack",
            )
            fig_gap.update_layout(
                margin=dict(l=0, r=0, t=10, b=0),
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                font=dict(size=12), xaxis_tickangle=-30,
                legend_title_text="",
            )
            fig_gap.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.07)")
            st.plotly_chart(fig_gap, use_container_width=True)

        st.divider()

        # Detailed gap list
        gap_filter = st.radio("Show", ["All gaps", "No only", "Partial only"], horizontal=True)
        shown = gaps if gap_filter == "All gaps" else [g for g in gaps if g["norm"] == gap_filter.split()[0]]

        for g in shown:
            norm = g["norm"]
            badge = badge_html(norm)
            bg = "#fff5f5" if norm == "No" else "#fffaf5"
            st.markdown(
                f'<div style="padding:10px 14px;border-radius:8px;margin-bottom:8px;'
                f'background:{bg};border:0.5px solid {"#f7c1c1" if norm=="No" else "#FAC775"}">'
                f'<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px">'
                f'<div style="flex:1">'
                f'<div style="font-size:11px;color:#999;margin-bottom:3px">{g["key"]} &nbsp;·&nbsp; {g["section_name"]}</div>'
                f'<div style="font-size:13px;color:#333;line-height:1.5">{g["question"]}</div>'
                f'</div>'
                f'<div style="flex-shrink:0">{badge}</div>'
                f'</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.divider()
        gap_df = pd.DataFrame([{
            "Key": g["key"], "Section": g["section_name"],
            "Question": g["question"], "Response": g["norm"],
        } for g in gaps])
        csv = gap_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇ Export gaps as CSV", data=csv,
                           file_name="tpcra_gaps.csv", mime="text/csv")

# ════════════════════════════════
# TAB 4 — ALL RESPONSES
# ════════════════════════════════
with tab_all:
    f1, f2 = st.columns([2, 3])
    with f1:
        resp_filter = st.multiselect(
            "Filter by response",
            options=["Yes", "No", "Partial", "N/A"],
            default=["Yes", "No", "Partial", "N/A"],
        )
    with f2:
        sec_filter = st.multiselect(
            "Filter by section",
            options=[f"{l} — {sections[l]['name']}" for l in sections],
            default=[f"{l} — {sections[l]['name']}" for l in sections],
        )

    sel_letters = {s.split(" — ")[0] for s in sec_filter}
    filtered = [
        it for it in all_items
        if it["norm"] in resp_filter and it["section"] in sel_letters
    ]

    st.caption(f"Showing {len(filtered)} of {len(all_items)} responses")

    rows_html = ""
    for it in filtered:
        badge = badge_html(it["norm"])
        resp_preview = it["response"] if len(it["response"]) <= 80 else it["response"][:77] + "…"
        rows_html += f"""<tr>
          <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#999;width:7%;vertical-align:top">{it['key']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#555;width:18%;vertical-align:top">{it['section_name']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;color:#333;width:55%;vertical-align:top;line-height:1.5">{it['question']}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:12px;color:#555;width:20%;vertical-align:top">{badge if it['norm'] in ('Yes','No','Partial','N/A') else resp_preview}</td>
        </tr>"""

    st.markdown(f"""
    <div style="border:1px solid #e9ecef;border-radius:10px;overflow:hidden;margin-top:8px">
      <table style="width:100%;border-collapse:collapse;table-layout:fixed">
        <thead>
          <tr style="background:#f8f9fa">
            <th style="padding:9px 10px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef;width:7%">Key</th>
            <th style="padding:9px 10px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef;width:18%">Section</th>
            <th style="padding:9px 10px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef;width:55%">Question</th>
            <th style="padding:9px 10px;text-align:left;font-size:11px;font-weight:600;color:#6c757d;text-transform:uppercase;letter-spacing:0.05em;border-bottom:1px solid #e9ecef;width:20%">Response</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>""", unsafe_allow_html=True)

    st.divider()
    exp1, exp2, _ = st.columns([1, 1, 3])
    all_df = pd.DataFrame([{
        "Key": it["key"], "Section": it["section_name"],
        "Question": it["question"], "Response": it["norm"],
    } for it in filtered])
    with exp1:
        st.download_button("⬇ Export CSV", data=all_df.to_csv(index=False).encode(),
                           file_name="tpcra_responses.csv", mime="text/csv",
                           use_container_width=True)
    with exp2:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            all_df.to_excel(writer, index=False, sheet_name="TPCRA Responses")
        st.download_button("⬇ Export Excel", data=buf.getvalue(),
                           file_name="tpcra_responses.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

st.divider()
st.caption("TPCRA — Third-Party Cyber Risk Assessment Dashboard")
