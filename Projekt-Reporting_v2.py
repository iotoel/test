import streamlit as st
import pandas as pd
from datetime import datetime
import re
import html as html_lib
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from pathlib import Path
import os

st.set_page_config(page_title="ICT-Projektübersicht", layout="wide", page_icon="🏢")

# Pfad zur Exceldatei
SCRIPT_DIR = Path(__file__).parent

#LOCAL_FILE = "https://github.com/iotoel/projektuebersicht/raw/refs/heads/main/Monatliches%20Projekt-Reporting.xlsx"
LOCAL_FILE = "https://raw.githubusercontent.com/iotoel/projektuebersicht/main/Monatliches%20Projekt-Reporting.xlsx"
FALLBACK_FILE = SCRIPT_DIR / "Monatliches Projekt-Reporting.xlsx" #Path(r"D:\Desktop\Monatliches Projekt-Reporting.xlsx")

try:
    pfad = LOCAL_FILE
except:
    pfad = FALLBACK_FILE

CSS = """
<style>
  section[data-testid="stSidebar"] { min-width: 220px !important; max-width: 220px !important; }
  .main { background-color: #f0f2f5; }
  .block-container { padding: 1rem 1.5rem; }
  .top-bar {
    display: flex; align-items: center; justify-content: space-between;
    background: white; border-radius: 10px; padding: 12px 20px;
    margin-bottom: 16px; box-shadow: 0 1px 4px rgba(0,0,0,0.1);
  }
  .logo-box { background: #cc0000; color: white; font-weight: bold; font-size: 18px;
    padding: 6px 12px; border-radius: 4px; display:inline-block; }
  .logo-sub { color: #666; font-size: 11px; margin-top: 2px; }
  .title-main { font-size: 22px; font-weight: 700; color: #1a1a2e; }
  .update-info { font-size: 11px; color: #888; text-align: right; }
  .kanban-col { background: #e8eaf0; border-radius: 10px; padding: 10px; min-height: 200px; }
  .col-header {
    font-size: 11px; font-weight: 700; color: #555;
    text-transform: uppercase; letter-spacing: 0.5px;
    margin-bottom: 8px; padding: 4px 6px; border-left: 3px solid #3a86ff;
  }
  .project-card {
    background: white; border-radius: 8px; padding: 12px;
    margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    border-top: 3px solid #3a86ff; cursor: pointer;
    transition: box-shadow 0.15s, transform 0.15s;
  }
  .project-card:hover { box-shadow: 0 4px 12px rgba(58,134,255,0.25); transform: translateY(-2px); }
  .card-id { font-size: 10px; color: #888; font-weight: 600; }
  .card-project { font-size: 12px; font-weight: 700; color: #1a1a2e; margin: 2px 0; }
  .card-pm { font-size: 11px; color: #555; }
  .card-phase-end { font-size: 10px; font-weight: 600; padding: 2px 6px; border-radius: 3px; display: inline-block; margin: 3px 0; }
  .phase-end-ok  { background: #d4edda; color: #155724; }
  .phase-end-warn{ background: #fff3cd; color: #856404; }
  .phase-end-late{ background: #f8d7da; color: #721c24; }
  .status-badge { font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 10px; display: inline-block; margin: 3px 0; }
  .status-green  { background: #d4edda; color: #155724; }
  .status-yellow { background: #fff3cd; color: #856404; }
  .status-red    { background: #f8d7da; color: #721c24; }
  .budget-row    { font-size: 10px; color: #666; margin: 2px 0; }
  .section-label { font-size: 10px; font-weight: 700; color: #888; margin-top: 6px; }
  .section-text  { font-size: 10px; color: #333; margin: 2px 0 4px 0; line-height: 1.4; white-space: pre-wrap; word-break: break-word; }
  .escalation-badge { font-size: 10px; background: #f8d7da; color: #721c24; padding: 2px 8px; border-radius: 3px; display: inline-block; margin: 3px 0; }
  /* Detail view */
  .detail-header { background: white; border-radius: 12px; padding: 20px 24px; margin-bottom: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); border-left: 5px solid #3a86ff; }
  .detail-title { font-size: 22px; font-weight: 800; color: #1a1a2e; margin: 4px 0; }
  .detail-meta { font-size: 13px; color: #555; margin: 3px 0; }
  .detail-section { background: white; border-radius: 10px; padding: 16px 20px; margin-bottom: 12px; box-shadow: 0 1px 4px rgba(0,0,0,0.07); }
  .detail-section h4 { font-size: 13px; font-weight: 700; color: #3a86ff; margin: 0 0 8px 0; text-transform: uppercase; letter-spacing: 0.5px; }
  .detail-section p { font-size: 13px; color: #333; line-height: 1.6; margin: 0; white-space: pre-wrap; }
  .detail-badge-lg { font-size: 13px; font-weight: 700; padding: 5px 14px; border-radius: 20px; display: inline-block; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

PHASE_ORDER = [
    "Entry (Erstellung des Projektantrages)",
    "Review Projektantrag (Review durch Transformation Agent)",
    "Projekt-Start (Initialisierung)",
    "Fertigstellung Konzept (Spezifikation)",
    "Fertigstellung Realisierung (Realisierung)",
    "Projekt Abschluss (Transition)",
]
PHASE_SHORT = {
    "Entry (Erstellung des Projektantrages)": "Entry",
    "Review Projektantrag (Review durch Transformation Agent)": "Review PA",
    "Projekt-Start (Initialisierung)": "Initialisierung",
    "Fertigstellung Konzept (Spezifikation)": "Spezifikation",
    "Fertigstellung Realisierung (Realisierung)": "Realisierung",
    "Projekt Abschluss (Transition)": "Transition",
}

def e(text):
    if text is None: return ""
    return html_lib.escape(str(text))

def classify_status(val):
    """Return 'gut', 'risiko', 'kritisch', or 'unbekannt'"""
    s = str(val).lower()
    if "gut" in s or "grün" in s or "green" in s: return "gut"
    if "risiko" in s or "gelb" in s or "yellow" in s: return "risiko"
    if "kritisch" in s or "rot" in s or "red" in s: return "kritisch"
    return "unbekannt"

def load_data(path_or_file):
    df = pd.read_excel(path_or_file, header=0)
    df.columns = [str(c) for c in df.columns]
    cols = list(df.columns)
    rename = {
        cols[0]: "id", cols[1]: "start", cols[2]: "end",
        cols[3]: "email", cols[4]: "name",
        cols[5]: "project", cols[6]: "pm",
        cols[7]: "phase", cols[8]: "phase_end_date",
        cols[9]: "date_changed_reason", cols[10]: "status",
        cols[11]: "budget_ext", cols[12]: "budget_int",
        cols[13]: "successes", cols[14]: "challenges",
        cols[15]: "escalation", cols[16]: "outlook", cols[17]: "misc",
    }
    df = df.rename(columns=rename)
    df["phase_end_date"] = pd.to_datetime(df["phase_end_date"], errors="coerce")
    df["start"] = pd.to_datetime(df["start"], errors="coerce")
    df["status_class"] = df["status"].apply(classify_status)
    df = df.sort_values("start").groupby("project").last().reset_index()
    return df

def is_empty(val):
    if val is None: return True
    s = str(val).strip()
    return s.lower() in {"", "nan", "-", "n/a", "none", "nein", "no", ".", "unknown", "keine"}

def val_or_none(val, n=None):
    if is_empty(val): return None
    t = str(val).strip()
    if n: return t[:n] + ("…" if len(t) > n else "")
    return t

def status_html(status_val, large=False):
    cls_map = {"gut": "status-green", "risiko": "status-yellow", "kritisch": "status-red"}
    label_map = {"gut": "🟢 Gut", "risiko": "🟡 Risiko", "kritisch": "🔴 Kritisch", "unbekannt": "❓ Unbekannt"}
    sc = classify_status(status_val)
    css_class = cls_map.get(sc, "status-yellow")
    label = label_map.get(sc, "❓")
    badge_class = "detail-badge-lg" if large else "status-badge"
    return f'<span class="{badge_class} {css_class}">{label}</span>'

def phase_end_html(dt, large=False):
    if pd.isna(dt): return ""
    today = datetime.now()
    diff = (dt - today).days
    label = dt.strftime("%d.%m.%Y")
    cls = "phase-end-ok" if diff >= 14 else ("phase-end-warn" if diff >= 0 else "phase-end-late")
    size = "font-size:13px; padding:4px 10px;" if large else ""
    return f'<span class="card-phase-end {cls}" style="{size}">📅 {e(label)}</span>'

def render_card_html(row, idx):
    proj_id = str(row.get("project", ""))
    pm = val_or_none(row.get("pm")) or ""
    pa_match = re.search(r"(PA-\d+)", proj_id)
    pa_id = pa_match.group(1) if pa_match else ""
    proj_name = proj_id.replace(pa_id, "").strip().strip("()").strip()
    budget_ext = val_or_none(row.get("budget_ext"))
    budget_int = val_or_none(row.get("budget_int"))
    escalation = val_or_none(row.get("escalation"), 100)
    successes = val_or_none(row.get("successes"), 180)
    challenges = val_or_none(row.get("challenges"), 180)
    outlook = val_or_none(row.get("outlook"), 150)

    parts = [f'<div class="project-card" onclick="window.location.href=\'?card={idx}\'" style="cursor:pointer">',
        f'  <div class="card-id">{e(pa_id)}</div>',
        f'  <div class="card-project">{e(proj_name or proj_id)}</div>']
    if pm: parts.append(f'  <div class="card-pm">👤 {e(pm)}</div>')
    parts.append(phase_end_html(row.get("phase_end_date")))
    parts.append(status_html(row.get("status", "")))
    if budget_ext: parts.append(f'  <div class="budget-row">💰 Ext: {e(budget_ext)}</div>')
    if budget_int: parts.append(f'  <div class="budget-row">🏠 Int: {e(budget_int)}</div>')
    if escalation: parts.append(f'  <div class="escalation-badge">🚨 {e(escalation)}</div>')
    if successes: parts.append(f'  <div class="section-label">🏆 Erfolge</div><div class="section-text">{e(successes)}</div>')
    if challenges: parts.append(f'  <div class="section-label">⚠️ Herausforderungen</div><div class="section-text">{e(challenges)}</div>')
    if outlook: parts.append(f'  <div class="section-label">📅 Ausblick</div><div class="section-text">{e(outlook)}</div>')
    parts.append('</div>')
    return "\n".join(parts)

def detail_section(icon, title, content):
    if not content: return ""
    return f"""
<div class="detail-section">
  <h4>{icon} {e(title)}</h4>
  <p>{e(content)}</p>
</div>"""

def render_detail_view(row, df_index):
    proj_id = str(row.get("project", ""))
    pa_match = re.search(r"(PA-\d+)", proj_id)
    pa_id = pa_match.group(1) if pa_match else ""
    proj_name = proj_id.replace(pa_id, "").strip().strip("()").strip() or proj_id
    pm = val_or_none(row.get("pm")) or "–"
    phase = val_or_none(row.get("phase")) or "–"
    start = row.get("start")
    start_str = start.strftime("%d.%m.%Y") if pd.notna(start) else "–"

    st.markdown(f"""
    <div class="detail-header">
      <div style="font-size:12px;color:#888;font-weight:600;">{e(pa_id)}</div>
      <div class="detail-title">{e(proj_name)}</div>
      <div class="detail-meta">👤 {e(pm)} &nbsp;|&nbsp; 📊 {e(PHASE_SHORT.get(phase, phase))} &nbsp;|&nbsp; 🗓 Start: {e(start_str)}</div>
      <div style="margin-top:8px">
        {phase_end_html(row.get("phase_end_date"), large=True)}
        &nbsp;{status_html(row.get("status",""), large=True)}
      </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(detail_section("💰", "Budget extern", val_or_none(row.get("budget_ext"))), unsafe_allow_html=True)
        st.markdown(detail_section("🏠", "Budget intern", val_or_none(row.get("budget_int"))), unsafe_allow_html=True)
        st.markdown(detail_section("🏆", "Erfolge", val_or_none(row.get("successes"))), unsafe_allow_html=True)
        st.markdown(detail_section("📅", "Ausblick", val_or_none(row.get("outlook"))), unsafe_allow_html=True)
    with col2:
        st.markdown(detail_section("⏳", "Terminänderung Begründung", val_or_none(row.get("date_changed_reason"))), unsafe_allow_html=True)
        st.markdown(detail_section("⚠️", "Herausforderungen", val_or_none(row.get("challenges"))), unsafe_allow_html=True)
        st.markdown(detail_section("🚨", "Eskalation", val_or_none(row.get("escalation"))), unsafe_allow_html=True)
        st.markdown(detail_section("📝", "Sonstiges", val_or_none(row.get("misc"))), unsafe_allow_html=True)

    # PDF export for single project
    st.markdown("---")
    if st.button("📄 Als PDF exportieren", key="pdf_single"):
        pdf_bytes = export_single_pdf(row, pa_id, proj_name, pm, phase, start_str)
        st.download_button("⬇️ PDF herunterladen", data=pdf_bytes,
            file_name=f"{pa_id}_{proj_name[:30]}.pdf", mime="application/pdf")

def export_single_pdf(row, pa_id, proj_name, pm, phase, start_str):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("title", fontSize=18, fontName="Helvetica-Bold", textColor=colors.HexColor("#1a1a2e"), spaceAfter=4)
    meta_style = ParagraphStyle("meta", fontSize=10, textColor=colors.HexColor("#555555"), spaceAfter=2)
    heading_style = ParagraphStyle("heading", fontSize=11, fontName="Helvetica-Bold", textColor=colors.HexColor("#3a86ff"), spaceBefore=10, spaceAfter=4)
    body_style = ParagraphStyle("body", fontSize=10, leading=14, textColor=colors.HexColor("#333333"), spaceAfter=6)

    story = []
    story.append(Paragraph(f"{pa_id} – {proj_name}", title_style))
    story.append(Paragraph(f"Projektleiter: {pm}  |  Phase: {PHASE_SHORT.get(phase, phase)}  |  Start: {start_str}", meta_style))
    sc = classify_status(row.get("status",""))
    status_colors = {"gut": "#155724", "risiko": "#856404", "kritisch": "#721c24"}
    story.append(Paragraph(f"<font color='{status_colors.get(sc,'#333')}'><b>Status: {row.get('status','')}</b></font>", meta_style))
    pe = row.get("phase_end_date")
    if pd.notna(pe):
        story.append(Paragraph(f"Phase-Ende: {pe.strftime('%d.%m.%Y')}", meta_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#dddddd"), spaceAfter=10))

    fields = [
        ("Budget extern", "budget_ext"), ("Budget intern", "budget_int"),
        ("Terminänderung Begründung", "date_changed_reason"), ("Erfolge", "successes"),
        ("Herausforderungen", "challenges"), ("Eskalation", "escalation"),
        ("Ausblick", "outlook"), ("Sonstiges", "misc"),
    ]
    for label, key in fields:
        v = val_or_none(row.get(key))
        if v:
            story.append(Paragraph(label, heading_style))
            story.append(Paragraph(v.replace("\n", "<br/>"), body_style))

    story.append(Spacer(1, 20))
    story.append(Paragraph(f"Exportiert am {datetime.now().strftime('%d.%m.%Y %H:%M')}", ParagraphStyle("footer", fontSize=8, textColor=colors.grey)))
    doc.build(story)
    return buf.getvalue()

def export_overview_pdf(filtered_df):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    title_s = ParagraphStyle("t", fontSize=16, fontName="Helvetica-Bold", textColor=colors.HexColor("#1a1a2e"), spaceAfter=4)
    sub_s = ParagraphStyle("s", fontSize=9, textColor=colors.grey, spaceAfter=12)
    col_s = ParagraphStyle("c", fontSize=9, fontName="Helvetica-Bold", textColor=colors.HexColor("#3a86ff"), spaceBefore=10, spaceAfter=4)
    card_title_s = ParagraphStyle("ct", fontSize=9, fontName="Helvetica-Bold", textColor=colors.HexColor("#1a1a2e"))
    card_body_s = ParagraphStyle("cb", fontSize=8, leading=11, textColor=colors.HexColor("#444444"))

    story = [Paragraph("ICT-Projektübersicht", title_s),
             Paragraph(f"Empa – Exportiert am {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  {len(filtered_df)} Projekte", sub_s),
             HRFlowable(width="100%", thickness=1, color=colors.HexColor("#cccccc"), spaceAfter=10)]

    status_color_map = {"gut": colors.HexColor("#155724"), "risiko": colors.HexColor("#856404"), "kritisch": colors.HexColor("#721c24")}

    for phase in PHASE_ORDER:
        phase_rows = filtered_df[filtered_df["phase"].str.strip() == phase.strip()]
        if phase_rows.empty: continue
        short = PHASE_SHORT.get(phase, phase)
        story.append(Paragraph(f"{short.upper()}  ({len(phase_rows)})", col_s))

        table_data = [["Projekt", "Projektleiter", "Phase-Ende", "Status", "Erfolge / Herausforderungen"]]
        for _, row in phase_rows.iterrows():
            proj_id = str(row.get("project",""))
            pa_match = re.search(r"(PA-\d+)", proj_id)
            pa_id = pa_match.group(1) if pa_match else ""
            proj_name = proj_id.replace(pa_id,"").strip().strip("()").strip() or proj_id
            pm = val_or_none(row.get("pm")) or "–"
            pe = row.get("phase_end_date")
            pe_str = pe.strftime("%d.%m.%Y") if pd.notna(pe) else "–"
            sc = classify_status(row.get("status",""))
            status_label = {"gut":"🟢 Gut","risiko":"🟡 Risiko","kritisch":"🔴 Kritisch"}.get(sc,"❓")
            suc = (val_or_none(row.get("successes"),"") or "")[:120]
            chal = (val_or_none(row.get("challenges"),"") or "")[:120]
            detail = ""
            if suc: detail += f"✅ {suc}"
            if suc and chal: detail += "\n"
            if chal: detail += f"⚠ {chal}"
            table_data.append([
                Paragraph(f"<b>{pa_id}</b><br/>{proj_name}", card_title_s),
                Paragraph(pm, card_body_s),
                Paragraph(pe_str, card_body_s),
                Paragraph(status_label, card_body_s),
                Paragraph(detail.replace("\n","<br/>"), card_body_s),
            ])

        t = Table(table_data, colWidths=[3.5*cm, 3*cm, 2.2*cm, 2*cm, 7.3*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#3a86ff")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 8),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f8f9fa")]),
            ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#dddddd")),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))

    doc.build(story)
    return buf.getvalue()

# ──────────────────────────────────────────
# Load data
# ──────────────────────────────────────────
df = load_data(pfad)
#uploaded = st.sidebar.file_uploader("📂 Excel laden (.xlsx)", type=["xlsx"])
#if uploaded:
#    df = load_data(uploaded)
#else:
#    try:
#        df = load_data("/mnt/user-data/uploads/1780925656110_Monatliches_Projekt-Reporting.xlsx")
#    except Exception:
#        df = pd.DataFrame()

now_str = datetime.now().strftime("%d. %B %Y")
st.markdown(f"""
<div class="top-bar">
  <div><div class="logo-box">Empa</div><div class="logo-sub">Materials Science and Technology</div></div>
  <div class="title-main">ICT-Projektübersicht</div>
  <div class="update-info">Letzte Aktualisierung: {now_str}</div>
</div>""", unsafe_allow_html=True)

if df.empty:
    st.info("Bitte eine Excel-Datei hochladen.")
    st.stop()

# ──────────────────────────────────────────
# Sidebar filters
# ──────────────────────────────────────────
st.sidebar.markdown("### 🔍 Filter")
all_pms = sorted(df["pm"].dropna().unique().tolist())
selected_pms = st.sidebar.multiselect("Projektleiter", all_pms)
# Use status_class for reliable filtering
selected_statuses = st.sidebar.multiselect("Status", ["🟢 Gut", "🟡 Risiko", "🔴 Kritisch"])
search = st.sidebar.text_input("Projektname suchen", "")
st.sidebar.markdown("---")
st.sidebar.markdown("### 📤 Export")
export_pdf_btn = st.sidebar.button("📄 Übersicht als PDF")

filtered = df.copy()
if selected_pms:
    filtered = filtered[filtered["pm"].isin(selected_pms)]
if selected_statuses:
    allowed = set()
    for sel in selected_statuses:
        if "gut" in sel.lower(): allowed.add("gut")
        if "risiko" in sel.lower(): allowed.add("risiko")
        if "kritisch" in sel.lower(): allowed.add("kritisch")
    filtered = filtered[filtered["status_class"].isin(allowed)]
if search:
    filtered = filtered[filtered["project"].str.contains(search, case=False, na=False)]

if export_pdf_btn:
    pdf_bytes = export_overview_pdf(filtered)
    st.sidebar.download_button("⬇️ PDF herunterladen", data=pdf_bytes,
        file_name=f"ICT_Projektuebersicht_{datetime.now().strftime('%Y%m%d')}.pdf",
        mime="application/pdf")

# ──────────────────────────────────────────
# State: selected card
# ──────────────────────────────────────────
if "selected_card" not in st.session_state:
    st.session_state.selected_card = None

# ──────────────────────────────────────────
# Stats
# ──────────────────────────────────────────
total  = len(filtered)
green  = (filtered["status_class"] == "gut").sum()
yellow = (filtered["status_class"] == "risiko").sum()
red    = (filtered["status_class"] == "kritisch").sum()

c1,c2,c3,c4 = st.columns(4)
c1.metric("📋 Projekte gesamt", total)
c2.metric("🟢 Gut", green)
c3.metric("🟡 Risiko", yellow)
c4.metric("🔴 Kritisch", red)

st.markdown("---")

# ──────────────────────────────────────────
# Detail view
# ──────────────────────────────────────────
if st.session_state.selected_card is not None:
    idx = st.session_state.selected_card
    if idx in df.index:
        row = df.loc[idx]
        col_back, col_title = st.columns([1, 8])
        with col_back:
            if st.button("← Zurück"):
                st.session_state.selected_card = None
                st.rerun()
        with col_title:
            proj_id = str(row.get("project",""))
            pa_match = re.search(r"(PA-\d+)", proj_id)
            pa_id = pa_match.group(1) if pa_match else ""
            proj_name = proj_id.replace(pa_id,"").strip().strip("()").strip() or proj_id
            st.markdown(f"### {pa_id} – {proj_name}")
        render_detail_view(row, idx)
    else:
        st.session_state.selected_card = None
        st.rerun()
else:
    # ──────────────────────────────────────
    # Kanban board with clickable buttons
    # ──────────────────────────────────────
    cols = st.columns(len(PHASE_ORDER))
    for i, phase in enumerate(PHASE_ORDER):
        rows_in_phase = filtered[filtered["phase"].str.strip() == phase.strip()]
        count = len(rows_in_phase)
        short = PHASE_SHORT.get(phase, phase)

        with cols[i]:
            st.markdown(f'<div class="kanban-col"><div class="col-header">{e(short)} ({count})</div>', unsafe_allow_html=True)

            for idx, row in rows_in_phase.iterrows():
                proj_id = str(row.get("project",""))
                pa_match = re.search(r"(PA-\d+)", proj_id)
                pa_id = pa_match.group(1) if pa_match else ""
                proj_name = proj_id.replace(pa_id,"").strip().strip("()").strip() or proj_id
                pm = val_or_none(row.get("pm")) or ""
                budget_ext = val_or_none(row.get("budget_ext"))
                budget_int = val_or_none(row.get("budget_int"))
                escalation = val_or_none(row.get("escalation"), 80)
                successes = val_or_none(row.get("successes"), 140)
                challenges = val_or_none(row.get("challenges"), 140)
                outlook = val_or_none(row.get("outlook"), 100)

                card_parts = [f'<div class="project-card">',
                    f'<div class="card-id">{e(pa_id)}</div>',
                    f'<div class="card-project">{e(proj_name)}</div>']
                if pm: card_parts.append(f'<div class="card-pm">👤 {e(pm)}</div>')
                card_parts.append(phase_end_html(row.get("phase_end_date")))
                card_parts.append(status_html(row.get("status","")))
                if budget_ext: card_parts.append(f'<div class="budget-row">💰 Ext: {e(budget_ext)}</div>')
                if budget_int: card_parts.append(f'<div class="budget-row">🏠 Int: {e(budget_int)}</div>')
                if escalation: card_parts.append(f'<div class="escalation-badge">🚨 {e(escalation)}</div>')
                if successes: card_parts.append(f'<div class="section-label">🏆 Erfolge</div><div class="section-text">{e(successes)}</div>')
                if challenges: card_parts.append(f'<div class="section-label">⚠️ Herausforderungen</div><div class="section-text">{e(challenges)}</div>')
                if outlook: card_parts.append(f'<div class="section-label">📅 Ausblick</div><div class="section-text">{e(outlook)}</div>')
                card_parts.append('</div>')
                st.markdown("\n".join(card_parts), unsafe_allow_html=True)

                # Invisible Streamlit button overlaid via custom CSS trick
                btn_key = f"card_{idx}"
                if st.button(f"🔍 Details", key=btn_key, help=f"{pa_id} – {proj_name} öffnen",
                             use_container_width=True):
                    st.session_state.selected_card = idx
                    st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)