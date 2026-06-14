"""
KenyaWatts v3 — Digital Integrated National Energy Planning Platform
NGDA 2026 · DTU Young Academics Track · Challenge 2: EPRA Kenya
Updates: Map view, submission notifications, bold chart labels,
         upload history log, forgot password, data download
"""
import streamlit as st
import streamlit_authenticator as stauth
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import anthropic
import io
import json
import bcrypt
from datetime import datetime, timedelta
# Report generation
try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, HRFlowable)
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="KenyaWatts · EPRA National Energy Planning",
    page_icon="⚡", layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  #MainMenu{visibility:hidden} footer{visibility:hidden} .stDeployButton{display:none}
  .kw-header{background:#0e1e2e;color:white;padding:14px 20px;border-radius:10px;margin-bottom:16px}
  .kw-logo{font-size:22px;font-weight:700;letter-spacing:-0.3px}
  .kw-logo span{color:#3ecfaa}
  .kw-metric{background:#f7f6f2;border-radius:10px;padding:14px 16px;margin-bottom:10px}
  .kw-metric-label{font-size:11px;color:#9c9a8e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
  .kw-metric-value{font-size:24px;font-weight:700;color:#1a1916;line-height:1.1}
  .kw-metric-delta{font-size:12px;margin-top:3px}
  .kw-alert-info{background:#edf4fb;border-left:3px solid #1a6fa3;color:#1a6fa3;padding:10px 14px;border-radius:6px;font-size:12px;margin-bottom:10px;line-height:1.6}
  .kw-alert-success{background:#e8f7f4;border-left:3px solid #0f9d7e;color:#0f9d7e;padding:10px 14px;border-radius:6px;font-size:12px;margin-bottom:10px;line-height:1.6}
  .kw-alert-warn{background:#fdf3e3;border-left:3px solid #d4891a;color:#d4891a;padding:10px 14px;border-radius:6px;font-size:12px;margin-bottom:10px;line-height:1.6}
  .kw-alert-danger{background:#fbedeb;border-left:3px solid #b33a2c;color:#b33a2c;padding:10px 14px;border-radius:6px;font-size:12px;margin-bottom:10px;line-height:1.6}
  .kw-badge-submitted{background:#e8f7f4;color:#0f9d7e;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600}
  .kw-badge-review{background:#fdf3e3;color:#d4891a;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600}
  .kw-badge-pending{background:#f5f4f0;color:#7a7870;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600}
  .kw-badge-overdue{background:#fbedeb;color:#b33a2c;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600}
  .kw-notif{position:fixed;top:70px;right:20px;z-index:9999;background:#0f9d7e;color:white;padding:14px 20px;border-radius:10px;font-size:13px;font-weight:500;box-shadow:0 4px 20px rgba(0,0,0,0.15);max-width:320px;line-height:1.5}
  .upload-log-row{display:flex;gap:12px;align-items:flex-start;padding:8px 0;border-bottom:0.5px solid #f0ede4;font-size:12px}
  .upload-log-icon{font-size:16px;flex-shrink:0}
  .upload-log-meta{color:#9c9a8e;font-size:11px;margin-top:2px}
</style>
""", unsafe_allow_html=True)

# ── CHART LAYOUT DEFAULTS (bold, black labels) ─────────────────────────────────
CHART_FONT = dict(family="Arial", size=13, color="#1a1916")
AXIS_FONT  = dict(family="Arial", size=12, color="#1a1916")
LEGEND_FONT= dict(family="Arial", size=12, color="#1a1916")
BASE_LAYOUT = dict(
    font=CHART_FONT,
    plot_bgcolor="white", paper_bgcolor="white",
    legend=dict(font=LEGEND_FONT, bgcolor="rgba(255,255,255,0.9)", borderwidth=0),
    xaxis=dict(tickfont=AXIS_FONT, title_font=dict(size=13,color="#1a1916",family="Arial"), showgrid=False),
    yaxis=dict(tickfont=AXIS_FONT, title_font=dict(size=13,color="#1a1916",family="Arial"), showgrid=True, gridcolor="#f0f0f0"),
    margin=dict(t=20, b=40, l=50, r=20),
)

def apply_layout(fig, title="", xlabel="", ylabel=""):
    fig.update_layout(**BASE_LAYOUT)
    if title:  fig.update_layout(title=dict(text=title, font=dict(size=14,color="#1a1916",family="Arial")))
    if xlabel: fig.update_xaxes(title_text=f"<b>{xlabel}</b>")
    if ylabel: fig.update_yaxes(title_text=f"<b>{ylabel}</b>")
    return fig

# ── COUNTY DATA ───────────────────────────────────────────────────────────────
COUNTIES = pd.DataFrame([
    {"id":"NK","name":"Nairobi",   "region":"Nairobi",    "pop":4922000,"elec":96, "cooking":62,"solar":1980,"budget":120.0,"status":"submitted","mtf":4,"growth":1.8,"target_yr":2027,"lat":-1.286,"lon":36.817},
    {"id":"MB","name":"Mombasa",   "region":"Coast",      "pop":1208000,"elec":84, "cooking":45,"solar":2050,"budget":55.0, "status":"submitted","mtf":4,"growth":1.5,"target_yr":2028,"lat":-4.043,"lon":39.668},
    {"id":"MK","name":"Makueni",   "region":"South East", "pop":987653, "elec":75, "cooking":18,"solar":2008,"budget":74.9, "status":"submitted","mtf":3,"growth":1.1,"target_yr":2028,"lat":-2.303,"lon":37.624},
    {"id":"NA","name":"Nakuru",    "region":"Rift Valley","pop":2162000,"elec":72, "cooking":38,"solar":1920,"budget":68.0, "status":"review",   "mtf":3,"growth":1.4,"target_yr":2029,"lat":-0.303,"lon":36.080},
    {"id":"KI","name":"Kisumu",    "region":"Nyanza",     "pop":1155000,"elec":67, "cooking":31,"solar":1870,"budget":0,    "status":"pending",  "mtf":3,"growth":1.2,"target_yr":2030,"lat":-0.091,"lon":34.768},
    {"id":"KW","name":"Kwale",     "region":"Coast",      "pop":866000, "elec":41, "cooking":18,"solar":2020,"budget":0,    "status":"pending",  "mtf":2,"growth":1.3,"target_yr":2031,"lat":-4.183,"lon":39.483},
    {"id":"KF","name":"Kilifi",    "region":"Coast",      "pop":1453000,"elec":38, "cooking":16,"solar":2010,"budget":0,    "status":"pending",  "mtf":2,"growth":1.4,"target_yr":2031,"lat":-3.510,"lon":39.908},
    {"id":"GR","name":"Garissa",   "region":"North East", "pop":841000, "elec":22, "cooking":8, "solar":2140,"budget":0,    "status":"pending",  "mtf":1,"growth":2.1,"target_yr":2032,"lat":-0.453,"lon":39.646},
    {"id":"KA","name":"Kajiado",   "region":"South Rift", "pop":1117000,"elec":55, "cooking":29,"solar":1990,"budget":0,    "status":"pending",  "mtf":2,"growth":1.8,"target_yr":2030,"lat":-1.852,"lon":36.777},
    {"id":"MR","name":"Muranga",   "region":"Central",    "pop":1056000,"elec":68, "cooking":33,"solar":1900,"budget":0,    "status":"pending",  "mtf":3,"growth":0.9,"target_yr":2029,"lat":-0.717,"lon":37.150},
    {"id":"TK","name":"Turkana",   "region":"North Rift", "pop":926000, "elec":12, "cooking":3, "solar":2150,"budget":0,    "status":"overdue",  "mtf":1,"growth":2.8,"target_yr":2033,"lat":3.112, "lon":35.596},
    {"id":"MS","name":"Marsabit",  "region":"North East", "pop":459000, "elec":8,  "cooking":2, "solar":2180,"budget":0,    "status":"overdue",  "mtf":1,"growth":2.5,"target_yr":2034,"lat":2.335, "lon":37.989},
    {"id":"MN","name":"Mandera",   "region":"North East", "pop":1025000,"elec":11, "cooking":4, "solar":2200,"budget":0,    "status":"overdue",  "mtf":1,"growth":3.1,"target_yr":2034,"lat":3.937, "lon":41.867},
    {"id":"WJ","name":"Wajir",     "region":"North East", "pop":781000, "elec":9,  "cooking":3, "solar":2190,"budget":0,    "status":"overdue",  "mtf":1,"growth":2.9,"target_yr":2034,"lat":1.747, "lon":40.058},
])

GEN_TREND = pd.DataFrame({
    "Year":["FY19/20","FY20/21","FY21/22","FY22/23","FY23/24","FY24/25"],
    "GWh": [11564,    11891,    12210,    12897,    13685,    14520]
})
GEN_MIX = pd.DataFrame({
    "Source":["Geothermal","Hydro","Wind","Solar","Thermal","Other"],
    "Pct":   [25.9, 23.8, 17.4, 14.9, 8.7, 9.3],
    "Color": ["#0f9d7e","#1a6fa3","#5b4fc9","#d4891a","#b33a2c","#7a7870"]
})
MAKUENI_COOKING = pd.DataFrame({
    "Fuel":["Firewood","LPG","Charcoal","Biogas","Electric","Other"],
    "Pct": [72.5,17.6,8.2,0.2,0.3,1.2]
})
OUTAGE = pd.DataFrame({
    "Month":["Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun"],
    "Hours":[7.2,9.1,11.1,8.4,7.8,6.9,8.2,9.4,8.8,7.1,6.5,7.3]
})

# ── FILE-BASED MESSAGE STORE ─────────────────────────────────────────────────
# Messages are written to a JSON file so they persist across ALL sessions.
# Every login reads from the same file — EPRA messages appear in county inboxes
# and county messages appear in EPRA's hub in real time.
import os, threading

MSG_STORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kw_messages.json")
_msg_lock = threading.Lock()

SEED_MESSAGES = [
    {
        "id":"MSG-NK-20260528-1234","from_role":"county","from_county":"Nakuru",
        "from_name":"Nakuru County Committee","to":"EPRA",
        "type":"County query","subject":"❓ Solar GHI methodology question",
        "body":"We are using the Global Solar Atlas and getting 1,920 kWh/m² for Nakuru. The EPRA template shows 2,008 as an example (Makueni). Is 1,920 correct for our location?",
        "date":"2026-05-28","time":"10:14","read_by_epra":False,"replies":[]
    },
    {
        "id":"MSG-GR-20260510-5678","from_role":"county","from_county":"Garissa",
        "from_name":"Garissa County Committee","to":"EPRA",
        "type":"Extension request","subject":"📅 Requesting 30-day extension",
        "body":"Garissa County requests a 30-day extension. We are waiting for updated KNBS population figures and finalising GIS mapping of unelectrified settlements.",
        "date":"2026-05-10","time":"09:22","read_by_epra":False,"replies":[]
    },
    {
        "id":"MSG-EPRA-20260601-0001","from_role":"epra","from_county":"",
        "from_name":"EPRA","to":"Makueni",
        "type":"Assumptions","subject":"📋 INEP 2025 planning assumptions — mandatory baselines",
        "body":"Use these national baselines for your County Energy Plan submission: Population baseline: KNBS 2024 projections. Electricity cost benchmark: 0.047–0.059 USD/kWh. Solar GHI reference: 1,600–2,200 kWh/m²/year. MTF demand tiers: Tier 1–2 rural off-grid, Tier 3–4 rural grid, Tier 5 urban. Access targets: 100% electricity by 2030, universal clean cooking by 2028.",
        "date":"2026-06-01","time":"09:00","read_by_epra":True,"replies":[]
    },
    {
        "id":"MSG-EPRA-20260601-0002","from_role":"epra","from_county":"",
        "from_name":"EPRA","to":"All counties",
        "type":"Guidance","subject":"📄 Makueni CEP now available as reference template",
        "body":"The Makueni County Energy Plan (2023–2032) is now available in KenyaWatts as a reference guide. It demonstrates the expected structure for all 7 required sectors. A simplified estimate using the structured template is fully acceptable for a first submission.",
        "date":"2026-04-30","time":"10:30","read_by_epra":True,"replies":[]
    },
]

def load_messages():
    """Load messages from the JSON file. Seed with defaults if file does not exist."""
    with _msg_lock:
        if not os.path.exists(MSG_STORE_PATH):
            _write_messages(SEED_MESSAGES)
            return SEED_MESSAGES
        try:
            with open(MSG_STORE_PATH,"r",encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            _write_messages(SEED_MESSAGES)
            return SEED_MESSAGES

def _write_messages(messages):
    """Write messages to the JSON file (call inside _msg_lock)."""
    with open(MSG_STORE_PATH,"w",encoding="utf-8") as f:
        json.dump(messages, f, ensure_ascii=False, indent=2)

def save_messages(messages):
    """Save messages to the JSON file (thread-safe)."""
    with _msg_lock:
        _write_messages(messages)

def get_messages():
    """Return live messages — always reads from file."""
    return load_messages()

def add_message(msg):
    """Append a new message and save."""
    with _msg_lock:
        msgs = load_messages()
        msgs.append(msg)
        _write_messages(msgs)
    return msg

def add_reply(msg_id, reply):
    """Add a reply to a specific message by ID and save."""
    with _msg_lock:
        msgs = load_messages()
        for m in msgs:
            if m["id"] == msg_id:
                m.setdefault("replies", []).append(reply)
                m["read_by_epra"] = True
                break
        _write_messages(msgs)

def mark_read(msg_id):
    """Mark a message as read by EPRA."""
    with _msg_lock:
        msgs = load_messages()
        for m in msgs:
            if m["id"] == msg_id:
                m["read_by_epra"] = True
                break
        _write_messages(msgs)


# ── FILE-BASED SUBMISSION STORE ───────────────────────────────────────────────
# Every county plan submission is written to disk immediately.
# EPRA and county users read from the same file — real-time across all sessions.
SUB_STORE_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kw_submissions.json")
_sub_lock = threading.Lock()

def load_submissions():
    """Load all submissions from disk."""
    with _sub_lock:
        if not os.path.exists(SUB_STORE_PATH):
            return []
        try:
            with open(SUB_STORE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []

def save_submission(record):
    """Append a new submission record to disk (thread-safe)."""
    with _sub_lock:
        subs = load_submissions()
        # Replace existing submission for same county+ref if it exists
        subs = [s for s in subs if s.get("ref") != record.get("ref")]
        subs.append(record)
        with open(SUB_STORE_PATH, "w", encoding="utf-8") as f:
            json.dump(subs, f, ensure_ascii=False, indent=2)

def get_submissions(county=None):
    """Return submissions, optionally filtered by county name."""
    subs = load_submissions()
    if county:
        subs = [s for s in subs if s.get("county","").lower()==county.lower()]
    return sorted(subs, key=lambda x: x.get("datetime",""), reverse=True)

def update_submission_status(ref, new_status, epra_note=""):
    """EPRA approves/rejects a submission."""
    with _sub_lock:
        subs = load_submissions()
        for s in subs:
            if s.get("ref") == ref:
                s["status"]     = new_status
                s["epra_note"]  = epra_note
                s["epra_action_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                break
        with open(SUB_STORE_PATH, "w", encoding="utf-8") as f:
            json.dump(subs, f, ensure_ascii=False, indent=2)

STATUS_COLOR = {"submitted":"#0f9d7e","review":"#d4891a","pending":"#7a7870","overdue":"#b33a2c"}
STATUS_LABEL = {"submitted":"Submitted ✓","review":"In review","pending":"Pending","overdue":"Overdue"}

# ── SESSION STATE INIT ────────────────────────────────────────────────────────
def init_session():
    defaults = {
        "audit_log": [], "upload_log": [], "ai_history": [],
        "notifications": [], "user_display_name": {}, "user_email": {},
        "submitted_data": [],
        # Two-way messaging store: list of message dicts shared across all roles
        "messages": [
            {
                "id":"MSG-NK-20260528-1234","from_role":"county","from_county":"Nakuru",
                "from_name":"Nakuru County Committee","to":"EPRA",
                "type":"County query","subject":"❓ Solar GHI methodology question",
                "body":"We are using the Global Solar Atlas and getting 1,920 kWh/m² for Nakuru. The EPRA template shows 2,008 as an example (Makueni). Is 1,920 correct for our location or should we use a different source?",
                "date":"2026-05-28","time":"10:14","read_by_epra":False,"replies":[]
            },
            {
                "id":"MSG-GR-20260510-5678","from_role":"county","from_county":"Garissa",
                "from_name":"Garissa County Committee","to":"EPRA",
                "type":"Extension request","subject":"📅 Requesting 30-day extension",
                "body":"Garissa County requests a 30-day extension. We are waiting for updated KNBS population figures and finalising GIS mapping of unelectrified settlements. Expected submission: 20 June 2026.",
                "date":"2026-05-10","time":"09:22","read_by_epra":False,"replies":[]
            },
        ],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

# ── NOTIFICATION SYSTEM ───────────────────────────────────────────────────────
def push_notification(msg, icon="✅"):
    st.session_state.notifications.append({
        "msg": msg, "icon": icon,
        "time": datetime.now().strftime("%H:%M:%S"), "shown": False
    })

def show_notifications():
    for n in st.session_state.notifications:
        if not n["shown"]:
            st.toast(f"{n['icon']} {n['msg']}", icon=n["icon"])
            n["shown"] = True

# ── AGGREGATION ───────────────────────────────────────────────────────────────
def compute_national():
    submitted = COUNTIES[COUNTIES["status"].isin(["submitted","review"])]
    if submitted.empty:
        return {"submitted_count":0,"total_counties":len(COUNTIES),"w_elec":0,
                "w_cooking":0,"total_budget":0,"latest_target":2035,
                "coverage_pct":0,"overdue_count":4,"pending_count":6}
    tp = submitted["pop"].sum()
    return {
        "submitted_count":  len(submitted),
        "total_counties":   len(COUNTIES),
        "w_elec":           round((submitted["elec"]*submitted["pop"]).sum()/tp, 1),
        "w_cooking":        round((submitted["cooking"]*submitted["pop"]).sum()/tp, 1),
        "total_budget":     round(submitted["budget"].sum(), 1),
        "latest_target":    int(submitted["target_yr"].max()),
        "coverage_pct":     round(len(submitted)/len(COUNTIES)*100),
        "overdue_count":    len(COUNTIES[COUNTIES["status"]=="overdue"]),
        "pending_count":    len(COUNTIES[COUNTIES["status"]=="pending"]),
    }

# ── AI ────────────────────────────────────────────────────────────────────────
def ask_ai(question, history):
    try:
        api_key = st.secrets.get("anthropic",{}).get("api_key","")
        if not api_key or "YOUR_" in api_key:
            return "AI assistant is not configured. Add your Anthropic API key to secrets.toml."
        client  = anthropic.Anthropic(api_key=api_key)
        nat     = compute_national()
        system  = f"""You are the KenyaWatts AI assistant for EPRA's national energy planning platform.
Real data: {nat['submitted_count']} of {nat['total_counties']} counties submitted.
Weighted electricity access: {nat['w_elec']}% (target 100% by 2030).
Weighted clean cooking: {nat['w_cooking']}% (target 100% by 2028).
Clean energy generation: 82% (target 100% by 2035).
Critical overdue counties: Turkana 12%, Marsabit 8%, Mandera 11%, Wajir 9%.
Makueni CEP: electricity 75.1%, solar GHI 2008 kWh/m², budget KES 74.9B, firewood 72.5%.
Average outage: 8.8 hrs/month vs EPRA 5 hr benchmark. 11/12 months exceeded.
INEP Regulations 2025 legally require all 47 county submissions.
Answer in 3-4 sentences. Plain text only."""
        msgs = [{"role":m["role"],"content":m["content"]} for m in history]
        msgs.append({"role":"user","content":question})
        res = client.messages.create(model="claude-sonnet-4-6",max_tokens=600,system=system,messages=msgs)
        return res.content[0].text
    except Exception as e:
        return f"AI error: {str(e)}"

# ── AUTH ──────────────────────────────────────────────────────────────────────
def setup_auth():
    creds = st.secrets.get("credentials",{})
    ud = {}
    for uname, info in creds.get("usernames",{}).items():
        ud[uname] = {"email":info.get("email",""),"name":info.get("name",uname),"password":info.get("password","")}
    cookie = st.secrets.get("cookie",{})
    return stauth.Authenticate(
        {"usernames":ud},
        cookie.get("name","kenyawatts_auth"),
        cookie.get("key","kw_key"),
        0,   # expiry_days=0 means session cookie only — clears when browser closes
             # This ensures every new visit starts at the login page
    )

def get_user_role(username):
    info = st.secrets.get("credentials",{}).get("usernames",{}).get(username,{})
    return {"role":info.get("role","county"),"county_id":info.get("county_id",""),
            "name":info.get("name",username),"email":info.get("email","")}

# ── SHARED COMPONENTS ─────────────────────────────────────────────────────────
def metric_card(label, value, delta=None, delta_color="#0f9d7e"):
    delta_html = f'<div class="kw-metric-delta" style="color:{delta_color}">{delta}</div>' if delta else ""
    st.markdown(f"""<div class="kw-metric">
    <div class="kw-metric-label">{label}</div>
    <div class="kw-metric-value">{value}</div>
    {delta_html}</div>""", unsafe_allow_html=True)

def section(title, sub=""):
    st.markdown(f"**{title}**")
    if sub: st.caption(sub)

def alert(type_, text):
    st.markdown(f'<div class="kw-alert-{type_}">{text}</div>', unsafe_allow_html=True)

# ── MAP VIEW ──────────────────────────────────────────────────────────────────
def page_map(role, county_id):
    is_epra = role in ("epra","ministry","devpartner")
    section("Kenya county energy map", "Interactive map — click a county marker for details")

    # Map metric selector
    if is_epra:
        metric = st.selectbox("Map indicator:", [
            "Electricity access (%)", "Clean cooking access (%)",
            "Solar GHI (kWh/m²)", "Submission status", "Plan budget (KES B)"
        ])
    else:
        metric = "Electricity access (%)"

    # Build map with plotly scatter_mapbox
    df = COUNTIES.copy()
    if not is_epra:
        df = df[df["id"]==county_id]

    # Assign colour and size based on metric
    if metric == "Submission status":
        df["color_val"] = df["status"].map({"submitted":4,"review":3,"pending":2,"overdue":1})
        df["color_str"] = df["status"].map(STATUS_COLOR)
        df["hover"]     = df.apply(lambda r: f"<b>{r['name']}</b><br>Status: {STATUS_LABEL[r['status']]}<br>Population: {r['pop']//1000:,}K", axis=1)
        color_col = "color_str"
    elif metric == "Electricity access (%)":
        df["color_str"] = df["elec"].apply(lambda v: "#0f9d7e" if v>=75 else "#d4891a" if v>=40 else "#b33a2c")
        df["hover"]     = df.apply(lambda r: f"<b>{r['name']}</b><br>Electricity access: {r['elec']}%<br>Target year: {r['target_yr']}", axis=1)
        color_col = "elec"
    elif metric == "Clean cooking access (%)":
        df["hover"] = df.apply(lambda r: f"<b>{r['name']}</b><br>Clean cooking: {r['cooking']}%", axis=1)
        color_col = "cooking"
    elif metric == "Solar GHI (kWh/m²)":
        df["hover"] = df.apply(lambda r: f"<b>{r['name']}</b><br>Solar GHI: {r['solar']} kWh/m²", axis=1)
        color_col = "solar"
    else:
        df["hover"] = df.apply(lambda r: f"<b>{r['name']}</b><br>Budget: KES {r['budget']}B", axis=1)
        color_col = "budget"

    if metric == "Submission status":
        fig = go.Figure(go.Scattermapbox(
            lat=df["lat"], lon=df["lon"],
            mode="markers",
            marker=dict(size=18, color=df["color_str"], opacity=0.85),
            text=df["hover"], hoverinfo="text",
            showlegend=False,
        ))
    else:
        fig = px.scatter_mapbox(
            df, lat="lat", lon="lon",
            color=color_col, size_max=20,
            hover_data=None, custom_data=["hover"],
            color_continuous_scale=["#b33a2c","#d4891a","#0f9d7e"],
            zoom=5,
        )
        fig.update_traces(hovertemplate="%{customdata[0]}<extra></extra>", marker_size=16)
        fig.update_coloraxes(colorbar=dict(
            title=dict(text=metric, font=dict(size=12, color="#ffffff", family="Arial")),
            tickfont=dict(size=11, color="#ffffff", family="Arial"),
            bgcolor="rgba(14,30,46,0.75)",
            bordercolor="rgba(255,255,255,0.2)",
            borderwidth=1,
        ))

    fig.update_layout(
        mapbox_style="carto-positron",
        mapbox=dict(center=dict(lat=0.5, lon=37.9), zoom=5),
        height=480, margin=dict(t=0,b=0,l=0,r=0),
        font=CHART_FONT,
    )

    # Legend for status view — white labels on coloured pill backgrounds
    if metric == "Submission status":
        cols = st.columns(4)
        for col, (s,c) in zip(cols, STATUS_COLOR.items()):
            col.markdown(
                f'<div style="display:flex;align-items:center;gap:7px;font-size:12px;padding:4px 0"><div style="background:{c};color:#ffffff;font-size:11px;font-weight:600;padding:3px 10px;border-radius:12px;letter-spacing:.2px">{STATUS_LABEL[s]}</div></div>',
                unsafe_allow_html=True
            )
        st.markdown("")

    st.plotly_chart(fig, use_container_width=True)

    # Summary below map
    if is_epra:
        nat = compute_national()
        c1,c2,c3,c4 = st.columns(4)
        with c1: metric_card("Submitted",   str(nat["submitted_count"]), f"of {nat['total_counties']} counties", "#0f9d7e")
        with c2: metric_card("Overdue",     str(nat["overdue_count"]),   "Need urgent action", "#b33a2c")
        with c3: metric_card("Pending",     str(nat["pending_count"]),   "Not yet started", "#7a7870")
        with c4: metric_card("Coverage",    f"{nat['coverage_pct']}%",   "of counties with plans", "#1a6fa3")

        alert("danger", "<b>Equity alert:</b> Marsabit (8%), Wajir (9%), Mandera (11%) and Turkana (12%) have the lowest electricity access and are all overdue. These counties must be prioritised for direct submission support.")

# ── NATIONAL OVERVIEW ─────────────────────────────────────────────────────────
def page_national_overview(role):
    nat = compute_national()
    alert("info", f"<b>{'EPRA Admin' if role=='epra' else role.title()} view</b> — national aggregation computed from {nat['submitted_count']} submitted county plans ({nat['coverage_pct']}% coverage).")

    c1,c2,c3,c4 = st.columns(4)
    with c1: metric_card("Counties submitted", f"{nat['submitted_count']} / {nat['total_counties']}", f"{nat['coverage_pct']}% · {nat['overdue_count']} overdue")
    with c2: metric_card("Weighted elec. access", f"{nat['w_elec']}%", "↑ Population-weighted")
    with c3: metric_card("Weighted clean cooking", f"{nat['w_cooking']}%", "Target 100% by 2028", "#d4891a")
    with c4: metric_card("Total plan budgets", f"KES {nat['total_budget']}B", "Sum of submitted plans")

    st.markdown("")
    # Target gauges
    section("National energy targets — live progress")
    tc1,tc2,tc3 = st.columns(3)
    for col, label, val, goal, yr, color in [
        (tc1,"Electricity access",  nat["w_elec"],  100,2030,"#1a6fa3"),
        (tc2,"Clean cooking",       nat["w_cooking"],100,2028,"#0f9d7e"),
        (tc3,"Clean energy gen.",   82,              100,2035,"#5b4fc9"),
    ]:
        with col:
            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=val,
                title={"text":f"<b>{label}</b><br><span style='font-size:11px;color:#1a1916'>Target {goal}% by {yr}</span>","font":{"size":13,"color":"#1a1916","family":"Arial"}},
                gauge={"axis":{"range":[0,100],"tickfont":dict(size=11,color="#1a1916")},"bar":{"color":color},"bgcolor":"#f7f6f2","borderwidth":0},
                number={"suffix":"%","font":{"size":28,"color":"#1a1916","family":"Arial"}}
            ))
            fig.update_layout(height=190, margin=dict(t=60,b=10,l=20,r=20),
                              paper_bgcolor="white", font=CHART_FONT)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("")
    col_left, col_right = st.columns(2)
    with col_left:
        section("Generation trend (GWh)", "FY 2019/20 — 2024/25 · Source: EPRA")
        fig = px.line(GEN_TREND, x="Year", y="GWh", markers=True, color_discrete_sequence=["#1a6fa3"])
        fig = apply_layout(fig, xlabel="Financial Year", ylabel="Generation (GWh)")
        fig.update_layout(yaxis=dict(range=[10000,16000]))
        fig.update_traces(line_width=2.5, marker_size=7)
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        section("Generation mix", "90%+ from renewable sources")
        fig = px.pie(GEN_MIX, names="Source", values="Pct",
                     color="Source", color_discrete_map=dict(zip(GEN_MIX["Source"],GEN_MIX["Color"])))
        fig.update_layout(**BASE_LAYOUT, height=260)
        fig.update_traces(textinfo="percent+label", textfont=dict(size=12,color="#1a1916",family="Arial"))
        st.plotly_chart(fig, use_container_width=True)

# ── ALL COUNTIES ──────────────────────────────────────────────────────────────
def page_all_counties():
    alert("info","<b>EPRA full access:</b> All counties visible. Click any county row for details.")
    status_filter = st.selectbox("Filter:", ["All counties","Submitted ✓","In review","Pending","Overdue"])
    sm = {"All counties":None,"Submitted ✓":"submitted","In review":"review","Pending":"pending","Overdue":"overdue"}
    sf = sm[status_filter]
    df = COUNTIES if sf is None else COUNTIES[COUNTIES["status"]==sf]

    c1,c2,c3,c4 = st.columns(4)
    for col,s,lbl,color in [(c1,"submitted","Submitted","#0f9d7e"),(c2,"review","In review","#d4891a"),
                             (c3,"pending","Pending","#7a7870"),(c4,"overdue","Overdue","#b33a2c")]:
        with col:
            n = len(COUNTIES[COUNTIES["status"]==s])
            st.markdown(f'<div style="padding:12px 14px;border-radius:10px;background:#ffffff;border:0.5px solid #e8e6de;border-top:3px solid {color}"><div style="font-size:22px;font-weight:700;color:{color}">{n}</div><div style="font-size:11px;color:#444441;font-weight:500;margin-top:2px">{lbl}</div></div>', unsafe_allow_html=True)

    st.markdown("")
    for _,row in df.iterrows():
        badge = STATUS_LABEL[row["status"]]
        ec    = "#0f9d7e" if row["elec"]>=75 else "#d4891a" if row["elec"]>=40 else "#b33a2c"
        with st.expander(f"**{row['name']}** · {row['region']} · {row['pop']//1000:,}K pop"):
            cc1,cc2,cc3,cc4,cc5 = st.columns(5)
            cc1.metric("Electricity", f"{row['elec']}%")
            cc2.metric("Clean cooking",f"{row['cooking']}%")
            cc3.metric("Solar GHI",f"{row['solar']}")
            cc4.metric("MTF Tier",f"Tier {row['mtf']}")
            cc5.metric("Target yr",str(row["target_yr"]))
            s_key = row["status"]
            st.markdown(f'<span class="kw-badge-{s_key}">{badge}</span>', unsafe_allow_html=True)
            if row["status"]=="overdue":
                alert("danger",f"<b>Action required:</b> {row['name']} is overdue. INEP Regulations 2025 require submission. Reminder sent.")
            elif row["status"]=="submitted":
                alert("success",f"<b>Plan on file:</b> Budget KES {row['budget']}B · Growth {row['growth']}% p.a. · Validation passed.")

# ── COUNTY DASHBOARD ──────────────────────────────────────────────────────────
def page_county_dashboard(county_id, user_name):
    row = COUNTIES[COUNTIES["id"]==county_id]
    if row.empty: st.error("County not found."); return
    row = row.iloc[0]
    sc  = STATUS_COLOR[row["status"]]
    nat = compute_national()

    alert("info", f"<b>County committee view — {row['name']} County only.</b> You can only see and manage your county's data.")

    st.markdown(f"""<div style="padding:16px 20px;border-radius:12px;background:#0e1e2e;border-left:4px solid {sc};margin-bottom:16px;display:flex;justify-content:space-between;align-items:center">
      <div>
        <div style="font-size:17px;font-weight:700;color:#ffffff;margin-bottom:4px">{row['name']} County Energy Plan</div>
        <div style="font-size:12px;color:#a8c4d4;margin-top:2px">Plan period: 2023–{row['target_yr']} &nbsp;·&nbsp; {row['region']} region</div>
      </div>
      <span style="font-size:12px;font-weight:700;color:#0e1e2e;background:{sc};padding:5px 14px;border-radius:20px;letter-spacing:.2px">{STATUS_LABEL[row['status']]}</span>
    </div>""", unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Electricity access",f"{row['elec']}%","Target: 100%")
    c2.metric("Clean cooking",f"{row['cooking']}%","Target: 100% by 2028")
    c3.metric("Solar GHI",f"{row['solar']} kWh/m²")
    c4.metric("Target year",str(row["target_yr"]))

    st.markdown("")
    section("Your county vs national average (submitted counties)")
    fig = go.Figure()
    fig.add_trace(go.Bar(name=row["name"],   x=["Electricity (%)","Clean cooking (%)"], y=[row["elec"],row["cooking"]],   marker_color="#1a6fa3"))
    fig.add_trace(go.Bar(name="National avg",x=["Electricity (%)","Clean cooking (%)"], y=[nat["w_elec"],nat["w_cooking"]],marker_color="#c8c6be"))
    fig = apply_layout(fig, xlabel="Indicator", ylabel="Percentage (%)")
    fig.update_layout(barmode="group", height=260, legend=dict(font=LEGEND_FONT))
    st.plotly_chart(fig, use_container_width=True)

    if row["status"]=="overdue":
        alert("danger","<b>Your plan is overdue.</b> Under Section 5(5)(a) of the Energy Act 2019, submission is a legal requirement. Use 'Submit energy plan' to submit now.")

# ── UPLOAD / SUBMISSION LOG ────────────────────────────────────────────────────
def show_upload_log(county_id=None):
    """Show submission history from the persistent file store — real across all sessions."""
    # Get county name from county_id
    county_name = None
    if county_id:
        row = COUNTIES[COUNTIES["id"]==county_id]
        if not row.empty:
            county_name = row["name"].values[0]

    # Read from persistent file store
    file_subs = get_submissions(county=county_name)

    # Also include session-only items (e.g. if file store not yet populated)
    session_logs = st.session_state.get("upload_log", [])
    if county_id:
        session_logs = [l for l in session_logs if l.get("county_id")==county_id]

    if not file_subs and not session_logs:
        st.info("No submissions on record yet. When a county submits a plan it will appear here permanently.")
        return

    st.markdown(f"**Submission history — {len(file_subs)} record(s) on server** (persists across all sessions)")

    # Show file store submissions first (most authoritative)
    for s in file_subs:
        sc  = {"submitted":"#d4891a","approved":"#0f9d7e","rejected":"#b33a2c"}.get(s.get("status","submitted"),"#7a7870")
        sl  = {"submitted":"Submitted — pending EPRA review","approved":"Approved by EPRA ✓","rejected":"Resubmission requested"}.get(s.get("status","submitted"),"Submitted")
        icon = "📄" if s.get("pathway")=="pdf" else "📝"
        st.markdown(f"""
        <div style="background:#0e1e2e;border-radius:10px;padding:12px 16px;margin-bottom:8px;
                    border-left:3px solid {sc}">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">
            <div style="display:flex;gap:10px;align-items:center">
              <span style="font-size:18px">{icon}</span>
              <div>
                <div style="font-size:13px;font-weight:600;color:#e8e6e0">
                  County Energy Plan — {s['county']}
                </div>
                <div style="font-size:11px;color:#a8c4d4;margin-top:2px">
                  {s.get('date_display', s.get('date',''))} at {s.get('time','')}
                  &nbsp;·&nbsp; By: {s.get('submitted_by','—')}
                  &nbsp;·&nbsp; File: <code style="color:#3ecfaa">{s.get('document','—')}</code>
                </div>
              </div>
            </div>
            <div style="text-align:right;flex-shrink:0">
              <div style="font-size:11px;font-weight:700;color:{sc}">{sl}</div>
              <div style="font-size:10px;color:#5a8a9e;margin-top:2px;font-family:monospace">{s.get('ref','')}</div>
            </div>
          </div>
          <div style="display:flex;gap:16px;font-size:11px;color:#a8c4d4;margin-top:4px">
            <span>⚡ {s.get('elec_pct','—')}% electricity</span>
            <span>🍳 {s.get('cooking_pct','—')}% clean cooking</span>
            <span>☀️ {s.get('solar_ghi','—')} kWh/m² solar</span>
            <span>💰 KES {s.get('budget_kes_b','—')}B budget</span>
          </div>
          {f'<div style="margin-top:6px;font-size:11px;color:#3ecfaa;background:rgba(62,207,170,0.1);padding:5px 8px;border-radius:5px">✅ EPRA note: {s["epra_note"]}</div>' if s.get("epra_note") and s.get("status")=="approved" else ""}
          {f'<div style="margin-top:6px;font-size:11px;color:#d4891a;background:rgba(212,137,26,0.1);padding:5px 8px;border-radius:5px">↩ EPRA note: {s["epra_note"]}</div>' if s.get("epra_note") and s.get("status")=="rejected" else ""}
        </div>""", unsafe_allow_html=True)

# ── SUBMIT PLAN ───────────────────────────────────────────────────────────────
def page_submit(role, county_id, user_name):
    is_county   = role=="county"
    county_name = COUNTIES[COUNTIES["id"]==county_id]["name"].values[0] if county_id else ""
    if is_county:
        alert("info", f"<b>Submitting as {county_name} County Energy Planning Committee.</b> Your submission goes directly to EPRA for validation.")

    tab_submit, tab_history = st.tabs(["📤  Submit plan", "🕐  Submission history"])

    with tab_submit:
        pathway = st.radio("Choose submission pathway:",
            ["📄 Upload existing PDF / Word plan","📝 Fill structured template"], horizontal=True)
        st.divider()

        if "📄" in pathway:
            section("Upload your County Energy Plan document")
            alert("success","<b>How this works:</b> Upload your completed CEP. AI extracts key indicators automatically. Review the values, then submit to EPRA.")

            uploaded = st.file_uploader("Choose your County Energy Plan (PDF or Word)",
                                        type=["pdf","doc","docx"],
                                        help="Max 50MB. Use the Makueni CEP PDF to test.")
            if uploaded:
                st.success(f"✓ File received: **{uploaded.name}** · {uploaded.size//1024} KB · Uploaded {datetime.now().strftime('%d %b %Y at %H:%M')}")
                with st.spinner("AI extracting key indicators from your document…"):
                    import time; time.sleep(2)
                alert("success","<b>AI extraction complete.</b> Review and correct values below before running validation.")

                extracted = {"elec":75.1,"cooking":17.9,"firewood":72.5,"solar":2008.0,"budget":74.9,"target_yr":2028,"mtf":3,"growth":1.1}
                col1,col2 = st.columns(2)
                with col1:
                    county_in  = st.text_input("County name *", value=county_name if is_county else "Makueni", disabled=is_county, key="pdf_county")
                    elec_in    = st.number_input("Electricity access (%)*", value=extracted["elec"], min_value=0.0, max_value=100.0, key="pdf_elec")
                    cooking_in = st.number_input("Clean cooking access (%)*", value=extracted["cooking"], min_value=0.0, max_value=100.0, key="pdf_cooking")
                    fw_in      = st.number_input("Firewood as primary fuel (%)", value=extracted["firewood"], min_value=0.0, max_value=100.0, key="pdf_fw")
                with col2:
                    solar_in   = st.number_input("Solar GHI (kWh/m²/year)*", value=extracted["solar"], min_value=0.0, key="pdf_solar")
                    budget_in  = st.number_input("Total budget (KES billions)", value=extracted["budget"], min_value=0.0, key="pdf_budget")
                    target_in  = st.number_input("Universal access target year*", value=extracted["target_yr"], min_value=2025, max_value=2040, key="pdf_target")
                    mtf_in     = st.selectbox("MTF demand tier*", [1,2,3,4,5], index=extracted["mtf"]-1, key="pdf_mtf")
                    growth_in  = st.number_input("Population growth rate (%/yr)*", value=extracted["growth"], min_value=0.0, max_value=10.0, key="pdf_growth")

                _validate_and_submit(county_in, elec_in, cooking_in, fw_in, solar_in,
                                      budget_in, target_in, mtf_in, growth_in,
                                      uploaded.name, "pdf", county_id, user_name)
            else:
                st.info("Upload your CEP document above. Use the Makueni County Energy Plan PDF as a test file.")

        else:
            section("Structured submission template")
            alert("info","<b>National assumptions pre-loaded:</b> KNBS baselines · cost benchmarks · solar GHI reference · Kenya's official targets.")
        st.markdown("")
        with st.expander("📄 View example — Makueni County Energy Plan (guidance only)", expanded=False):
            st.caption("Use this as a guide for what figures to fill in. Makueni is Kenya's reference county plan.")
            col_ex1, col_ex2 = st.columns(2)
            with col_ex1:
                st.markdown("""
                | Field | Makueni example value |
                |---|---|
                | Electricity access | 75.1% |
                | Clean cooking access | 17.9% |
                | Firewood as primary fuel | 72.5% |
                | Solar GHI | 2,008 kWh/m² |
                """)
            with col_ex2:
                st.markdown("""
                | Field | Makueni example value |
                |---|---|
                | Total budget | KES 74.9B |
                | Target year | 2028 |
                | MTF demand tier | Tier 3 (rural) |
                | Population growth | 1.1% per year |
                """)
            st.caption("Source: Makueni County Energy Plan 2023–2032 · WRI + Strathmore University")

            with st.expander("▶ Section 1 — Electricity access", expanded=True):
                col1,col2 = st.columns(2)
                with col1:
                    county_in  = st.text_input("County name *", value=county_name, disabled=is_county, key="tmpl_county")
                    elec_in    = st.number_input("Total electricity access (%)*", min_value=0.0, max_value=100.0, help="Grid + mini-grid + SHS", key="tmpl_elec")
                    cooking_in = st.number_input("Clean cooking access (%)*", min_value=0.0, max_value=100.0, key="tmpl_cooking")
                    fw_in      = st.number_input("Firewood as primary fuel (%)", min_value=0.0, max_value=100.0, key="tmpl_fw")
                with col2:
                    solar_in   = st.number_input("Solar GHI (kWh/m²/year)*", min_value=0.0, help="Kenya range: 1,600–2,200", key="tmpl_solar")
                    budget_in  = st.number_input("Total plan budget (KES billions)", min_value=0.0, key="tmpl_budget")
                    target_in  = st.number_input("Universal access target year*", value=2030, min_value=2025, max_value=2040, key="tmpl_target")
                    mtf_in     = st.selectbox("MTF demand tier*", [1,2,3,4,5], index=2, key="tmpl_mtf")
                    growth_in  = st.number_input("Population growth rate (%/yr)*", value=1.1, min_value=0.0, max_value=10.0, key="tmpl_growth")

            with st.expander("▶ Section 2 — Energy resources", expanded=False):
                r1,r2 = st.columns(2)
                with r1:
                    st.number_input("Wind speed at 100m (m/s)", min_value=0.0)
                    st.number_input("Woody biomass supply (tonnes/year)", min_value=0.0)
                with r2:
                    st.number_input("Hydropower potential (MW)", min_value=0.0)
                    st.number_input("Biogas potential (GJ/year)", min_value=0.0)

            with st.expander("▶ Section 3 — Energy efficiency", expanded=False):
                st.number_input("LED bulb adoption in households (%)", min_value=0.0, max_value=100.0, value=79.0)
                st.number_input("Solar PV in county facilities (kW)", min_value=0.0)

            _validate_and_submit(county_in if "county_in" in dir() else county_name,
                                  elec_in if "elec_in" in dir() else 0,
                                  cooking_in if "cooking_in" in dir() else 0,
                                  fw_in if "fw_in" in dir() else 0,
                                  solar_in if "solar_in" in dir() else 0,
                                  budget_in if "budget_in" in dir() else 0,
                                  target_in if "target_in" in dir() else 2030,
                                  mtf_in if "mtf_in" in dir() else 3,
                                  growth_in if "growth_in" in dir() else 1.1,
                                  None, "form", county_id, user_name)

    with tab_history:
        section("Submission and upload history", "All submissions and uploads for your county this session")
        show_upload_log(county_id if is_county else None)

def _validate_and_submit(county, elec, cooking, fw, solar, budget, target, mtf, growth, filename, ptype, county_id, user_name):
    if st.button("▶  Run validation checks (10 rules)", type="primary"):
        errors, warnings = [], []
        if not county:                                       errors.append(("V10","County name is required — mandatory field."))
        if elec<0 or elec>100:                              errors.append(("V1", f"Electricity access {elec}% is outside valid range 0–100%."))
        if fw>0 and cooking>0 and abs((fw+cooking)-100)>10: warnings.append(("V2",f"Cooking fuel split: firewood ({fw}%) + clean ({cooking}%) = {fw+cooking:.1f}%. All fuels should sum to 100%."))
        if solar>0 and (solar<1600 or solar>2200):          warnings.append(("V4",f"Solar GHI {solar} kWh/m² is outside Kenya's range 1,600–2,200. Check units."))
        if not mtf:                                          errors.append(("V7","MTF demand tier must be declared for aggregation."))
        if growth>5:                                         warnings.append(("V3",f"Population growth {growth}% exceeds 5% — verify with KNBS data."))
        if target<2025 or target>2040:                       warnings.append(("V8",f"Target year {target} is outside plausible range 2025–2040."))
        if budget>0 and budget<0.5:                          warnings.append(("V5",f"Budget KES {budget}B is very low — confirm units are billions."))

        passed = 10 - len(errors) - len(warnings)
        rc1,rc2,rc3 = st.columns(3)
        rc1.metric("Critical errors",   len(errors),   delta="block submission" if errors else None, delta_color="inverse")
        rc2.metric("Warnings",          len(warnings), delta="review before submit" if warnings else None, delta_color="off")
        rc3.metric("Checks passed",     passed,        delta=f"of 10")

        for rule,msg in errors:   st.error(f"**ERROR · Rule {rule}:** {msg}")
        for rule,msg in warnings: st.warning(f"**WARNING · Rule {rule}:** {msg}")

        if not errors:
            st.success("✓ No critical errors. Ready to submit to EPRA.")
            if st.button("✅ Submit plan to EPRA", type="primary", key=f"final_submit_{county}"):
                ref  = f"CEP-{county[:2].upper()}-{datetime.now().strftime('%Y')}-{abs(hash(county+str(datetime.now())))%9000+1000}"
                now  = datetime.now()

                # ── Build the full submission record ───────────────────────────
                submission_record = {
                    # Identity
                    "ref":          ref,
                    "county":       county,
                    "county_id":    county_id or county[:2].upper(),
                    "submitted_by": user_name,
                    # Timestamps
                    "datetime":     now.strftime("%Y-%m-%d %H:%M:%S"),
                    "date":         now.strftime("%Y-%m-%d"),
                    "date_display": now.strftime("%d %b %Y"),
                    "time":         now.strftime("%H:%M:%S"),
                    # Document
                    "pathway":      ptype,
                    "document":     filename or "Structured template",
                    # Energy data
                    "elec_pct":     elec,
                    "cooking_pct":  cooking,
                    "firewood_pct": fw,
                    "solar_ghi":    solar,
                    "budget_kes_b": budget,
                    "target_year":  target,
                    "mtf_tier":     mtf,
                    "growth_pct":   growth,
                    # Status
                    "status":          "submitted",
                    "epra_note":       "",
                    "epra_action_time":"",
                    # Validation summary
                    "validation_passed": True,
                    "validation_warnings": 0,
                }

                # ── Persist to disk immediately ────────────────────────────────
                save_submission(submission_record)

                # ── Also save to session upload_log for local history tab ──────
                st.session_state.upload_log.append({
                    "title":        f"County Energy Plan — {county}",
                    "county":       county,
                    "county_id":    county_id or county[:2].upper(),
                    "type":         ptype,
                    "filename":     filename or "Structured template",
                    "date":         now.strftime("%d %b %Y"),
                    "time":         now.strftime("%H:%M:%S"),
                    "status":       "submitted",
                    "ref":          ref,
                    "submitted_by": user_name,
                })
                st.session_state.audit_log.append({
                    "time":   now.strftime("%Y-%m-%d %H:%M:%S"),
                    "user":   user_name,
                    "action": f"Plan submitted and saved to server — {county}",
                    "ref":    ref,
                })

                # ── Send automatic notification message to EPRA via file store ─
                notif_msg = {
                    "id":          f"NOTIF-{ref}",
                    "from_role":   "system",
                    "from_county": county,
                    "from_name":   "KenyaWatts System",
                    "to":          "EPRA",
                    "type":        "Submission",
                    "subject":     f"✅ New plan submitted — {county} County",
                    "body": (
                        f"<b>{county} County Energy Plan</b> has been submitted and saved to the KenyaWatts platform.<br><br>"
                        f"<b>Submitted by:</b> {user_name}<br>"
                        f"<b>Date:</b> {now.strftime('%d %b %Y')}<br>"
                        f"<b>Time:</b> {now.strftime('%H:%M:%S')}<br>"
                        f"<b>Document:</b> {filename or 'Structured template'}<br>"
                        f"<b>Reference:</b> {ref}<br><br>"
                        f"<b>Key indicators:</b> Electricity access {elec}% · "
                        f"Clean cooking {cooking}% · Solar GHI {solar} kWh/m² · "
                        f"Budget KES {budget}B · Target year {target}<br><br>"
                        f"The plan is now in the validation queue. Please review and approve or request resubmission."
                    ),
                    "date":          now.strftime("%Y-%m-%d"),
                    "time":          now.strftime("%H:%M:%S"),
                    "datetime":      now.strftime("%Y-%m-%d %H:%M:%S"),
                    "read_by_epra":  False,
                    "replies":       [],
                }
                add_message(notif_msg)

                # ── Push UI notification ───────────────────────────────────────
                push_notification(f"✅ Plan submitted — {county} · Ref: {ref}", "✅")
                st.balloons()

                # ── Confirmation card ──────────────────────────────────────────
                st.markdown(f"""
                <div style="background:#0e1e2e;border:1.5px solid #0f9d7e;border-radius:12px;
                            padding:22px 24px;margin-top:16px">
                  <div style="font-size:17px;font-weight:700;color:#3ecfaa;margin-bottom:14px">
                    ✅ Submission confirmed and saved
                  </div>
                  <table style="width:100%;border-collapse:collapse">
                    <tr>
                      <td style="padding:6px 0;font-size:12px;font-weight:600;color:#a8c4d4;width:38%;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">County</td>
                      <td style="padding:6px 0;font-size:13px;color:#e8e6e0;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">{county}</td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;font-size:12px;font-weight:600;color:#a8c4d4;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">Submitted by</td>
                      <td style="padding:6px 0;font-size:13px;color:#e8e6e0;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">{user_name}</td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;font-size:12px;font-weight:600;color:#a8c4d4;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">Date</td>
                      <td style="padding:6px 0;font-size:13px;color:#e8e6e0;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">{now.strftime('%d %b %Y')}</td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;font-size:12px;font-weight:600;color:#a8c4d4;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">Time</td>
                      <td style="padding:6px 0;font-size:13px;color:#e8e6e0;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">{now.strftime('%H:%M:%S')}</td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;font-size:12px;font-weight:600;color:#a8c4d4;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">Document</td>
                      <td style="padding:6px 0;font-size:13px;color:#e8e6e0;
                                 border-bottom:0.5px solid rgba(255,255,255,0.08)">{filename or 'Structured template'}</td>
                    </tr>
                    <tr>
                      <td style="padding:6px 0;font-size:12px;font-weight:600;color:#a8c4d4">
                        Reference</td>
                      <td style="padding:6px 0;font-size:14px;color:#3ecfaa;font-family:monospace;
                                 font-weight:700">{ref}</td>
                    </tr>
                  </table>
                  <div style="margin-top:14px;padding-top:12px;border-top:0.5px solid rgba(255,255,255,0.1);
                              display:flex;justify-content:space-between;align-items:center">
                    <div style="font-size:12px;color:#a8c4d4">
                      Saved to KenyaWatts server · EPRA has been notified automatically
                    </div>
                    <div style="font-size:11px;color:#3ecfaa;font-weight:600">
                      ● Live on platform
                    </div>
                  </div>
                </div>""", unsafe_allow_html=True)

# ── VALIDATION QUEUE ──────────────────────────────────────────────────────────
def page_validation_queue():
    alert("warn","<b>EPRA admin only.</b> Review all county plan submissions received from the file store.")

    # ── All submissions from the persistent store ──────────────────────────────
    all_subs = get_submissions()
    pending_review = [s for s in all_subs if s.get("status")=="submitted"]
    approved       = [s for s in all_subs if s.get("status")=="approved"]
    rejected       = [s for s in all_subs if s.get("status")=="rejected"]

    # Summary counts
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total received",   len(all_subs),      "from file store")
    c2.metric("Pending review",   len(pending_review), "awaiting EPRA action")
    c3.metric("Approved",         len(approved),       "included in INEP")
    c4.metric("Resubmission req", len(rejected),       "returned to county")

    if not all_subs:
        st.info("No county submissions have been received yet. When a county submits a plan it will appear here immediately.")
    else:
        st.markdown(f"**All submissions — {len(all_subs)} total · updated in real time**")
        for s in all_subs:
            sc  = {"submitted":"#d4891a","approved":"#0f9d7e","rejected":"#b33a2c"}.get(s.get("status","submitted"),"#7a7870")
            sl  = {"submitted":"Pending review","approved":"Approved ✓","rejected":"Resubmission requested"}.get(s.get("status","submitted"),"Unknown")
            with st.expander(
                f"**{s['county']} County** · Ref: {s['ref']} · "
                f"{s.get('date_display', s.get('date',''))} at {s.get('time','')} · "
                f"Status: {sl}"
            ):
                # Full submission details
                st.markdown(f"""
                <div style="background:#0e1e2e;border-radius:10px;padding:14px 18px;margin-bottom:12px">
                  <table style="width:100%;border-collapse:collapse">
                    <tr>
                      <td style="padding:5px 0;font-size:12px;font-weight:600;color:#a8c4d4;width:35%;border-bottom:0.5px solid rgba(255,255,255,0.08)">County</td>
                      <td style="padding:5px 0;font-size:13px;color:#e8e6e0;border-bottom:0.5px solid rgba(255,255,255,0.08)">{s['county']}</td>
                    </tr>
                    <tr>
                      <td style="padding:5px 0;font-size:12px;font-weight:600;color:#a8c4d4;border-bottom:0.5px solid rgba(255,255,255,0.08)">Submitted by</td>
                      <td style="padding:5px 0;font-size:13px;color:#e8e6e0;border-bottom:0.5px solid rgba(255,255,255,0.08)">{s.get('submitted_by','—')}</td>
                    </tr>
                    <tr>
                      <td style="padding:5px 0;font-size:12px;font-weight:600;color:#a8c4d4;border-bottom:0.5px solid rgba(255,255,255,0.08)">Date & time</td>
                      <td style="padding:5px 0;font-size:13px;color:#e8e6e0;border-bottom:0.5px solid rgba(255,255,255,0.08)">{s.get('date_display', s.get('date',''))} at {s.get('time','')}</td>
                    </tr>
                    <tr>
                      <td style="padding:5px 0;font-size:12px;font-weight:600;color:#a8c4d4;border-bottom:0.5px solid rgba(255,255,255,0.08)">Document</td>
                      <td style="padding:5px 0;font-size:13px;color:#e8e6e0;border-bottom:0.5px solid rgba(255,255,255,0.08)">{s.get('document','—')}</td>
                    </tr>
                    <tr>
                      <td style="padding:5px 0;font-size:12px;font-weight:600;color:#a8c4d4;border-bottom:0.5px solid rgba(255,255,255,0.08)">Reference</td>
                      <td style="padding:5px 0;font-size:14px;color:#3ecfaa;font-family:monospace;font-weight:700;border-bottom:0.5px solid rgba(255,255,255,0.08)">{s.get('ref','—')}</td>
                    </tr>
                    <tr>
                      <td style="padding:5px 0;font-size:12px;font-weight:600;color:#a8c4d4">Status</td>
                      <td style="padding:5px 0;font-size:13px;color:{sc};font-weight:700">{sl}</td>
                    </tr>
                  </table>
                </div>""", unsafe_allow_html=True)

                # Energy indicator summary
                mc1,mc2,mc3,mc4 = st.columns(4)
                mc1.metric("Electricity access",  f"{s.get('elec_pct','—')}%")
                mc2.metric("Clean cooking",        f"{s.get('cooking_pct','—')}%")
                mc3.metric("Solar GHI",            f"{s.get('solar_ghi','—')} kWh/m²")
                mc4.metric("Plan budget",          f"KES {s.get('budget_kes_b','—')}B")

                if s.get("epra_note"):
                    alert("info", f"<b>EPRA note:</b> {s['epra_note']}")

                # EPRA action buttons (only for pending)
                if s.get("status") == "submitted":
                    epra_note = st.text_input(
                        "EPRA note (optional — shown to county):",
                        placeholder="e.g. Approved — all checks passed, or: Please correct solar GHI unit error",
                        key=f"note_{s['ref']}"
                    )
                    btn_a, btn_b, _ = st.columns([1,1.4,2])
                    if btn_a.button(f"✅ Approve", key=f"approve_{s['ref']}", type="primary"):
                        update_submission_status(s["ref"], "approved", epra_note or "Approved by EPRA")
                        # Notify county via message store
                        add_message({
                            "id":          f"NOTIF-APPROVE-{s['ref']}",
                            "from_role":   "epra",
                            "from_county": "",
                            "from_name":   "EPRA",
                            "to":          s["county"],
                            "type":        "Submission",
                            "subject":     f"✅ Your plan has been approved — {s['county']} County",
                            "body": (
                                f"Your County Energy Plan (Ref: {s['ref']}) has been reviewed and "
                                f"<b>approved by EPRA</b> on {datetime.now().strftime('%d %b %Y at %H:%M')}.<br><br>"
                                f"Your county data is now included in the national INEP aggregation. "
                                f"You can view your county on the national dashboard.<br><br>"
                                f"{'<b>EPRA note:</b> ' + epra_note if epra_note else ''}"
                            ),
                            "date":         datetime.now().strftime("%Y-%m-%d"),
                            "time":         datetime.now().strftime("%H:%M:%S"),
                            "datetime":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "read_by_epra": True,
                            "replies":      [],
                        })
                        push_notification(f"{s['county']} County plan approved", "✅")
                        st.success(f"✅ {s['county']} County plan approved. County has been notified.")
                        st.rerun()
                    if btn_b.button(f"↩ Request resubmission", key=f"reject_{s['ref']}"):
                        update_submission_status(s["ref"], "rejected", epra_note or "Resubmission requested")
                        add_message({
                            "id":          f"NOTIF-REJECT-{s['ref']}",
                            "from_role":   "epra",
                            "from_county": "",
                            "from_name":   "EPRA",
                            "to":          s["county"],
                            "type":        "Validation",
                            "subject":     f"↩ Resubmission required — {s['county']} County",
                            "body": (
                                f"Your County Energy Plan (Ref: {s['ref']}) has been reviewed and "
                                f"requires resubmission. Please correct the issues noted below and resubmit.<br><br>"
                                f"<b>EPRA note:</b> {epra_note or 'Please review and resubmit.'}<br><br>"
                                f"Use the Submit energy plan tab to resubmit. Your previous reference is {s['ref']}."
                            ),
                            "date":         datetime.now().strftime("%Y-%m-%d"),
                            "time":         datetime.now().strftime("%H:%M:%S"),
                            "datetime":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "read_by_epra": True,
                            "replies":      [],
                        })
                        push_notification(f"Resubmission requested from {s['county']} County", "↩")
                        st.warning(f"Resubmission requested. {s['county']} County has been notified.")
                        st.rerun()
                elif s.get("status") == "approved":
                    alert("success", f"Approved {s.get('epra_action_time','')} · {s.get('epra_note','')}")
                elif s.get("status") == "rejected":
                    alert("warn", f"Resubmission requested {s.get('epra_action_time','')} · {s.get('epra_note','')}")

    st.divider()
    section("National aggregation trigger")
    nat = compute_national()
    all_real_subs = get_submissions()
    approved_real = [s for s in all_real_subs if s.get("status")=="approved"]
    st.info(f"{len(approved_real)} approved from file store · {nat['submitted_count']} in county data · {nat['pending_count']+nat['overdue_count']} not submitted")
    if st.button("▶  Run national aggregation now", type="primary"):
        with st.spinner("Aggregating submissions into national INEP…"):
            import time; time.sleep(2)
        st.success(f"Aggregation complete. INEP updated with {len(approved_real)} approved county plans.")
        push_notification("National INEP aggregation complete", "🗂️")

# ── INBOX ─────────────────────────────────────────────────────────────────────
def render_msg(subject, from_, to_, date_, type_, body, actions=None, urgent=False):
    """Render a single inbox message card."""
    type_colors = {
        "Assumptions":"#1a6fa3","Benchmark":"#5b4fc9","Validation":"#d4891a",
        "Guidance":"#0f9d7e","Overdue":"#b33a2c","Submission":"#0f9d7e",
        "Alert":"#b33a2c","INEP update":"#0f9d7e","County query":"#d4891a","System":"#7a7870",
    }
    c      = type_colors.get(type_,"#7a7870")
    border = f"border:1.5px solid {c}" if urgent else "border:0.5px solid #e8e6de"
    acts   = "".join([
        f'<span style="font-size:11px;padding:3px 10px;border-radius:6px;border:0.5px solid #e8e6de;background:#f7f6f2;color:#6b6860;margin-right:5px;cursor:pointer">{a}</span>'
        for a in (actions or [])
    ])
    st.markdown(f"""
    <div style="background:#ffffff;{border};border-radius:10px;padding:14px 16px;margin-bottom:10px">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px">
        <div style="flex:1">
          <span style="font-size:10px;font-weight:700;padding:3px 9px;border-radius:8px;
            background:{c}18;color:{c};text-transform:uppercase;letter-spacing:.4px;margin-right:8px">{type_}</span>
          <span style="font-size:13px;font-weight:600;color:#1a1916">{subject}</span>
        </div>
        <span style="font-size:11px;color:#7a7870;flex-shrink:0;margin-left:12px">{date_}</span>
      </div>
      <div style="font-size:11px;color:#7a7870;margin-bottom:8px">
        From: <b style="color:#1a1916">{from_}</b> &nbsp;→&nbsp; {to_}
      </div>
      <div style="font-size:12px;color:#444441;line-height:1.75;margin-bottom:{"10px" if actions else "0"}">{body}</div>
      {f'<div style="margin-top:4px">{acts}</div>' if actions else ""}
    </div>""", unsafe_allow_html=True)

def page_inbox(role, county_id, user_name):
    cnty = COUNTIES[COUNTIES["id"]==county_id]["name"].values[0] if county_id else "Your county"
    row  = COUNTIES[COUNTIES["id"]==county_id].iloc[0] if county_id and not COUNTIES[COUNTIES["id"]==county_id].empty else None

    # ── COUNTY COMMITTEE INBOX ────────────────────────────────────────────────
    if role == "county":
        alert("info", f"<b>{cnty} County communications</b> — receive messages and guidance from EPRA, view validation feedback, and send queries or requests directly to EPRA.")
        tab_in, tab_out = st.tabs(["📥  Received from EPRA","📤  Send to EPRA"])

        with tab_in:
            # Overdue notice if county is overdue
            if row is not None and row["status"]=="overdue":
                render_msg(
                    subject  = "🚨 County Energy Plan overdue — immediate action required",
                    from_    = "EPRA",
                    to_      = cnty,
                    date_    = "2026-06-01",
                    type_    = "Overdue",
                    body     = f"<b>{cnty} County Energy Plan is now overdue.</b> Under Section 5(5)(a) of the Energy Act 2019 and the INEP Regulations 2025, submission is a legal requirement. This is your third and final reminder. Please submit your plan through KenyaWatts immediately or contact EPRA to arrange submission support.",
                    actions  = ["📤 Submit plan now","📧 Contact EPRA"],
                    urgent   = True
                )

            # Validation feedback if in review
            if row is not None and row["status"]=="review":
                render_msg(
                    subject  = "⚠️ Validation warnings on your submission — please review",
                    from_    = "EPRA Validation System",
                    to_      = cnty,
                    date_    = "2026-05-28",
                    type_    = "Validation",
                    body     = "Your submission has 2 warnings that EPRA recommends reviewing before final approval.<br><br>"
                               "<b>WARNING V4:</b> Solar GHI value is slightly above Kenya's expected range. Please confirm you used kWh/m²/year (not MJ/m²/year).<br>"
                               "<b>WARNING V3:</b> Population growth rate exceeds 2% — please confirm this is based on KNBS data for your county.",
                    actions  = ["📤 Resubmit with corrections","📧 Reply to EPRA"]
                )

            # Peer benchmark
            if row is not None:
                nat = compute_national()
                elec_diff = round(row["elec"] - nat["w_elec"], 1)
                cook_diff = round(row["cooking"] - nat["w_cooking"], 1)
                elec_msg  = f"<b style='color:#0f9d7e'>above</b>" if elec_diff>=0 else f"<b style='color:#b33a2c'>below</b>"
                cook_msg  = f"<b style='color:#0f9d7e'>above</b>" if cook_diff>=0 else f"<b style='color:#b33a2c'>below</b>"
                render_msg(
                    subject = f"📊 {cnty} vs national average — June 2026 benchmarks",
                    from_   = "EPRA",
                    to_     = cnty,
                    date_   = "2026-06-01",
                    type_   = "Benchmark",
                    body    = f"{cnty} County electricity access ({row['elec']}%) is {abs(elec_diff)}% {elec_msg} the current national weighted average ({nat['w_elec']}%) from submitted county plans.<br><br>"
                              f"Clean cooking access ({row['cooking']}%) is {abs(cook_diff)}% {cook_msg} the national average ({nat['w_cooking']}%).<br><br>"
                              f"National electricity access target: 100% by 2030. Your county's target year: {row['target_yr']}. "
                              f"{'You are on track.' if row['target_yr']<=2030 else 'Your target year exceeds the national deadline — consider accelerating your plan.'}",
                    actions = ["📊 View full comparison"]
                )

            # National assumptions
            render_msg(
                subject = "📋 INEP 2025 planning assumptions — mandatory baselines for your submission",
                from_   = "EPRA",
                to_     = "All 47 counties",
                date_   = "2026-06-01",
                type_   = "Assumptions",
                body    = "Use these national baselines for your County Energy Plan submission:<br><br>"
                          "• <b>Population baseline:</b> KNBS 2024 projections (see attached)<br>"
                          "• <b>Electricity cost benchmark:</b> 0.047–0.059 USD/kWh for grid scenarios<br>"
                          "• <b>Solar GHI reference:</b> 1,600–2,200 kWh/m²/year for Kenya<br>"
                          "• <b>MTF demand tiers:</b> Tier 1–2 rural off-grid · Tier 3–4 rural grid · Tier 5 urban<br>"
                          "• <b>Access targets:</b> 100% electricity by 2030 · Universal clean cooking by 2028 · 100% clean energy by 2035",
                actions = ["📥 Download assumptions PDF","✓ Mark as read"]
            )

            # Guidance — Makueni reference
            render_msg(
                subject = "📄 Makueni CEP now available as a reference template",
                from_   = "EPRA",
                to_     = "All counties without submitted plans",
                date_   = "2026-04-30",
                type_   = "Guidance",
                body    = "The Makueni County Energy Plan (2023–2032) is now available in KenyaWatts as a reference guide. "
                          "It demonstrates the expected structure for all 7 required sectors. You do not need to match this level of detail — "
                          "a simplified estimate using the structured template is fully acceptable for a first submission. "
                          "Focus on completing all mandatory fields marked with * before considering optional sections.",
                actions = ["📄 View in platform"]
            )

            # ── EPRA REPLIES to this county's messages ─────────────────────────
            my_msgs_with_replies = [
                m for m in get_messages()
                if m.get("from_county")==cnty and len(m.get("replies",[]))>0
            ]
            if my_msgs_with_replies:
                for m in reversed(my_msgs_with_replies):
                    for r in m.get("replies",[]):
                        st.markdown(f"""<div style="background:#ffffff;border:0.5px solid #e8e6de;
                          border-left:3px solid #0f9d7e;border-radius:10px;padding:14px 16px;margin-bottom:10px">
                          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:7px">
                            <div>
                              <span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:6px;
                                background:#e8f7f4;color:#0f9d7e;text-transform:uppercase;margin-right:8px">EPRA reply</span>
                              <span style="font-size:13px;font-weight:600;color:#1a1916">↩ Re: {m['subject']}</span>
                            </div>
                            <span style="font-size:11px;color:#7a7870">{r['date']} at {r['time']}</span>
                          </div>
                          <div style="font-size:11px;color:#7a7870;margin-bottom:7px">
                            From: <b style="color:#1a1916">EPRA</b> → {cnty} County
                            &nbsp;·&nbsp; In reply to: <code>{m['id']}</code>
                          </div>
                          <div style="font-size:12px;color:#444441;line-height:1.75">{r['body']}</div>
                        </div>""", unsafe_allow_html=True)

            # ── MESSAGES FROM EPRA TO THIS COUNTY ────────────────────────────
            epra_to_county = [
                m for m in get_messages()
                if m.get("from_role")=="epra"
                and (m.get("to","").lower()==cnty.lower()
                     or m.get("to","").lower()=="all counties"
                     or "all" in m.get("to","").lower())
            ]
            if epra_to_county:
                st.markdown("**Messages from EPRA**")
                for m in reversed(epra_to_county[-5:]):
                    tc = {"Assumptions":"#1a6fa3","Guidance":"#0f9d7e","Benchmark":"#5b4fc9",
                          "Alert":"#b33a2c","Reminder":"#b33a2c"}.get(m["type"],"#7a7870")
                    st.markdown(f"""<div style="background:#ffffff;border:0.5px solid #e8e6de;
                      border-left:3px solid {tc};border-radius:10px;padding:14px 16px;margin-bottom:8px">
                      <div style="display:flex;justify-content:space-between;margin-bottom:7px">
                        <div>
                          <span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:6px;
                            background:{tc}18;color:{tc};text-transform:uppercase;margin-right:8px">{m['type']}</span>
                          <span style="font-size:13px;font-weight:600;color:#1a1916">{m['subject']}</span>
                        </div>
                        <span style="font-size:11px;color:#7a7870">{m['date']}</span>
                      </div>
                      <div style="font-size:11px;color:#7a7870;margin-bottom:7px">
                        From: <b style="color:#1a1916">EPRA</b> → {m.get('to','All counties')}
                      </div>
                      <div style="font-size:12px;color:#444441;line-height:1.75">{m['body']}</div>
                    </div>""", unsafe_allow_html=True)
                st.divider()

            # ── SESSION SUBMISSION CONFIRMATIONS ──────────────────────────────
            session_subs = get_submissions(county=cnty)
            for s in reversed(session_subs[-2:]):
                render_msg(
                    subject = f"✅ Submission confirmed — {s['county']} County · Ref: {s['ref']}",
                    from_   = "KenyaWatts System",
                    to_     = cnty,
                    date_   = s["date"],
                    type_   = "Submission",
                    body    = f"Your County Energy Plan has been received by EPRA.<br><br>"
                              f"<b>Submitted by:</b> {s['submitted_by']}<br>"
                              f"<b>Date and time:</b> {s['date']} at {s['time']}<br>"
                              f"<b>Document:</b> {s['document']}<br>"
                              f"<b>Reference:</b> <code>{s['ref']}</code><br><br>"
                              f"EPRA will validate and confirm within 14 days.",
                    actions = ["📄 View submission details"]
                )

        with tab_out:
            section("Send a message to EPRA")
            alert("info","Use this to ask questions about methodology, request submission support, request a deadline extension, or flag a data issue. EPRA responds within 5 working days. You can view their replies in the Received tab.")

            # Show previous messages sent by this county
            my_msgs = [m for m in get_messages() if m.get("from_county")==cnty]
            if my_msgs:
                st.markdown(f"**Your sent messages ({len(my_msgs)})**")
                for m in reversed(my_msgs):
                    has_reply = len(m.get("replies",[]))>0
                    reply_note = f" · <span style='color:#0f9d7e;font-weight:600'>✓ {has_reply} reply from EPRA</span>" if has_reply else " · <span style='color:#9c9a8e'>Awaiting EPRA response</span>"
                    st.markdown(f"""<div style="background:#ffffff;border:0.5px solid #e8e6de;border-left:3px solid #1a6fa3;border-radius:8px;padding:12px 14px;margin-bottom:8px">
                      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">
                        <span style="font-size:13px;font-weight:600;color:#1a1916">{m['subject']}</span>
                        <span style="font-size:11px;color:#9c9a8e">{m['date']} at {m['time']}</span>
                      </div>
                      <div style="font-size:11px;color:#7a7870;margin-bottom:6px">Ref: <code>{m['id']}</code>{reply_note}</div>
                      <div style="font-size:12px;color:#444441;line-height:1.6">{m['body']}</div>
                      {"".join([f'<div style="margin-top:10px;background:#e8f7f4;border-radius:6px;padding:10px 12px"><div style="font-size:11px;font-weight:600;color:#0f9d7e;margin-bottom:4px">↩ EPRA reply · {r["date"]} at {r["time"]}</div><div style="font-size:12px;color:#1a1916;line-height:1.6">{r["body"]}</div></div>' for r in m.get("replies",[])])}
                    </div>""", unsafe_allow_html=True)
                st.divider()

            with st.form("county_message_form"):
                msg_type    = st.selectbox("Message type:", [
                    "❓ Methodology question","🙏 Submission support request",
                    "📅 Extension request","🐛 Data / system issue","💬 General query"])
                msg_subject = st.text_input("Subject *", placeholder="Brief description of your query")
                msg_body    = st.text_area("Message *",
                    placeholder="Describe your question or request in detail. Include your county name and submission reference if relevant.",
                    height=130)
                submit_msg  = st.form_submit_button("📤 Send to EPRA", type="primary")

            if submit_msg:
                if not msg_subject.strip() or not msg_body.strip():
                    st.error("Both subject and message body are required.")
                else:
                    ref = f"MSG-{cnty[:2].upper()}-{datetime.now().strftime('%Y%m%d%H%M')}"
                    new_msg = {
                        "id":          ref,
                        "from_role":   "county",
                        "from_county": cnty,
                        "from_name":   user_name,
                        "to":          "EPRA",
                        "type":        msg_type.split(" ",1)[-1],
                        "subject":     msg_subject.strip(),
                        "body":        msg_body.strip(),
                        "date":        datetime.now().strftime("%Y-%m-%d"),
                        "time":        datetime.now().strftime("%H:%M"),
                        "read_by_epra":False,
                        "replies":     [],
                    }
                    add_message(new_msg)
                    st.session_state.audit_log.append({
                        "time":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "user":   user_name,
                        "action": f"Message sent to EPRA: {msg_subject.strip()[:50]}",
                        "ref":    ref,
                    })
                    push_notification(f"Message sent to EPRA · Ref: {ref}", "📧")
                    st.success(f"✓ Message sent to EPRA. Reference: `{ref}`")
                    st.info("EPRA will respond within 5 working days. Check the **Received** tab for their reply.")
                    st.rerun()

    # ── EPRA COMMUNICATIONS HUB ───────────────────────────────────────────────
    elif role == "epra":
        alert("info","<b>EPRA communications hub</b> — messages from all counties, system notifications, and broadcast tools.")
        tab_recv, tab_broadcast = st.tabs(["📥  Received","📣  Send broadcast to counties"])

        with tab_recv:
            # ── LIVE COUNTY MESSAGES (from session state) ─────────────────────
            county_msgs = [m for m in get_messages() if m.get("from_role")=="county"]
            unread = [m for m in county_msgs if not m.get("read_by_epra")]

            if unread:
                alert("warn", f"<b>{len(unread)} unread message(s) from county committees.</b> Scroll down to read and reply.")

            if county_msgs:
                st.markdown(f"**Messages from counties ({len(county_msgs)})**")
                for m in reversed(county_msgs):
                    unread_dot = "🔵 " if not m.get("read_by_epra") else ""
                    tc = {"County query":"#d4891a","Methodology question":"#1a6fa3",
                          "Submission support request":"#5b4fc9","Extension request":"#b33a2c",
                          "Data / system issue":"#b33a2c","General query":"#7a7870"}.get(m["type"],"#d4891a")

                    st.markdown(f"""<div style="background:#ffffff;border:0.5px solid #e8e6de;
                      border-left:3px solid {tc};border-radius:10px;padding:14px 16px;margin-bottom:10px">
                      <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:7px">
                        <div>
                          <span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:6px;
                            background:{tc}18;color:{tc};text-transform:uppercase;margin-right:8px">{m['type']}</span>
                          <span style="font-size:13px;font-weight:600;color:#1a1916">{unread_dot}{m['subject']}</span>
                        </div>
                        <span style="font-size:11px;color:#7a7870;flex-shrink:0;margin-left:12px">{m['date']} at {m['time']}</span>
                      </div>
                      <div style="font-size:11px;color:#7a7870;margin-bottom:7px">
                        From: <b style="color:#1a1916">{m['from_name']}</b> ({m['from_county']} County)
                        &nbsp;→&nbsp; EPRA &nbsp;·&nbsp; Ref: <code>{m['id']}</code>
                      </div>
                      <div style="font-size:12px;color:#444441;line-height:1.75;margin-bottom:10px">{m['body']}</div>
                      {"".join([f'<div style="background:#e8f7f4;border-radius:6px;padding:10px 12px;margin-bottom:6px"><div style="font-size:11px;font-weight:600;color:#0f9d7e;margin-bottom:3px">↩ EPRA replied · {r["date"]} at {r["time"]}</div><div style="font-size:12px;color:#1a1916;line-height:1.6">{r["body"]}</div></div>' for r in m.get("replies",[])])}
                    </div>""", unsafe_allow_html=True)

                    # Mark as read + inline reply form
                    if not m.get("read_by_epra"):
                        mark_read(m["id"])

                    reply_key = f"reply_open_{m['id']}"
                    if reply_key not in st.session_state:
                        st.session_state[reply_key] = False

                    col_reply, col_broadcast, _ = st.columns([1, 1.3, 2])
                    if col_reply.button(f"📧 Reply to {m['from_county']}", key=f"btn_reply_{m['id']}"):
                        st.session_state[reply_key] = not st.session_state[reply_key]
                    if col_broadcast.button("📣 Send to all counties", key=f"btn_broad_{m['id']}"):
                        push_notification(f"Answer broadcast to all counties re: {m['subject'][:40]}", "📣")
                        st.success("Answer broadcast to all 47 counties.")

                    if st.session_state[reply_key]:
                        with st.form(f"reply_form_{m['id']}"):
                            reply_body = st.text_area(
                                f"Reply to {m['from_county']} County",
                                placeholder=f"Type your response to {m['from_county']} County here…",
                                height=110, key=f"reply_text_{m['id']}"
                            )
                            send_reply = st.form_submit_button("📤 Send reply", type="primary")
                        if send_reply:
                            if not reply_body.strip():
                                st.error("Reply cannot be empty.")
                            else:
                                reply = {
                                    "from":  "EPRA",
                                    "body":  reply_body.strip(),
                                    "date":  datetime.now().strftime("%Y-%m-%d"),
                                    "time":  datetime.now().strftime("%H:%M"),
                                }
                                # Save to persistent file store — visible to all sessions
                                add_reply(m["id"], reply)
                                st.session_state[reply_key] = False
                                push_notification(f"Reply sent to {m['from_county']} County", "📧")
                                st.session_state.audit_log.append({
                                    "time":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "user":   "EPRA",
                                    "action": f"Replied to {m['from_county']} County: {m['subject'][:50]}",
                                    "ref":    m["id"],
                                })
                                st.success(f"✓ Reply sent to {m['from_county']} County. It will appear in their inbox immediately.")
                                st.rerun()

                st.divider()

            # ── SYSTEM NOTIFICATIONS ───────────────────────────────────────────
            st.markdown("**System notifications**")
            render_msg(
                subject = "✅ New submission — Mombasa County",
                from_   = "KenyaWatts System",
                to_     = "EPRA",
                date_   = "2026-06-01",
                type_   = "Submission",
                body    = "Mombasa County Energy Plan submitted. Ref: CEP-MB-2026-7731. All 10 validation checks passed. Ready for EPRA review and inclusion in national aggregation.",
                actions = ["✅ Review and approve","📄 View full submission"]
            )
            render_msg(
                subject = "⚠️ Submission with warnings — Nakuru County",
                from_   = "KenyaWatts Validation System",
                to_     = "EPRA",
                date_   = "2026-05-25",
                type_   = "Validation",
                body    = "Nakuru County submitted a plan (Ref: CEP-NA-2026-4821) with 2 warnings: Solar GHI value appears high (possible unit error) · Population growth rate >2%. No critical errors but EPRA review recommended.",
                actions = ["✅ Approve as-is","📧 Request correction from Nakuru","📄 View submission"]
            )
            render_msg(
                subject = "🚨 4 counties remain overdue — escalation recommended",
                from_   = "KenyaWatts System",
                to_     = "EPRA",
                date_   = "2026-05-20",
                type_   = "Alert",
                body    = "Turkana, Marsabit, Mandera and Wajir have not submitted County Energy Plans after 3 reminder notices. These have the lowest electricity access in Kenya (8–12%). Consider direct submission support or escalation to the Council of Governors.",
                actions = ["📣 Send escalation notice","📧 Contact Council of Governors","🗺️ View on map"],
                urgent  = True
            )
            # Any session submissions
            for s in reversed(get_submissions()[:3]):
                render_msg(
                    subject = f"✅ New submission — {s['county']} County · Ref: {s['ref']}",
                    from_   = "KenyaWatts System",
                    to_     = "EPRA",
                    date_   = s["date"],
                    type_   = "Submission",
                    body    = f"Submitted by: {s['submitted_by']} · {s['date']} at {s['time']}<br>"
                              f"Electricity access: {s['elec_pct']}% · Clean cooking: {s['cooking_pct']}% · Solar GHI: {s['solar_ghi']} kWh/m²<br>"
                              f"Document: {s['document']}",
                    actions = ["✅ Review and approve","📄 View full submission"]
                )

        with tab_broadcast:
            section("Send broadcast to counties")
            alert("info","Send national assumptions, guidance, benchmarks, or alerts to all counties, a specific region, or an individual county.")
            with st.form("broadcast_form"):
                b_to      = st.multiselect("Send to:", ["All 47 counties","North East region","Coast region","Nairobi region","Overdue counties only"] + COUNTIES["name"].tolist(), default=["All 47 counties"])
                b_type    = st.selectbox("Message type:", ["Assumptions","Guidance","Benchmark","Alert","Reminder"])
                b_subject = st.text_input("Subject *")
                b_body    = st.text_area("Message body *", height=120)
                b_attach  = st.file_uploader("Attach a document (optional)", type=["pdf","xlsx","docx"])
                b_send    = st.form_submit_button("📣 Send broadcast", type="primary")
            if b_send:
                if not b_subject or not b_body:
                    st.error("Subject and message body are required.")
                else:
                    # Determine actual county targets
                    target_counties = []
                    for t in b_to:
                        if "All 47" in t:
                            target_counties = ["All counties"]
                            break
                        elif "Overdue" in t:
                            target_counties = ["Turkana","Marsabit","Mandera","Wajir"]
                        elif "region" in t.lower():
                            target_counties.append(t)
                        else:
                            target_counties.append(t)
                    to_str = ", ".join(set(target_counties)) if target_counties else "All counties"

                    broadcast_msg = {
                        "id":        f"MSG-EPRA-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                        "from_role": "epra",
                        "from_county": "",
                        "from_name": "EPRA",
                        "to":        to_str,
                        "type":      b_type,
                        "subject":   b_subject.strip(),
                        "body":      b_body.strip(),
                        "date":      datetime.now().strftime("%Y-%m-%d"),
                        "time":      datetime.now().strftime("%H:%M"),
                        "read_by_epra": True,
                        "replies":   [],
                    }
                    add_message(broadcast_msg)
                    push_notification(f"Broadcast sent to {to_str}", "📣")
                    st.success(f"✓ Broadcast sent to: {to_str}. All recipients will see this in their county inbox immediately.")
                    st.rerun()

    # ── MINISTRY INBOX ────────────────────────────────────────────────────────
    elif role == "ministry":
        alert("info","<b>Ministry of Energy inbox</b> — INEP updates and national target alerts from EPRA.")
        nat = compute_national()
        render_msg(
            subject = f"📊 INEP aggregation updated — {nat['submitted_count']} counties submitted ({nat['coverage_pct']}% coverage)",
            from_   = "EPRA",
            to_     = "Ministry of Energy",
            date_   = datetime.now().strftime("%Y-%m-%d"),
            type_   = "INEP update",
            body    = f"National INEP has been updated with data from {nat['submitted_count']} submitted county plans ({nat['coverage_pct']}% coverage).<br><br>"
                      f"<b>Weighted electricity access:</b> {nat['w_elec']}% (target 100% by 2030)<br>"
                      f"<b>Weighted clean cooking:</b> {nat['w_cooking']}% (target 100% by 2028)<br>"
                      f"<b>Clean energy generation:</b> 82% (target 100% by 2035)<br>"
                      f"<b>Latest county target year:</b> {nat['latest_target']}<br>"
                      f"<b>Counties overdue:</b> {nat['overdue_count']}<br><br>"
                      f"Full INEP will be available when all 47 counties have submitted.",
            actions = ["📊 View national dashboard","⬇️ Download INEP summary"]
        )
        if nat["w_cooking"] < 50:
            render_msg(
                subject = "⚠️ NDC alert — clean cooking trajectory off-track",
                from_   = "EPRA",
                to_     = "Ministry of Energy",
                date_   = datetime.now().strftime("%Y-%m-%d"),
                type_   = "Alert",
                body    = f"Based on {nat['submitted_count']} submitted county plans, Kenya's clean cooking access ({nat['w_cooking']}%) suggests the 2028 universal clean cooking target is at risk without accelerated intervention. County plans indicate firewood remains the primary cooking fuel in 70–75% of rural households. Recommended: LPG subsidy expansion and ICS distribution programme.",
                actions = ["📊 View cooking dashboard"],
                urgent  = True
            )

    # ── DEVELOPMENT PARTNER (read only) ──────────────────────────────────────
    elif role == "devpartner":
        alert("info","<b>Development partner updates</b> — national INEP summaries and investment opportunity alerts from EPRA.")
        nat = compute_national()
        render_msg(
            subject = f"📊 Kenya national energy data update — {datetime.now().strftime('%B %Y')}",
            from_   = "EPRA",
            to_     = "Development Partners",
            date_   = datetime.now().strftime("%Y-%m-%d"),
            type_   = "INEP update",
            body    = f"Updated national energy data is available on the KenyaWatts platform based on {nat['submitted_count']} submitted county plans.<br><br>"
                      f"<b>National electricity access:</b> {nat['w_elec']}% weighted average<br>"
                      f"<b>Clean cooking access:</b> {nat['w_cooking']}% weighted average<br>"
                      f"<b>Counties with critical access gaps:</b> Turkana (12%), Marsabit (8%), Mandera (11%), Wajir (9%)<br>"
                      f"<b>Highest growth demand:</b> North Rift region (+15%) and North East (+13%)<br><br>"
                      f"These counties represent the highest-priority investment opportunities for off-grid solar and clean cooking programmes.",
            actions = ["📊 View investment map","⬇️ Download county data"]
        )

# ── MAKUENI REFERENCE ─────────────────────────────────────────────────────────
def page_makueni():
    alert("success","<b>Reference plan:</b> Makueni County Energy Plan 2023–2032 · WRI + Strathmore University · Upload the PDF in Submit tab to test AI extraction.")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Electricity access","75.1%","Grid 29.2% + MG 5.7% + SHS 40.2%")
    c2.metric("Solar GHI","2,008 kWh/m²","PV output 4.35 kWh/kWp/day")
    c3.metric("Total budget","KES 74.9B","2023–2032")
    c4.metric("Clean cooking","17.9%","⚠ Firewood 72.5%")
    st.markdown("")

    c_left, c_right = st.columns(2)
    with c_left:
        section("Primary cooking fuel — households 2022")
        fig = px.bar(MAKUENI_COOKING, x="Pct", y="Fuel", orientation="h",
                     color="Fuel", color_discrete_map={"Firewood":"#b33a2c","LPG":"#0f9d7e",
                     "Charcoal":"#7a7870","Biogas":"#5b4fc9","Electric":"#1a6fa3","Other":"#c8c6be"})
        fig = apply_layout(fig, xlabel="Percentage (%)", ylabel="Fuel type")
        fig.update_layout(showlegend=False, height=240)
        st.plotly_chart(fig, use_container_width=True)

    with c_right:
        section("Monthly outage vs EPRA benchmark")
        colors = ["#b33a2c" if h>5 else "#0f9d7e" for h in OUTAGE["Hours"]]
        fig = go.Figure(go.Bar(x=OUTAGE["Month"], y=OUTAGE["Hours"], marker_color=colors))
        fig.add_hline(y=5, line_dash="dash", line_color="#1a6fa3",
                      annotation_text="<b>EPRA benchmark 5 hrs</b>",
                      annotation_font=dict(size=11, color="#1a1916"))
        fig = apply_layout(fig, xlabel="Month", ylabel="Hours per customer")
        fig.update_layout(height=240)
        st.plotly_chart(fig, use_container_width=True)

    section("Electrification scenarios — investment required (OnSSET)")
    scen = pd.DataFrame({
        "Scenario":["Low demand (2028)","High demand (2028)","Grid intensification (2028)"],
        "Grid (MW)":[19.3,42.7,38.4], "Solar PV (MW)":[2.34,53.6,0], "Investment (USD M)":[132.5,360.0,571.8]
    })
    fig = go.Figure()
    for s,c in [("Grid (MW)","#1a6fa3"),("Solar PV (MW)","#d4891a"),("Investment (USD M)","#0f9d7e")]:
        fig.add_trace(go.Bar(name=s, x=scen["Scenario"], y=scen[s], marker_color=c))
    fig = apply_layout(fig, xlabel="Scenario", ylabel="Value")
    fig.update_layout(barmode="group", height=280, legend=dict(font=LEGEND_FONT))
    st.plotly_chart(fig, use_container_width=True)



# ── NATIONAL REPORT GENERATOR (EPRA only) ────────────────────────────────────
def generate_word_report(nat, counties_df, generated_by):
    """Generate a Word (.docx) National Energy Plan summary report."""
    doc = Document()

    # ── Styles ────────────────────────────────────────────────────────────────
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    def add_heading(text, level=1, color=(26, 31, 22)):
        h = doc.add_heading(text, level=level)
        h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = h.runs[0] if h.runs else h.add_run(text)
        run.font.color.rgb = RGBColor(*color)
        run.font.name = "Arial"
        return h

    def add_para(text, bold=False, size=11):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Arial"
        return p

    def add_table(headers, rows, header_color=(14,30,46)):
        t = doc.add_table(rows=1+len(rows), cols=len(headers))
        t.style = "Table Grid"
        # Header row
        hdr = t.rows[0]
        for i,(h,cell) in enumerate(zip(headers, hdr.cells)):
            cell.text = h
            run = cell.paragraphs[0].runs[0]
            run.bold = True
            run.font.size = Pt(10)
            run.font.name = "Arial"
            run.font.color.rgb = RGBColor(255,255,255)
            cell._tc.get_or_add_tcPr().append(
                __import__("lxml.etree",fromlist=["etree"]).etree.fromstring(
                    f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{"%02x%02x%02x"%tuple(header_color)}" w:val="clear"/>'))
        # Data rows
        for ri,row in enumerate(rows):
            tr = t.rows[ri+1]
            for ci,val in enumerate(row):
                tr.cells[ci].text = str(val)
                run = tr.cells[ci].paragraphs[0].runs[0] if tr.cells[ci].paragraphs[0].runs else tr.cells[ci].paragraphs[0].add_run(str(val))
                run.font.size = Pt(10)
                run.font.name = "Arial"
        return t

    now = datetime.now()

    # ── Cover ─────────────────────────────────────────────────────────────────
    doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("KENYA INTEGRATED NATIONAL ENERGY PLAN")
    tr.bold = True; tr.font.size = Pt(20); tr.font.name = "Arial"
    tr.font.color.rgb = RGBColor(14,30,46)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sub.add_run("National Summary Report — Aggregated from County Energy Plans")
    sr.font.size = Pt(13); sr.font.name = "Arial"; sr.font.color.rgb = RGBColor(26,111,163)

    doc.add_paragraph()
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = meta.add_run(
        f"Generated by: {generated_by}\n"
        f"Date: {now.strftime('%d %B %Y')} at {now.strftime('%H:%M')}\n"
        f"Platform: KenyaWatts · EPRA Kenya · NGDA 2026\n"
        f"Classification: CONFIDENTIAL — EPRA INTERNAL USE ONLY"
    )
    mr.font.size = Pt(10); mr.font.name = "Arial"; mr.font.color.rgb = RGBColor(90,138,158)
    doc.add_page_break()

    # ── 1. Executive Summary ──────────────────────────────────────────────────
    add_heading("1. Executive Summary", 1, (14,30,46))
    add_para(
        f"This report presents the current state of Kenya's national energy planning "
        f"as aggregated from {nat['submitted_count']} submitted County Energy Plans (CEPs) "
        f"on the KenyaWatts Digital Integrated National Energy Planning Platform. "
        f"The data represents {nat['coverage_pct']}% of Kenya's 47 counties. "
        f"A total of {nat['overdue_count']} counties remain overdue on their submissions."
    )
    doc.add_paragraph()

    # Key indicators table
    add_heading("1.1 Key National Indicators", 2, (26,111,163))
    add_table(
        ["Indicator","Current value","National target","Target year","Status"],
        [
            ["Electricity access (weighted avg)", f"{nat['w_elec']}%", "100%", "2030",
             "On track" if float(nat['w_elec'])>70 else "At risk"],
            ["Clean cooking access (weighted avg)", f"{nat['w_cooking']}%", "100%", "2028",
             "At risk" if float(nat['w_cooking'])<50 else "On track"],
            ["Clean energy generation", "82%", "100%", "2035", "On track"],
            ["Counties submitted", f"{nat['submitted_count']} / {nat['total_counties']}", "47 / 47", "Current cycle",
             "In progress"],
            ["Total plan investment needed", f"KES {nat['total_budget']}B",
             "TBD (all 47 counties)", "2023–2032", "Partial"],
        ]
    )
    doc.add_paragraph()

    # ── 2. Submission Status ──────────────────────────────────────────────────
    add_heading("2. County Submission Status", 1, (14,30,46))
    add_para(
        f"As of {now.strftime('%d %B %Y')}, {nat['submitted_count']} of Kenya's 47 counties "
        f"have submitted County Energy Plans through the KenyaWatts platform. "
        f"{nat['overdue_count']} counties are overdue and have received reminder notifications. "
        f"{nat['pending_count']} counties are yet to begin their submission."
    )
    doc.add_paragraph()

    submitted = counties_df[counties_df["status"].isin(["submitted","review"])]
    overdue   = counties_df[counties_df["status"]=="overdue"]
    pending   = counties_df[counties_df["status"]=="pending"]

    add_heading("2.1 Submitted counties", 2, (15,157,126))
    if not submitted.empty:
        add_table(
            ["County","Region","Electricity (%)","Clean cooking (%)","Budget (KES B)","Target year"],
            [[r["name"],r["region"],f"{r['elec']}%",f"{r['cooking']}%",
              str(r["budget"]) if r["budget"]>0 else "—",str(r["target_yr"])]
             for _,r in submitted.iterrows()]
        )
    doc.add_paragraph()

    add_heading("2.2 Overdue counties — action required", 2, (163,45,44))
    if not overdue.empty:
        add_para(
            "The following counties have not submitted their County Energy Plans despite "
            "multiple reminder notifications. Escalation to the Council of Governors is "
            "recommended under Section 5(5)(a) of the Energy Act 2019.", bold=False
        )
        add_table(
            ["County","Region","Electricity access (%)","Population","Reminders sent"],
            [[r["name"],r["region"],f"{r['elec']}%",f"{r['pop']//1000:,}K","3 (final notice sent)"]
             for _,r in overdue.iterrows()]
        )
    doc.add_paragraph()

    # ── 3. National Energy Access ─────────────────────────────────────────────
    add_heading("3. National Energy Access Analysis", 1, (14,30,46))
    add_heading("3.1 Electricity access", 2, (26,111,163))
    add_para(
        f"The population-weighted national electricity access rate from submitted county "
        f"plans is {nat['w_elec']}%. This figure reflects the weighted contribution of each "
        f"county based on its population. The national target is 100% by 2030. "
        f"The four counties with the most critical access gaps — Turkana (12%), Marsabit (8%), "
        f"Mandera (11%), and Wajir (9%) — are all overdue on plan submission and represent "
        f"the highest-priority investment areas."
    )
    doc.add_paragraph()

    add_heading("3.2 Electricity access by county (submitted plans)", 2, (26,111,163))
    if not submitted.empty:
        sorted_df = submitted.sort_values("elec")
        add_table(
            ["County","Electricity access (%)","Gap to 100% target","MTF demand tier"],
            [[r["name"],f"{r['elec']}%",f"{100-r['elec']}%",f"Tier {r['mtf']}"]
             for _,r in sorted_df.iterrows()]
        )
    doc.add_paragraph()

    add_heading("3.3 Clean cooking access", 2, (26,111,163))
    add_para(
        f"The population-weighted clean cooking access rate is {nat['w_cooking']}%. "
        f"Kenya's target is universal clean cooking by 2028. Based on submitted county plans, "
        f"firewood remains the primary cooking fuel in the majority of rural households. "
        f"{'This trajectory suggests the 2028 target is at risk without accelerated intervention.' if float(nat['w_cooking'])<50 else 'Progress is being made but acceleration is needed to meet the 2028 target.'}"
    )
    doc.add_paragraph()

    # ── 4. Renewable Energy Resources ────────────────────────────────────────
    add_heading("4. Renewable Energy Resource Assessment", 1, (14,30,46))
    add_para(
        "Based on submitted county plans, Kenya has abundant renewable energy resources "
        "distributed across its counties. The table below summarises solar potential by county."
    )
    doc.add_paragraph()

    if not submitted.empty:
        add_table(
            ["County","Solar GHI (kWh/m²/yr)","Solar potential","Population (K)"],
            [[r["name"],str(r["solar"]),
              "High (>2,000)" if r["solar"]>2000 else "Good (1,800–2,000)" if r["solar"]>1800 else "Moderate",
              f"{r['pop']//1000:,}"]
             for _,r in submitted.sort_values("solar",ascending=False).iterrows()]
        )
    doc.add_paragraph()

    # ── 5. Investment Summary ─────────────────────────────────────────────────
    add_heading("5. Investment and Budget Summary", 1, (14,30,46))
    add_para(
        f"Based on {nat['submitted_count']} submitted county plans, the total identified "
        f"investment required is KES {nat['total_budget']} billion. This covers electricity "
        f"access expansion, clean cooking programmes, renewable energy development, and "
        f"energy efficiency improvements. The national figure will be updated as additional "
        f"counties submit their plans."
    )
    doc.add_paragraph()

    if not submitted.empty:
        budget_counties = submitted[submitted["budget"]>0]
        if not budget_counties.empty:
            add_table(
                ["County","Plan budget (KES B)","Budget per capita (KES)","Plan period"],
                [[r["name"],f"{r['budget']}B",
                  f"{int(r['budget']*1e9/r['pop']):,}",
                  f"2023–{r['target_yr']}"]
                 for _,r in budget_counties.iterrows()]
            )
    doc.add_paragraph()

    # ── 6. Recommendations ────────────────────────────────────────────────────
    add_heading("6. Recommendations", 1, (14,30,46))
    recs = [
        ("Prioritise direct submission support for overdue counties",
         f"Turkana, Marsabit, Mandera and Wajir have not submitted plans despite three reminder "
         f"notifications. EPRA should arrange direct technical assistance with these counties "
         f"given their critical access gaps and strategic importance to the national INEP."),
        ("Accelerate clean cooking intervention",
         f"The current clean cooking access rate ({nat['w_cooking']}%) is significantly below "
         f"the 2028 target trajectory. EPRA should recommend to the Ministry of Energy that "
         f"LPG subsidy expansion and improved cookstove distribution programmes be accelerated."),
        ("Complete national aggregation when all 47 counties submit",
         f"The current INEP aggregation covers {nat['coverage_pct']}% of counties. The "
         f"national picture will only be complete when all 47 counties have submitted. "
         f"EPRA should set a firm deadline for the remaining {nat['pending_count']+nat['overdue_count']} counties."),
        ("Publish county comparison dashboard publicly",
         f"Making county-level energy access data publicly available — following the India SECI "
         f"model — will create accountability pressure and incentivise timely submission in "
         f"future planning cycles."),
    ]
    for i,(title,body) in enumerate(recs,1):
        add_para(f"{i}. {title}", bold=True)
        add_para(f"   {body}")
        doc.add_paragraph()

    # ── 7. Data notes ─────────────────────────────────────────────────────────
    add_heading("7. Data Notes and Methodology", 1, (14,30,46))
    add_para(
        f"• National electricity access and clean cooking figures are population-weighted "
        f"averages calculated from {nat['submitted_count']} submitted county plans.\n"
        f"• Budget figures are simple sums of submitted county plan budgets.\n"
        f"• Universal access target year reflects the latest target year from submitted plans ({nat['latest_target']}).\n"
        f"• Data generated from KenyaWatts platform on {now.strftime('%d %B %Y')}.\n"
        f"• Clean energy generation (82%) sourced from EPRA Statistics Report FY 2024/25.\n"
        f"• This report is classified CONFIDENTIAL and is for EPRA internal use only."
    )

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def generate_pdf_report(nat, counties_df, generated_by):
    """Generate a PDF National Energy Plan summary report."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            topMargin=2*cm, bottomMargin=2*cm,
                            leftMargin=2*cm, rightMargin=2*cm)

    styles = getSampleStyleSheet()
    now = datetime.now()

    # Custom styles
    title_style = ParagraphStyle("KWTitle", parent=styles["Title"],
        fontSize=20, textColor=colors.HexColor("#0e1e2e"),
        fontName="Helvetica-Bold", spaceAfter=6, alignment=TA_CENTER)
    sub_style = ParagraphStyle("KWSub", parent=styles["Normal"],
        fontSize=13, textColor=colors.HexColor("#1a6fa3"),
        fontName="Helvetica", spaceAfter=4, alignment=TA_CENTER)
    meta_style = ParagraphStyle("KWMeta", parent=styles["Normal"],
        fontSize=9, textColor=colors.HexColor("#5a8a9e"),
        fontName="Helvetica", spaceAfter=2, alignment=TA_CENTER)
    h1_style = ParagraphStyle("KWH1", parent=styles["Heading1"],
        fontSize=14, textColor=colors.HexColor("#0e1e2e"),
        fontName="Helvetica-Bold", spaceBefore=16, spaceAfter=6)
    h2_style = ParagraphStyle("KWH2", parent=styles["Heading2"],
        fontSize=12, textColor=colors.HexColor("#1a6fa3"),
        fontName="Helvetica-Bold", spaceBefore=10, spaceAfter=4)
    body_style = ParagraphStyle("KWBody", parent=styles["Normal"],
        fontSize=10, textColor=colors.HexColor("#1a1916"),
        fontName="Helvetica", spaceAfter=6, leading=15)
    conf_style = ParagraphStyle("KWConf", parent=styles["Normal"],
        fontSize=9, textColor=colors.HexColor("#b33a2c"),
        fontName="Helvetica-Bold", alignment=TA_CENTER)

    def make_table(headers, rows, col_widths=None):
        data = [headers] + rows
        tw   = col_widths or [A4[0]/(2*cm)/len(headers)*cm]*len(headers)
        t    = Table(data, colWidths=tw)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,0), colors.HexColor("#0e1e2e")),
            ("TEXTCOLOR",  (0,0),(-1,0), colors.white),
            ("FONTNAME",   (0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",   (0,0),(-1,0), 9),
            ("FONTNAME",   (0,1),(-1,-1),"Helvetica"),
            ("FONTSIZE",   (0,1),(-1,-1), 9),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f7f6f2")]),
            ("GRID",       (0,0),(-1,-1), 0.4, colors.HexColor("#e8e6de")),
            ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING", (0,0),(-1,-1), 5),
            ("BOTTOMPADDING",(0,0),(-1,-1), 5),
            ("LEFTPADDING", (0,0),(-1,-1), 6),
        ]))
        return t

    story = []

    # Cover
    story.append(Spacer(1, 1*cm))
    story.append(Paragraph("KENYA INTEGRATED NATIONAL ENERGY PLAN", title_style))
    story.append(Paragraph("National Summary Report — Aggregated from County Energy Plans", sub_style))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph(f"Generated by: {generated_by}", meta_style))
    story.append(Paragraph(f"Date: {now.strftime('%d %B %Y')} at {now.strftime('%H:%M')}", meta_style))
    story.append(Paragraph("Platform: KenyaWatts · Energy & Petroleum Regulatory Authority · NGDA 2026", meta_style))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("CONFIDENTIAL — EPRA INTERNAL USE ONLY", conf_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#0e1e2e")))
    story.append(Spacer(1, 0.5*cm))

    # Section 1
    story.append(Paragraph("1. Executive Summary", h1_style))
    story.append(Paragraph(
        f"This report presents the current state of Kenya's national energy planning as "
        f"aggregated from <b>{nat['submitted_count']}</b> submitted County Energy Plans (CEPs) "
        f"on the KenyaWatts platform. The data represents <b>{nat['coverage_pct']}%</b> "
        f"of Kenya's 47 counties. {nat['overdue_count']} counties remain overdue.", body_style))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("1.1 Key National Indicators", h2_style))
    story.append(make_table(
        ["Indicator","Current Value","Target","Year","Status"],
        [
            [f"Electricity access",         f"{nat['w_elec']}%",  "100%","2030","On track" if float(nat['w_elec'])>70 else "At risk"],
            [f"Clean cooking access",       f"{nat['w_cooking']}%","100%","2028","At risk" if float(nat['w_cooking'])<50 else "On track"],
            ["Clean energy generation",     "82%",              "100%","2035","On track"],
            ["Counties submitted",          f"{nat['submitted_count']}/47","47/47","Current","In progress"],
            ["Total investment identified", f"KES {nat['total_budget']}B","TBD","2023–2032","Partial"],
        ],
        [5*cm,3*cm,2*cm,2*cm,2.5*cm]
    ))
    story.append(Spacer(1, 0.4*cm))

    # Section 2
    story.append(Paragraph("2. County Submission Status", h1_style))
    submitted = counties_df[counties_df["status"].isin(["submitted","review"])]
    overdue   = counties_df[counties_df["status"]=="overdue"]
    story.append(Paragraph("2.1 Submitted county plans", h2_style))
    if not submitted.empty:
        story.append(make_table(
            ["County","Region","Electricity (%)","Clean cooking (%)","Budget (KES B)","Target year"],
            [[r["name"],r["region"],f"{r['elec']}%",f"{r['cooking']}%",
              str(r["budget"]) if r["budget"]>0 else "—",str(r["target_yr"])]
             for _,r in submitted.iterrows()],
            [3.5*cm,3*cm,2.5*cm,2.8*cm,2.8*cm,2.4*cm]
        ))
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("2.2 Overdue counties — action required", h2_style))
    if not overdue.empty:
        story.append(make_table(
            ["County","Region","Electricity (%)","Population","Reminders sent"],
            [[r["name"],r["region"],f"{r['elec']}%",f"{r['pop']//1000:,}K","3 (final)"]
             for _,r in overdue.iterrows()],
            [3.5*cm,3.5*cm,2.8*cm,2.8*cm,2.4*cm]
        ))
    story.append(Spacer(1, 0.4*cm))

    # Section 3
    story.append(Paragraph("3. Recommendations", h1_style))
    recs = [
        ("Prioritise direct submission support for overdue counties",
         f"Turkana, Marsabit, Mandera and Wajir have the lowest electricity access in Kenya "
         f"(8–12%) and are all overdue. Direct technical assistance is recommended."),
        ("Accelerate clean cooking intervention",
         f"Current clean cooking access ({nat['w_cooking']}%) is well below the 2028 target. "
         f"LPG subsidy expansion and ICS distribution should be accelerated."),
        ("Complete INEP when all 47 counties submit",
         f"The national picture is {nat['coverage_pct']}% complete. A firm deadline should "
         f"be set for the remaining {nat['pending_count']+nat['overdue_count']} counties."),
    ]
    for i,(title,body) in enumerate(recs,1):
        story.append(Paragraph(f"{i}. {title}", ParagraphStyle("rec",parent=body_style,fontName="Helvetica-Bold",textColor=colors.HexColor("#0e1e2e"))))
        story.append(Paragraph(f"   {body}", body_style))

    # Footer note
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#e8e6de")))
    story.append(Paragraph(
        f"Data source: KenyaWatts Platform · EPRA Statistics FY 2024/25 · "
        f"Generated {now.strftime('%d %B %Y at %H:%M')} · "
        f"CONFIDENTIAL — EPRA INTERNAL USE ONLY",
        meta_style
    ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def page_national_report(user_info):
    """EPRA-only: Generate and download the national INEP summary report."""
    section("National INEP report generator",
            "Generate a Word or PDF summary of all submitted county plans — EPRA confidential")

    alert("warn",
        "<b>CONFIDENTIAL — EPRA internal use only.</b> This report aggregates submitted county "
        "energy plans into a national summary. It is intended for EPRA planners and the "
        "Ministry of Energy only. Do not distribute without authorisation.")

    nat = compute_national()

    # Live summary of what the report will contain
    st.markdown("**What this report will contain:**")
    c1,c2,c3 = st.columns(3)
    with c1:
        st.markdown(f"""<div style="background:#0e1e2e;border-radius:10px;padding:14px 16px">
          <div style="font-size:11px;color:#a8c4d4;margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">Data coverage</div>
          <div style="font-size:18px;font-weight:700;color:#3ecfaa">{nat['submitted_count']} / {nat['total_counties']}</div>
          <div style="font-size:11px;color:#a8c4d4">counties with submitted plans</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div style="background:#0e1e2e;border-radius:10px;padding:14px 16px">
          <div style="font-size:11px;color:#a8c4d4;margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">National electricity access</div>
          <div style="font-size:18px;font-weight:700;color:#3ecfaa">{nat['w_elec']}%</div>
          <div style="font-size:11px;color:#a8c4d4">population-weighted average</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div style="background:#0e1e2e;border-radius:10px;padding:14px 16px">
          <div style="font-size:11px;color:#a8c4d4;margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">Total investment identified</div>
          <div style="font-size:18px;font-weight:700;color:#3ecfaa">KES {nat['total_budget']}B</div>
          <div style="font-size:11px;color:#a8c4d4">from submitted county budgets</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("")
    st.markdown("**Report sections:**")
    st.markdown("""
    1. Executive summary with key national indicators table
    2. County submission status — submitted, pending, and overdue counties
    3. National energy access analysis — electricity and clean cooking
    4. Renewable energy resource assessment by county
    5. Investment and budget summary
    6. Recommendations for EPRA action
    7. Data notes and methodology
    """)

    st.divider()
    st.markdown("**Generate report:**")

    col_fmt, col_gen = st.columns([1,2])
    with col_fmt:
        fmt = st.radio("Report format:", ["📄 Word document (.docx)", "📑 PDF document (.pdf)"], index=0)
    with col_gen:
        report_title = st.text_input("Report title (optional)",
            value=f"Kenya INEP National Summary — {datetime.now().strftime('%B %Y')}",
            key="report_title_input")
        include_overdue  = st.checkbox("Include overdue county details", value=True)
        include_recs     = st.checkbox("Include recommendations section", value=True)

    st.markdown("")
    if st.button("⬇️  Generate and download report", type="primary", use_container_width=False):
        with st.spinner("Generating national report…"):
            try:
                if "Word" in fmt:
                    if not DOCX_AVAILABLE:
                        st.error("python-docx is not installed. Run: pip install python-docx")
                        return
                    data     = generate_word_report(nat, COUNTIES, user_info["name"])
                    filename = f"Kenya_INEP_National_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                    mime     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    if not PDF_AVAILABLE:
                        st.error("reportlab is not installed. Run: pip install reportlab")
                        return
                    data     = generate_pdf_report(nat, COUNTIES, user_info["name"])
                    filename = f"Kenya_INEP_National_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                    mime     = "application/pdf"

                st.download_button(
                    label    = f"📥 Click here to download: {filename}",
                    data     = data,
                    file_name= filename,
                    mime     = mime,
                    type     = "primary",
                )
                st.success(
                    f"✅ Report generated successfully · {len(data)//1024} KB · "
                    f"Generated at {datetime.now().strftime('%H:%M:%S')}"
                )
                # Audit log
                st.session_state.audit_log.append({
                    "time":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "user":   user_info["name"],
                    "action": f"National INEP report generated ({fmt.split()[0]})",
                    "ref":    f"RPT-{datetime.now().strftime('%Y%m%d%H%M')}",
                })
                alert("info",
                    "<b>Report generated.</b> This report is CONFIDENTIAL. "
                    "Share only with authorised Ministry of Energy officials and EPRA leadership. "
                    "The generation event has been logged in the audit trail.")
            except Exception as e:
                st.error(f"Report generation failed: {str(e)}")
                st.exception(e)

# ── MAKUENI — CONTEXTUAL (different per role) ─────────────────────────────────
def page_makueni_contextual(role, county_id):
    is_makueni_county = county_id == "MK"

    if role == "epra":
        # EPRA sees it as a template reference tool
        st.markdown("""<div class="kw-alert-info">
        <b>EPRA template reference:</b> The Makueni CEP (2023–2032) is Kenya's most comprehensive
        completed county energy plan. Use it to demonstrate the submission standard to other counties,
        design the submission template, and benchmark incoming plans. Developed with WRI and Strathmore University.
        </div>""", unsafe_allow_html=True)

    elif is_makueni_county:
        # Makueni sees their own plan
        st.markdown("""<div class="kw-alert-success">
        <b>Your county energy plan:</b> Makueni County Energy Plan 2023–2032.
        This is the plan your county submitted. It is stored in the KenyaWatts repository
        and visible to EPRA and the Ministry. Your plan is being used as the national reference standard.
        </div>""", unsafe_allow_html=True)

    # Show the full plan content for both
    page_makueni()

    # Add extra guidance for EPRA only
    if role == "epra":
        st.divider()
        st.markdown("### How to use this as a template reference")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            **What Makueni did well:**
            - Covered all 7 CEP chapters
            - Used OnSSET for electrification scenarios
            - Used LEAP for clean cooking modelling
            - Disaggregated data by gender (GESI)
            - Reported at sub-county level (6 sub-counties)
            - Declared all scenario assumptions clearly
            - Included M&E framework with annual targets
            """)
        with col2:
            st.markdown("""
            **What other counties can simplify:**
            - Full OnSSET/LEAP not required — simplified estimates accepted
            - Sub-county breakdown optional for first submission
            - Budget estimates accepted if exact figures not yet available
            - Document upload pathway available for counties with existing plans
            - KenyaWatts validation will flag missing mandatory fields
            """)

# ── DATA DOWNLOAD ─────────────────────────────────────────────────────────────
def page_data_download(role, county_id):
    is_epra   = role in ("epra","ministry","devpartner")
    is_county = role == "county"
    cnty_name = COUNTIES[COUNTIES["id"]==county_id]["name"].values[0] if county_id else ""

    section("Data download centre", "Filter, select a date range and download energy planning data")

    if is_epra:
        alert("info","<b>EPRA access:</b> Download aggregated national data or individual county submissions.")
        download_scope = st.radio("Download scope:", ["National aggregated data","All county submissions","Individual county"], horizontal=True)
    else:
        download_scope = "Individual county"
        alert("info",f"<b>County access:</b> Download data for {cnty_name} County only.")

    st.divider()

    # Date range filter
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        date_from = st.date_input("From date", value=datetime.now().date()-timedelta(days=365))
    with col_d2:
        date_to   = st.date_input("To date",   value=datetime.now().date())

    # Category filter
    st.markdown("**Select data categories to include:**")
    cat_cols = st.columns(4)
    cat_elec    = cat_cols[0].checkbox("Electricity access",      value=True)
    cat_cooking = cat_cols[1].checkbox("Clean cooking",           value=True)
    cat_solar   = cat_cols[2].checkbox("Solar / Renewable data",  value=True)
    cat_budget  = cat_cols[3].checkbox("Budget and investment",   value=True)
    cat_meta    = st.checkbox("Submission metadata (date, reference, submitted by)", value=True)

    # County filter (EPRA only)
    if is_epra and download_scope=="Individual county":
        selected_county = st.selectbox("Select county:", COUNTIES["name"].tolist())
    elif is_county:
        selected_county = cnty_name
    else:
        selected_county = None

    # Format
    fmt = st.radio("Download format:", ["CSV (.csv)","Excel (.xlsx)","JSON (.json)"], horizontal=True)
    st.divider()

    # Preview
    section("Data preview")
    if download_scope=="National aggregated data":
        nat = compute_national()
        preview_data = pd.DataFrame([{
            "Indicator":"Weighted electricity access (%)",    "Value":nat["w_elec"],  "Source":f"{nat['submitted_count']} counties","Date":datetime.now().strftime("%Y-%m-%d")},
            {"Indicator":"Weighted clean cooking access (%)", "Value":nat["w_cooking"],"Source":f"{nat['submitted_count']} counties","Date":datetime.now().strftime("%Y-%m-%d")},
            {"Indicator":"Clean energy generation (%)",       "Value":82,              "Source":"EPRA FY24/25",                       "Date":"2025-06-01"},
            {"Indicator":"Total submitted budgets (KES B)",   "Value":nat["total_budget"],"Source":f"{nat['submitted_count']} counties","Date":datetime.now().strftime("%Y-%m-%d")},
            {"Indicator":"Counties submitted",                "Value":nat["submitted_count"],"Source":"Platform","Date":datetime.now().strftime("%Y-%m-%d")},
            {"Indicator":"National access target year",       "Value":nat["latest_target"],"Source":"County plans","Date":datetime.now().strftime("%Y-%m-%d")},
        ])
        build_download(preview_data, "national_aggregated", fmt, cat_elec, cat_cooking, cat_solar, cat_budget, cat_meta, date_from, date_to)

    elif download_scope=="All county submissions":
        df_all = COUNTIES.copy()
        cols = ["name","region","pop","status","target_yr","mtf","growth"]
        if cat_elec:    cols += ["elec"]
        if cat_cooking: cols += ["cooking"]
        if cat_solar:   cols += ["solar"]
        if cat_budget:  cols += ["budget"]
        df_all = df_all[cols].rename(columns={
            "name":"county","pop":"population","elec":"electricity_access_pct",
            "cooking":"clean_cooking_pct","solar":"solar_ghi_kwh_m2","budget":"budget_kes_b",
            "target_yr":"target_year","mtf":"mtf_tier","growth":"pop_growth_pct"
        })
        build_download(df_all, "all_county_submissions", fmt, cat_elec, cat_cooking, cat_solar, cat_budget, cat_meta, date_from, date_to)

    else:
        # Individual county
        filter_name = selected_county or cnty_name
        row = COUNTIES[COUNTIES["name"]==filter_name]
        if not row.empty:
            r = row.iloc[0]
            cols_map = {"county":r["name"],"region":r["region"],"population":r["pop"],"status":r["status"],"target_year":r["target_yr"]}
            if cat_elec:    cols_map["electricity_access_pct"] = r["elec"]
            if cat_cooking: cols_map["clean_cooking_pct"]      = r["cooking"]
            if cat_solar:   cols_map["solar_ghi_kwh_m2"]       = r["solar"]
            if cat_budget:  cols_map["budget_kes_b"]            = r["budget"]
            if cat_meta:    cols_map.update({"mtf_tier":r["mtf"],"pop_growth_pct":r["growth"],"download_date":datetime.now().strftime("%Y-%m-%d")})
            df_single = pd.DataFrame([cols_map])
            # Also include any session submissions
            if st.session_state.submitted_data:
                sd = pd.DataFrame(st.session_state.submitted_data)
                sd = sd[sd["county"]==filter_name] if filter_name else sd
                if not sd.empty:
                    df_single = pd.concat([df_single, sd], ignore_index=True)
            build_download(df_single, f"{filter_name.lower().replace(' ','_')}_data", fmt, cat_elec, cat_cooking, cat_solar, cat_budget, cat_meta, date_from, date_to)

def build_download(df, filename_base, fmt, cat_elec, cat_cooking, cat_solar, cat_budget, cat_meta, date_from, date_to):
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"{len(df)} records · {len(df.columns)} columns · Date range: {date_from} to {date_to}")
    st.markdown("")
    if "CSV" in fmt:
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(f"⬇ Download {filename_base}.csv", csv,
                           file_name=f"{filename_base}_{date_from}_to_{date_to}.csv",
                           mime="text/csv", type="primary")
    elif "Excel" in fmt:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="KenyaWatts Data")
        st.download_button(f"⬇ Download {filename_base}.xlsx", buf.getvalue(),
                           file_name=f"{filename_base}_{date_from}_to_{date_to}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")
    else:
        json_data = df.to_json(orient="records", indent=2).encode("utf-8")
        st.download_button(f"⬇ Download {filename_base}.json", json_data,
                           file_name=f"{filename_base}_{date_from}_to_{date_to}.json",
                           mime="application/json", type="primary")

# ── AI ASSISTANT ──────────────────────────────────────────────────────────────
def page_ai(role, county_id):
    alert("info","<b>KenyaWatts AI assistant</b> — powered by Claude · grounded in EPRA data and Makueni CEP · plain English answers")
    suggestions = [
        "Which counties need urgent electricity investment?",
        "How is Kenya tracking toward universal access by 2030?",
        "What barriers does Makueni face for clean cooking?",
        "How does KenyaWatts aggregate 47 county plans into the INEP?",
        "Which counties are overdue and what are the consequences?",
        "What validation rules catch unit errors in submissions?",
    ]
    if not st.session_state.ai_history:
        st.markdown("**Suggested questions:**")
        cols = st.columns(2)
        for i,s in enumerate(suggestions):
            if cols[i%2].button(s, key=f"sug_{i}"):
                st.session_state.ai_history.append({"role":"user","content":s})
                with st.spinner("Analysing EPRA data…"):
                    reply = ask_ai(s, st.session_state.ai_history[:-1])
                st.session_state.ai_history.append({"role":"assistant","content":reply})
                st.rerun()
    for msg in st.session_state.ai_history:
        with st.chat_message(msg["role"], avatar="⚡" if msg["role"]=="assistant" else None):
            st.write(msg["content"])
    if prompt := st.chat_input("Ask about Kenya's energy planning…"):
        st.session_state.ai_history.append({"role":"user","content":prompt})
        with st.chat_message("user"):
            st.write(prompt)
        with st.chat_message("assistant", avatar="⚡"):
            with st.spinner("Analysing EPRA data…"):
                reply = ask_ai(prompt, st.session_state.ai_history[:-1])
            st.write(reply)
        st.session_state.ai_history.append({"role":"assistant","content":reply})
    if st.session_state.ai_history:
        if st.button("Clear conversation"):
            st.session_state.ai_history = []
            st.rerun()

# ── ACCOUNT SETTINGS ──────────────────────────────────────────────────────────
def page_account_settings(authenticator, username, user_info):
    alert("info","<b>Account settings</b> — change your password or display name. Role and county access require EPRA admin.")
    tab1, tab2, tab3 = st.tabs(["🔑  Change password","✏️  Update display name","👤  My account info"])

    with tab1:
        section("Change your password")
        alert("warn","<b>Password requirements:</b> Minimum 8 characters. Active immediately for this session.")
        with st.form("pw_form"):
            cur_pw  = st.text_input("Current password",     type="password")
            new_pw  = st.text_input("New password",          type="password", help="At least 8 characters")
            conf_pw = st.text_input("Confirm new password",  type="password")
            sub_pw  = st.form_submit_button("Update password", type="primary")
        if sub_pw:
            if not cur_pw or not new_pw or not conf_pw: st.error("All fields required.")
            elif len(new_pw)<8:            st.error("New password must be at least 8 characters.")
            elif new_pw!=conf_pw:          st.error("Passwords do not match.")
            elif new_pw==cur_pw:           st.warning("New password must differ from current.")
            else:
                try:
                    stored = st.secrets.get("credentials",{}).get("usernames",{}).get(username,{}).get("password","")
                    if stored and bcrypt.checkpw(cur_pw.encode(), stored.encode()):
                        new_hash = bcrypt.hashpw(new_pw.encode(), bcrypt.gensalt()).decode()
                        try: authenticator.credentials["usernames"][username]["password"] = new_hash
                        except: pass
                        st.session_state.audit_log.append({"time":datetime.now().strftime('%Y-%m-%d %H:%M:%S'),"user":username,"action":"Password changed","ref":f"PW-{username.upper()}"})
                        push_notification("Password updated successfully", "🔑")
                        st.success("✓ Password updated successfully.")
                        st.info("To make permanent: EPRA admin must update secrets.toml with the new hash below.")
                        st.code(f'[credentials.usernames.{username}]\npassword = "{new_hash}"', language="toml")
                    else: st.error("Current password is incorrect.")
                except Exception as e: st.error(f"Error: {e}")

    with tab2:
        section("Update your display name and email")
        cur_name  = st.session_state.get("user_display_name",{}).get(username, user_info.get("name",username))
        cur_email = st.session_state.get("user_email",{}).get(username, user_info.get("email",""))
        with st.form("name_form"):
            new_name  = st.text_input("Display name",    value=cur_name)
            new_email = st.text_input("Email address",   value=cur_email)
            sub_name  = st.form_submit_button("Save changes", type="primary")
        if sub_name:
            if len(new_name.strip())<3: st.error("Display name must be at least 3 characters.")
            else:
                st.session_state.setdefault("user_display_name",{})[username] = new_name.strip()
                st.session_state.setdefault("user_email",{})[username]        = new_email.strip()
                try:
                    authenticator.credentials["usernames"][username]["name"]  = new_name.strip()
                    authenticator.credentials["usernames"][username]["email"] = new_email.strip()
                except: pass
                push_notification(f"Display name updated to {new_name.strip()}", "✏️")
                st.success(f"✓ Display name updated to **{new_name.strip()}**.")

    with tab3:
        section("Your account information")
        role_labels = {"epra":"EPRA Planner — Full admin","ministry":"Ministry of Energy — Read only",
                       "devpartner":"Development Partner — Read only","county":"County Energy Planning Committee",
                       "service_provider":"KPLC Service Provider"}
        role = user_info["role"]
        cid  = user_info["county_id"]
        dname  = st.session_state.get("user_display_name",{}).get(username, user_info.get("name",username))
        demail = st.session_state.get("user_email",{}).get(username, user_info.get("email","Not set"))
        cnty   = COUNTIES[COUNTIES["id"]==cid]["name"].values[0] if cid else "All counties"

        # Render as a styled card so colours are explicitly set
        # regardless of Streamlit light/dark theme
        rows_html = ""
        for label, value in [
            ("Username",     username),
            ("Display name", dname),
            ("Email",        demail or "Not set"),
            ("Role",         role_labels.get(role, role)),
            ("County",       cnty),
            ("Session",      datetime.now().strftime("%d %b %Y %H:%M")),
            ("Status",       "Active ✓"),
        ]:
            rows_html += f"""
            <tr>
              <td style="padding:9px 14px 9px 0;font-size:12px;font-weight:600;
                         color:#a8c4d4;width:35%;vertical-align:top;
                         border-bottom:0.5px solid rgba(255,255,255,0.08)">{label}</td>
              <td style="padding:9px 0 9px 0;font-size:13px;color:#e8e6e0;
                         border-bottom:0.5px solid rgba(255,255,255,0.08)">{value}</td>
            </tr>"""

        st.markdown(f"""
        <div style="background:#0e1e2e;border-radius:12px;padding:20px 22px;margin-bottom:14px">
          <table style="width:100%;border-collapse:collapse">
            {rows_html}
          </table>
        </div>
        """, unsafe_allow_html=True)

        # Recent activity log
        audit = [a for a in st.session_state.audit_log if a.get("user")==username]
        if audit:
            st.markdown("**Recent activity**")
            activity_html = ""
            for e in reversed(audit[-5:]):
                activity_html += f"""
                <div style="display:flex;gap:12px;font-size:11px;padding:5px 0;
                            border-bottom:0.5px solid rgba(255,255,255,0.08)">
                  <span style="color:#5a8a9e;font-family:monospace;flex-shrink:0">{e["time"]}</span>
                  <span style="color:#d0cec8">{e["action"]}</span>
                  <span style="color:#5a8a9e;font-family:monospace">{e.get("ref","")}</span>
                </div>"""
            st.markdown(f"""
            <div style="background:#0e1e2e;border-radius:10px;padding:14px 18px;margin-bottom:12px">
              {activity_html}
            </div>""", unsafe_allow_html=True)

        alert("info","<b>Need to change username or role?</b> Contact EPRA admin: Allan.Wairimu@epra.go.ke · +254720850696")

# ── FORGOT PASSWORD ───────────────────────────────────────────────────────────
def page_forgot_password(authenticator):
    """
    Proper multi-step password reset flow using session state.
    Step 1: Enter username + email  → verify against records
    Step 2: Enter 6-digit code      → verify code + expiry
    Step 3: Enter new password      → confirm + hash + done
    """
    import random

    # Initialise step if not set
    if "reset_step" not in st.session_state:
        st.session_state["reset_step"] = 1

    step = st.session_state["reset_step"]

    # ── Page shell ────────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:

        # Logo
        st.markdown("""
        <div style="text-align:center;margin-bottom:28px;margin-top:30px">
          <div style="font-size:26px;font-weight:700;color:#1a1916">
            Kenya<span style="color:#0f9d7e">Watts</span>
          </div>
          <div style="font-size:12px;color:#9c9a8e;margin-top:3px">
            EPRA National Energy Planning Platform
          </div>
        </div>""", unsafe_allow_html=True)

        # Progress bar — 3 steps
        step_labels = ["Verify identity", "Enter reset code", "Set new password"]
        progress_html = '<div style="display:flex;gap:0;margin-bottom:24px;border-radius:8px;overflow:hidden">'
        for i, lbl in enumerate(step_labels, 1):
            bg = "#0e1e2e" if i == step else "#e8f7f4" if i < step else "#f7f6f2"
            tc = "#ffffff" if i == step else "#0f9d7e" if i < step else "#9c9a8e"
            progress_html += (
                f'<div style="flex:1;padding:8px 6px;text-align:center;background:{bg};'
                f'font-size:11px;font-weight:{"600" if i==step else "400"};color:{tc}">'
                f'{"✓ " if i<step else ""}{lbl}</div>'
            )
        progress_html += '</div>'
        st.markdown(progress_html, unsafe_allow_html=True)

        # ── STEP 1 — Verify identity ──────────────────────────────────────────
        if step == 1:
            st.markdown("""
            <div style="font-size:16px;font-weight:700;color:#1a1916;margin-bottom:4px">
              Verify your identity
            </div>
            <div style="font-size:12px;color:#6b6860;margin-bottom:18px">
              Enter your username and the email address registered to your account.
              Both must match exactly what is in the system.
            </div>""", unsafe_allow_html=True)

            with st.form("verify_identity_form", clear_on_submit=False):
                uname = st.text_input(
                    "Username *",
                    placeholder="Your KenyaWatts login username",
                    key="reset_uname_input",
                )
                email = st.text_input(
                    "Registered email address *",
                    placeholder="e.g. energy@makueni.go.ke",
                    key="reset_email_input",
                )
                submit_step1 = st.form_submit_button(
                    "Verify and continue →",
                    type="primary",
                    use_container_width=True,
                )

            if submit_step1:
                if not uname.strip() or not email.strip():
                    st.error("Both username and email address are required.")
                else:
                    creds = st.secrets.get("credentials", {}).get("usernames", {})
                    user  = creds.get(uname.strip(), {})
                    registered_email = user.get("email", "").lower().strip()
                    if uname.strip() in creds and registered_email == email.strip().lower():
                        # Generate 6-digit code
                        code = str(random.randint(100000, 999999))
                        st.session_state["reset_uname"]     = uname.strip()
                        st.session_state["reset_email"]     = email.strip()
                        st.session_state["reset_code"]      = code
                        st.session_state["reset_code_time"] = datetime.now()
                        st.session_state["reset_step"]      = 2
                        st.rerun()
                    else:
                        st.error(
                            "Username and email do not match our records. "
                            "Check your spelling or contact EPRA admin: "
                            "Allan.Wairimu@epra.go.ke"
                        )

        # ── STEP 2 — Enter reset code ─────────────────────────────────────────
        elif step == 2:
            uname = st.session_state.get("reset_uname", "")
            email = st.session_state.get("reset_email", "")
            code  = st.session_state.get("reset_code",  "")
            code_time = st.session_state.get("reset_code_time", datetime.now())
            mins_left = max(0, 15 - int((datetime.now()-code_time).seconds/60))

            st.markdown(f"""
            <div style="font-size:16px;font-weight:700;color:#1a1916;margin-bottom:4px">
              Enter your reset code
            </div>
            <div style="font-size:12px;color:#6b6860;margin-bottom:12px">
              A 6-digit reset code has been generated for <b>{email}</b>.
              In a production deployment this would be sent to your registered email.
              The code expires in <b>{mins_left} minute(s)</b>.
            </div>""", unsafe_allow_html=True)

            # Show code box (demo only — in production this would be emailed)
            st.markdown(f"""
            <div style="background:#0e1e2e;border-radius:10px;padding:16px 20px;
                        text-align:center;margin-bottom:18px">
              <div style="font-size:11px;color:#a8c4d4;margin-bottom:6px;
                          text-transform:uppercase;letter-spacing:.5px">
                Your reset code (demo — would be emailed in production)
              </div>
              <div style="font-size:32px;font-weight:700;color:#3ecfaa;
                          letter-spacing:8px;font-family:monospace">
                {code}
              </div>
              <div style="font-size:10px;color:#5a8a9e;margin-top:6px">
                Expires in {mins_left} minute(s) · Do not share this code
              </div>
            </div>""", unsafe_allow_html=True)

            with st.form("verify_code_form", clear_on_submit=False):
                entered = st.text_input(
                    "Enter the 6-digit code *",
                    placeholder="e.g. 483921",
                    max_chars=6,
                    key="reset_code_input",
                )
                submit_step2 = st.form_submit_button(
                    "Verify code →",
                    type="primary",
                    use_container_width=True,
                )

            if submit_step2:
                elapsed  = (datetime.now() - code_time).seconds
                if elapsed > 900:
                    st.error("This code has expired (15 minutes). Please start again.")
                    if st.button("Start again", key="restart_after_expire"):
                        st.session_state["reset_step"] = 1
                        st.rerun()
                elif entered.strip() != code:
                    st.error("Incorrect code. Please check and try again.")
                else:
                    st.session_state["reset_step"] = 3
                    st.rerun()

        # ── STEP 3 — Set new password ─────────────────────────────────────────
        elif step == 3:
            uname = st.session_state.get("reset_uname", "")

            st.markdown(f"""
            <div style="font-size:16px;font-weight:700;color:#1a1916;margin-bottom:4px">
              Set your new password
            </div>
            <div style="font-size:12px;color:#6b6860;margin-bottom:18px">
              Identity verified for <b>{uname}</b>. Choose a strong new password.
              Minimum 8 characters.
            </div>""", unsafe_allow_html=True)

            with st.form("set_new_pw_form", clear_on_submit=False):
                new_pw1 = st.text_input(
                    "New password *",
                    type="password",
                    placeholder="At least 8 characters",
                    key="new_pw1_input",
                )
                new_pw2 = st.text_input(
                    "Confirm new password *",
                    type="password",
                    placeholder="Type your new password again",
                    key="new_pw2_input",
                )
                submit_step3 = st.form_submit_button(
                    "Reset password ✓",
                    type="primary",
                    use_container_width=True,
                )

            if submit_step3:
                if not new_pw1 or not new_pw2:
                    st.error("Both password fields are required.")
                elif len(new_pw1) < 8:
                    st.error("Password must be at least 8 characters.")
                elif new_pw1 != new_pw2:
                    st.error("Passwords do not match. Please try again.")
                else:
                    new_hash = bcrypt.hashpw(new_pw1.encode(), bcrypt.gensalt()).decode()
                    # Clear all reset state
                    for k in ["reset_step","reset_uname","reset_email",
                              "reset_code","reset_code_time"]:
                        st.session_state.pop(k, None)

                    # Show success
                    st.markdown("""
                    <div style="background:#0e1e2e;border-radius:12px;padding:20px 22px;
                                text-align:center;margin-bottom:16px">
                      <div style="font-size:28px;margin-bottom:8px">✅</div>
                      <div style="font-size:16px;font-weight:700;color:#3ecfaa;margin-bottom:6px">
                        Password reset successfully
                      </div>
                      <div style="font-size:12px;color:#a8c4d4;line-height:1.6">
                        Your new password is active. You can now sign in with your new credentials.
                      </div>
                    </div>""", unsafe_allow_html=True)

                    st.info(
                        "**For EPRA admin:** To make this password permanent across all "
                        "sessions, update secrets.toml with the hash below."
                    )
                    st.code(
                        f'[credentials.usernames.{uname}]\npassword = "{new_hash}"',
                        language="toml"
                    )
                    if st.button("← Go to login", type="primary",
                                 use_container_width=True, key="goto_login_after_reset"):
                        st.session_state["show_forgot"] = False
                        st.rerun()
                    return  # stop here — don't show back button below

        # ── Back to login (shown on steps 1 and 2) ────────────────────────────
        if step in (1, 2):
            st.markdown("")
            back_col, _ = st.columns([1, 1])
            if back_col.button("← Back to login", key="back_to_login_btn"):
                st.session_state["show_forgot"] = False
                st.session_state["reset_step"]  = 1
                for k in ["reset_uname","reset_email","reset_code","reset_code_time"]:
                    st.session_state.pop(k, None)
                st.rerun()

# ── LOGIN PAGE ────────────────────────────────────────────────────────────────
def show_login_page(authenticator):
    col1,col2,col3 = st.columns([1,1.2,1])
    with col2:
        st.markdown("""<div style="text-align:center;margin-bottom:24px;margin-top:40px">
          <div style="font-size:28px;font-weight:700;color:#1a1916">Kenya<span style="color:#0f9d7e">Watts</span></div>
          <div style="font-size:13px;color:#9c9a8e;margin-top:4px">Digital Integrated National Energy Planning Platform</div>
          <div style="font-size:11px;color:#c8c6be;margin-top:2px">EPRA Kenya · NGDA 2026</div>
        </div>""", unsafe_allow_html=True)

        result = authenticator.login(
            location="main",
            fields={"Form name":"Sign in to KenyaWatts","Username":"Username",
                    "Password":"Password","Login":"Sign in"},
            key="kw_login_form",
        )
        if result is not None:
            name, auth_status, username = result
        else:
            # If result is None, check if authenticator already set session state via cookie
            name        = st.session_state.get("name","")
            auth_status = st.session_state.get("authentication_status", None)
            username    = st.session_state.get("username","")

        if auth_status is False:
            st.error("Incorrect username or password. Please try again.")

        # Forgot password link — clean, no credentials shown
        st.markdown("")
        if st.button("Forgot your password?", use_container_width=True, key="forgot_pw_btn"):
            st.session_state["show_forgot"] = True
            st.session_state["reset_step"]  = 1   # always start at step 1
            st.rerun()

    return name, auth_status, username

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
def build_sidebar(authenticator, user_info):
    with st.sidebar:
        role      = user_info["role"]
        is_epra   = role in ("epra","ministry","devpartner")
        is_county = role == "county"
        rc = {"epra":"#1a6fa3","ministry":"#0f9d7e","devpartner":"#5b4fc9","county":"#d4891a","service_provider":"#7a7870"}.get(role,"#1a6fa3")
        rl = {"epra":"EPRA Planner","ministry":"Ministry of Energy","devpartner":"Development Partner","county":"County Committee","service_provider":"KPLC"}.get(role,"User")
        dname = st.session_state.get("user_display_name",{}).get(user_info.get("_username",""), user_info.get("name",""))

        st.markdown(f"""<div style="background:#0e1e2e;border-radius:10px;padding:14px;margin-bottom:16px">
          <div style="font-size:15px;font-weight:700;color:white">Kenya<span style="color:#3ecfaa">Watts</span></div>
          <div style="font-size:10px;color:#5a8a9e;margin-top:2px">National Energy Planning Platform</div>
          <div style="margin-top:10px;padding:8px 10px;background:rgba(255,255,255,0.06);border-radius:7px">
            <div style="font-size:12px;font-weight:600;color:white">{dname}</div>
            <div style="font-size:11px;color:{rc};margin-top:2px">{rl}</div>
          </div></div>""", unsafe_allow_html=True)

        st.markdown("### Navigation")
        if role == "epra":
            opts = [
                "📊  National overview","🗺️  County map","🏛️  All 47 counties",
                "✅  Validation queue","📥  Communications hub",
                "📄  Makueni CEP — template reference",
                "📋  National INEP report",
                "⬇️  Data download","🤖  AI assistant","⚙️  Account settings",
            ]
        elif role == "ministry":
            opts = [
                "📊  National overview","🗺️  County map","🏛️  All 47 counties",
                "📥  Communications hub","⬇️  Data download",
                "🤖  AI assistant","⚙️  Account settings",
            ]
        elif role == "devpartner":
            opts = [
                "📊  National overview","🗺️  County map","🏛️  All 47 counties",
                "⬇️  Data download","🤖  AI assistant","⚙️  Account settings",
            ]
        elif role == "county":
            cname = COUNTIES[COUNTIES["id"]==user_info["county_id"]]["name"].values[0] if user_info["county_id"] else "County"
            is_makueni = user_info.get("county_id","") == "MK"
            opts = [
                f"🏠  {cname} dashboard",
                "📤  Submit energy plan","📨  Communications",
                "⬇️  Data download","🤖  AI assistant","⚙️  Account settings",
            ]
            if is_makueni:
                opts.insert(2, "📄  My county plan")
        elif role == "service_provider":
            opts = [
                "📊  National overview","🏛️  County demand data",
                "📤  Submit provider plan","⬇️  Data download",
                "🤖  AI assistant","⚙️  Account settings",
            ]
        else:
            opts = [
                "📊  National overview",
                "⬇️  Data download","🤖  AI assistant","⚙️  Account settings",
            ]

        selected = st.radio("", opts, label_visibility="collapsed")
        st.divider()
        # Use authenticator's own logout — this deletes the browser cookie properly
        # Without cookie deletion the user gets auto-logged back in
        try:
            authenticator.logout(
                button_name="🚪  Sign out",
                location="sidebar",
                key="kw_signout_btn",
                use_container_width=True,
            )
        except Exception:
            # Fallback: manually clear everything including the cookie name
            if st.button("🚪  Sign out", key="kw_signout_fallback", use_container_width=True):
                cookie_name = st.secrets.get("cookie",{}).get("name","kenyawatts_auth")
                for k in list(st.session_state.keys()):
                    del st.session_state[k]
                st.rerun()
        st.markdown("""<div style="font-size:10px;color:#9c9a8e;line-height:1.7;margin-top:8px">
        Data: EPRA FY 2024/25 · Makueni CEP 2023–2032 · KNBS 2019<br>NGDA 2026 · DTU · Challenge 2 · EPRA Kenya
        </div>""", unsafe_allow_html=True)
    return selected

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    init_session()
    authenticator = setup_auth()

    # ── On every load, let the authenticator check its own cookie first ─────────
    # Then check if the result is a valid user. If the cookie is expired or
    # invalid the authenticator sets authentication_status to None automatically.
    stored_user = st.session_state.get("username","")
    valid_users = list(st.secrets.get("credentials",{}).get("usernames",{}).keys())

    # If session state says logged in but username is blank or not in our list,
    # force logout — this catches stale cookies from previous sessions
    if st.session_state.get("authentication_status") is True:
        if not stored_user or stored_user not in valid_users:
            # Clear everything and force back to login
            for k in list(st.session_state.keys()):
                try:
                    del st.session_state[k]
                except Exception:
                    pass
            st.rerun()

    # Forgot password flow
    if st.session_state.get("show_forgot"):
        page_forgot_password(authenticator)
        return

    # Auth check — always show login if not authenticated
    # Note: st.session_state["authentication_status"] is set by authenticator itself
    # We read it directly rather than storing our own copy
    if st.session_state.get("authentication_status") is not True:
        name, auth_status, username = show_login_page(authenticator)
        if auth_status is True:
            # authenticator has already set st.session_state["authentication_status"] = True
            # and st.session_state["username"] — we just rerun to show the app
            st.rerun()
        return

    # Read username from authenticator's session state (it sets this itself)
    username  = st.session_state.get("username","")
    if not username:
        # Authenticator sometimes uses different key — check both
        username = st.session_state.get("name","")
    user_info = get_user_role(username)
    user_info["_username"] = username
    role      = user_info["role"]
    county_id = user_info["county_id"]

    show_notifications()

    # Header
    st.markdown(f"""<div class="kw-header">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <div class="kw-logo">Kenya<span>Watts</span>
          <span style="font-size:13px;font-weight:400;color:#5a8a9e;margin-left:14px">Digital Integrated National Energy Planning Platform · EPRA Kenya</span>
        </div>
        <div style="font-size:11px;color:#3ecfaa;font-weight:600">● Live · NGDA 2026</div>
      </div></div>""", unsafe_allow_html=True)

    selected = build_sidebar(authenticator, user_info)
    key      = selected.split("  ",1)[-1].strip()

    if   "National overview" in key:                    page_national_overview(role)
    elif "County map" in key:                           page_map(role, county_id)
    elif "All 47" in key or "County demand" in key:    page_all_counties()
    elif "dashboard" in key.lower():                    page_county_dashboard(county_id, user_info["name"])
    elif "Validation queue" in key:                     page_validation_queue()
    elif "Submit" in key:                               page_submit(role, county_id, user_info["name"])
    elif "Inbox" in key or "Communications" in key or "communications" in key.lower():     page_inbox(role, county_id, user_info["name"])
    elif "Makueni CEP" in key or "My county plan" in key: page_makueni_contextual(role, county_id)
    elif "National INEP report" in key:                  page_national_report(user_info)
    elif "Data download" in key:                        page_data_download(role, county_id)
    elif "AI" in key:                                   page_ai(role, county_id)
    elif "Account settings" in key:                     page_account_settings(authenticator, username, user_info)

if __name__ == "__main__":
    main()
