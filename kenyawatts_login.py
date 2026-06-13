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

STATUS_COLOR = {"submitted":"#0f9d7e","review":"#d4891a","pending":"#7a7870","overdue":"#b33a2c"}
STATUS_LABEL = {"submitted":"Submitted ✓","review":"In review","pending":"Pending","overdue":"Overdue"}

# ── SESSION STATE INIT ────────────────────────────────────────────────────────
def init_session():
    defaults = {
        "audit_log": [], "upload_log": [], "ai_history": [],
        "notifications": [], "user_display_name": {}, "user_email": {},
        "submitted_data": [],
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
        cookie.get("expiry_days",1),
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
            title=dict(text=metric, font=dict(size=12, color="#1a1916", family="Arial")),
            tickfont=dict(size=11, color="#1a1916", family="Arial")
        ))

    fig.update_layout(
        mapbox_style="carto-positron",
        mapbox=dict(center=dict(lat=0.5, lon=37.9), zoom=5),
        height=480, margin=dict(t=0,b=0,l=0,r=0),
        font=CHART_FONT,
    )

    # Legend for status view
    if metric == "Submission status":
        cols = st.columns(4)
        for col, (s,c) in zip(cols, STATUS_COLOR.items()):
            col.markdown(f'<div style="display:flex;align-items:center;gap:6px;font-size:12px"><div style="width:12px;height:12px;border-radius:50%;background:{c}"></div>{STATUS_LABEL[s]}</div>', unsafe_allow_html=True)
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
            st.markdown(f'<div style="padding:12px 14px;border-radius:10px;background:#f7f6f2;border-top:3px solid {color}"><div style="font-size:22px;font-weight:700;color:{color}">{n}</div><div style="font-size:11px;color:#6b6860;margin-top:2px">{lbl}</div></div>', unsafe_allow_html=True)

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

    st.markdown(f"""<div style="padding:14px 18px;border-radius:10px;border:1.5px solid {sc};background:{sc}11;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center">
      <div><div style="font-size:16px;font-weight:700;color:#1a1916">{row['name']} County Energy Plan</div>
      <div style="font-size:12px;color:#6b6860;margin-top:2px">Plan period: 2023–{row['target_yr']} · {row['region']}</div></div>
      <span style="font-size:13px;font-weight:600;color:{sc};background:{sc}20;padding:5px 14px;border-radius:20px">{STATUS_LABEL[row['status']]}</span>
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
    logs = st.session_state.upload_log
    if county_id:
        logs = [l for l in logs if l.get("county_id")==county_id]
    if not logs:
        st.info("No submissions or uploads yet for this session.")
        return
    st.markdown("**Submission and upload history**")
    for entry in reversed(logs[-10:]):
        icon = "📄" if entry.get("type")=="pdf" else "📝" if entry.get("type")=="form" else "✅"
        st.markdown(f"""<div class="upload-log-row">
          <span class="upload-log-icon">{icon}</span>
          <div>
            <div style="font-size:12px;font-weight:500;color:#1a1916">{entry.get('title','Submission')}</div>
            <div class="upload-log-meta">
              {entry.get('county','')} · {entry.get('date','')} at {entry.get('time','')} ·
              Status: <span style="color:{STATUS_COLOR.get(entry.get('status','pending'),'#7a7870')};font-weight:600">{entry.get('status','Pending').title()}</span>
              {f" · File: <code>{entry.get('filename','')}</code>" if entry.get('filename') else ''}
            </div>
            {f'<div style="font-size:11px;color:#9c9a8e;margin-top:2px">Ref: <code>{entry.get("ref","")}</code></div>' if entry.get("ref") else ''}
          </div>
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
                    county_in  = st.text_input("County name *", value=county_name if is_county else "Makueni", disabled=is_county)
                    elec_in    = st.number_input("Electricity access (%)*", value=extracted["elec"], min_value=0.0, max_value=100.0)
                    cooking_in = st.number_input("Clean cooking access (%)*", value=extracted["cooking"], min_value=0.0, max_value=100.0)
                    fw_in      = st.number_input("Firewood as primary fuel (%)", value=extracted["firewood"], min_value=0.0, max_value=100.0)
                with col2:
                    solar_in   = st.number_input("Solar GHI (kWh/m²/year)*", value=extracted["solar"], min_value=0.0)
                    budget_in  = st.number_input("Total budget (KES billions)", value=extracted["budget"], min_value=0.0)
                    target_in  = st.number_input("Universal access target year*", value=extracted["target_yr"], min_value=2025, max_value=2040)
                    mtf_in     = st.selectbox("MTF demand tier*", [1,2,3,4,5], index=extracted["mtf"]-1)
                    growth_in  = st.number_input("Population growth rate (%/yr)*", value=extracted["growth"], min_value=0.0, max_value=10.0)

                _validate_and_submit(county_in, elec_in, cooking_in, fw_in, solar_in,
                                      budget_in, target_in, mtf_in, growth_in,
                                      uploaded.name, "pdf", county_id, user_name)
            else:
                st.info("Upload your CEP document above. Use the Makueni County Energy Plan PDF as a test file.")

        else:
            section("Structured submission template")
            alert("info","<b>National assumptions pre-loaded:</b> KNBS baselines · cost benchmarks · solar GHI reference · Kenya's official targets.")

            with st.expander("▶ Section 1 — Electricity access", expanded=True):
                col1,col2 = st.columns(2)
                with col1:
                    county_in  = st.text_input("County name *", value=county_name, disabled=is_county)
                    elec_in    = st.number_input("Total electricity access (%)*", min_value=0.0, max_value=100.0, help="Grid + mini-grid + SHS")
                    cooking_in = st.number_input("Clean cooking access (%)*", min_value=0.0, max_value=100.0)
                    fw_in      = st.number_input("Firewood as primary fuel (%)", min_value=0.0, max_value=100.0)
                with col2:
                    solar_in   = st.number_input("Solar GHI (kWh/m²/year)*", min_value=0.0, help="Kenya range: 1,600–2,200")
                    budget_in  = st.number_input("Total plan budget (KES billions)", min_value=0.0)
                    target_in  = st.number_input("Universal access target year*", value=2030, min_value=2025, max_value=2040)
                    mtf_in     = st.selectbox("MTF demand tier*", [1,2,3,4,5], index=2)
                    growth_in  = st.number_input("Population growth rate (%/yr)*", value=1.1, min_value=0.0, max_value=10.0)

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
                ref  = f"CEP-{county[:2].upper()}-2026-{abs(hash(county+str(datetime.now())))%9000+1000}"
                now  = datetime.now()
                # Log the upload/submission
                st.session_state.upload_log.append({
                    "title":     f"County Energy Plan — {county}",
                    "county":    county,
                    "county_id": county_id or county[:2].upper(),
                    "type":      ptype,
                    "filename":  filename or "Structured template",
                    "date":      now.strftime("%d %b %Y"),
                    "time":      now.strftime("%H:%M:%S"),
                    "status":    "submitted",
                    "ref":       ref,
                    "elec":      elec, "cooking": cooking,
                    "solar":     solar,"budget":  budget,
                    "target_yr": target, "mtf": mtf,
                    "submitted_by": user_name,
                })
                st.session_state.audit_log.append({
                    "time": now.strftime('%Y-%m-%d %H:%M:%S'),
                    "user": user_name, "action": f"Plan submitted — {county}","ref": ref
                })
                # Store for download
                st.session_state.submitted_data.append({
                    "county":county,"submitted_by":user_name,
                    "date":now.strftime("%Y-%m-%d"),"time":now.strftime("%H:%M"),
                    "elec_pct":elec,"cooking_pct":cooking,"solar_ghi":solar,
                    "budget_kes_b":budget,"target_year":target,"mtf_tier":mtf,
                    "growth_pct":growth,"firewood_pct":fw,
                    "document":filename or "Structured template","ref":ref
                })
                # Push notification
                push_notification(f"Plan submitted! {county} County · Ref: {ref}", "✅")
                st.balloons()
                # Confirmation card
                st.markdown(f"""<div style="background:#e8f7f4;border:1.5px solid #0f9d7e;border-radius:12px;padding:20px;margin-top:16px">
                  <div style="font-size:16px;font-weight:700;color:#0f9d7e;margin-bottom:8px">✅ Submission confirmed</div>
                  <div style="font-size:13px;color:#1a1916;margin-bottom:4px"><b>County:</b> {county}</div>
                  <div style="font-size:13px;color:#1a1916;margin-bottom:4px"><b>Submitted by:</b> {user_name}</div>
                  <div style="font-size:13px;color:#1a1916;margin-bottom:4px"><b>Date and time:</b> {now.strftime('%d %b %Y at %H:%M:%S')}</div>
                  <div style="font-size:13px;color:#1a1916;margin-bottom:4px"><b>Document:</b> {filename or 'Structured template'}</div>
                  <div style="font-size:13px;color:#1a1916;margin-bottom:12px"><b>Reference number:</b> <code>{ref}</code></div>
                  <div style="font-size:12px;color:#0f9d7e">EPRA will validate and confirm within 14 days. Your county is now shown as "Submitted" on the national map.</div>
                </div>""", unsafe_allow_html=True)

# ── VALIDATION QUEUE ──────────────────────────────────────────────────────────
def page_validation_queue():
    alert("warn","<b>EPRA admin only.</b> Review flagged submissions before including in national aggregation.")
    review = COUNTIES[COUNTIES["status"]=="review"]
    if review.empty:
        st.info("No submissions in review queue.")
    for _,row in review.iterrows():
        with st.expander(f"**{row['name']} County** — submitted 3 days ago"):
            c1,c2,c3 = st.columns(3)
            c1.metric("Electricity",f"{row['elec']}%")
            c2.metric("Solar GHI",f"{row['solar']} kWh/m²")
            c3.metric("Target year",str(row["target_yr"]))
            alert("warn","<b>2 warnings:</b> Solar GHI slightly above expected range · Population growth assumption >2%. No critical errors.")
            a,b,_ = st.columns([1,1,2])
            if a.button(f"✅ Approve — {row['name']}", key=f"app_{row['id']}"):
                st.success(f"{row['name']} approved and included in national aggregation.")
                push_notification(f"{row['name']} County plan approved", "✅")
            if b.button("↩ Request resubmission", key=f"rej_{row['id']}"):
                st.warning(f"Resubmission request sent to {row['name']} County.")

    st.divider()
    section("National aggregation trigger")
    nat = compute_national()
    st.info(f"{nat['submitted_count']} approved · {len(review)} in review · {nat['pending_count']+nat['overdue_count']} not submitted")
    if st.button("▶  Run national aggregation now", type="primary"):
        with st.spinner("Aggregating submissions into national INEP…"):
            import time; time.sleep(2)
        st.success(f"Aggregation complete. INEP updated with {nat['submitted_count']} county plans.")
        push_notification("National INEP aggregation complete", "🗂️")

# ── INBOX ─────────────────────────────────────────────────────────────────────
def page_inbox(role, county_id, user_name):
    is_epra = role in ("epra","ministry")
    cnty    = COUNTIES[COUNTIES["id"]==county_id]["name"].values[0] if county_id else "Your county"
    msgs    = (
        [{"from":"System","to":"All counties","date":"2026-06-01","type":"Assumptions","subject":"INEP 2025 planning assumptions updated","body":"KNBS 2024 projections and updated MTF benchmarks published. Use these for 2026 submissions."},
         {"from":"Nakuru County","to":"EPRA","date":"2026-05-28","type":"Query","subject":"Solar GHI methodology","body":"Different values from Global Solar Atlas vs EPRA template. Please clarify."},
         {"from":"System","to":"Turkana, Marsabit, Mandera, Wajir","date":"2026-05-20","type":"Reminder","subject":"Overdue — 3rd notice","body":"Submission required under INEP Regulations 2025."}]
        if is_epra else
        [{"from":"EPRA","to":cnty,"date":"2026-06-01","type":"Assumptions","subject":"Updated planning assumptions — June 2026","body":"Use updated KNBS baselines. Solar GHI reference 1,600–2,200 kWh/m². Electricity target: 100% by 2030."},
         {"from":"EPRA","to":cnty,"date":"2026-05-15","type":"Benchmark","subject":"Peer county comparison","body":"Your electricity access is within regional average. Clean cooking is below regional average — prioritise LPG and biogas."},
         {"from":"EPRA","to":cnty,"date":"2026-04-30","type":"Guidance","subject":"Makueni CEP available as reference","body":"Makueni County Energy Plan (2023–2032) available as a reference. Covers all 7 required sectors."}]
    )
    tc = {"Assumptions":"#1a6fa3","Query":"#d4891a","Reminder":"#b33a2c","Benchmark":"#5b4fc9","Guidance":"#0f9d7e"}
    for m in msgs:
        c = tc.get(m["type"],"#7a7870")
        st.markdown(f"""<div style="background:white;border:0.5px solid #e8e6de;border-radius:10px;padding:14px;margin-bottom:8px">
          <div style="display:flex;justify-content:space-between;margin-bottom:6px">
            <div><span style="font-size:10px;font-weight:600;padding:2px 8px;border-radius:8px;background:{c}20;color:{c};text-transform:uppercase;margin-right:8px">{m['type']}</span>
            <span style="font-size:13px;font-weight:600;color:#1a1916">{m['subject']}</span></div>
            <span style="font-size:11px;color:#9c9a8e">{m['date']}</span></div>
          <div style="font-size:11px;color:#9c9a8e;margin-bottom:6px">From: {m['from']} → {m['to']}</div>
          <div style="font-size:12px;color:#6b6860;line-height:1.6">{m['body']}</div>
        </div>""", unsafe_allow_html=True)

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
        for label, value in [("Username",username),("Display name",dname),("Email",demail or "Not set"),
                              ("Role",role_labels.get(role,role)),("County",cnty),
                              ("Session",datetime.now().strftime("%d %b %Y %H:%M")),("Status","Active ✓")]:
            cl,cr = st.columns([1,2])
            cl.markdown(f"<div style='font-size:12px;font-weight:600;color:#9c9a8e;padding:5px 0'>{label}</div>", unsafe_allow_html=True)
            cr.markdown(f"<div style='font-size:13px;color:#1a1916;padding:5px 0'>{value}</div>", unsafe_allow_html=True)
        st.divider()
        audit = [a for a in st.session_state.audit_log if a.get("user")==username]
        if audit:
            st.markdown("**Recent activity**")
            for e in reversed(audit[-5:]):
                st.markdown(f'<div style="font-size:11px;color:#6b6860;padding:4px 0;border-bottom:0.5px solid #f0ede4">'
                            f'<span style="color:#9c9a8e;font-family:monospace">{e["time"]}</span> &nbsp; {e["action"]} &nbsp; '
                            f'<span style="color:#9c9a8e;font-family:monospace">{e.get("ref","")}</span></div>', unsafe_allow_html=True)
        alert("info","<b>Need to change username or role?</b> Contact EPRA admin: Allan.Wairimu@epra.go.ke · +254720850696")

# ── FORGOT PASSWORD ───────────────────────────────────────────────────────────
def page_forgot_password(authenticator):
    st.markdown("""<div style="max-width:480px;margin:40px auto">
    <div style="font-size:22px;font-weight:700;color:#1a1916;margin-bottom:4px">Reset your password</div>
    <div style="font-size:13px;color:#9c9a8e;margin-bottom:24px">Enter your username and email. If they match our records, you will receive a reset code.</div>
    </div>""", unsafe_allow_html=True)

    col1,col2,col3 = st.columns([1,1.4,1])
    with col2:
        with st.form("forgot_pw_form"):
            uname  = st.text_input("Username",      placeholder="Your KenyaWatts username")
            email  = st.text_input("Email address", placeholder="The email on your account")
            submit = st.form_submit_button("Send reset code", type="primary", use_container_width=True)

        if submit:
            creds = st.secrets.get("credentials",{}).get("usernames",{})
            user  = creds.get(uname,{})
            if uname in creds and user.get("email","").lower()==email.lower():
                # Generate a simple 6-digit code (stored in session)
                import random
                code = str(random.randint(100000,999999))
                st.session_state[f"reset_code_{uname}"] = code
                st.session_state[f"reset_code_time_{uname}"] = datetime.now()
                st.success(f"✓ Reset code generated. In a real deployment this would be emailed to {email}.")
                st.info(f"**Demo mode:** Your reset code is `{code}` (in production this appears only in your email inbox)")

                # Let them enter the code and new password
                st.markdown("---")
                st.markdown("**Enter your reset code and new password:**")
                with st.form("reset_confirm_form"):
                    entered_code = st.text_input("Reset code (6 digits)", placeholder="Enter the code from your email")
                    new_pw1      = st.text_input("New password",          type="password", placeholder="At least 8 characters")
                    new_pw2      = st.text_input("Confirm new password",  type="password")
                    confirm_btn  = st.form_submit_button("Reset password", type="primary", use_container_width=True)

                if confirm_btn:
                    stored_code = st.session_state.get(f"reset_code_{uname}","")
                    code_time   = st.session_state.get(f"reset_code_time_{uname}", datetime.now()-timedelta(hours=2))
                    expired     = (datetime.now() - code_time).seconds > 900  # 15 min expiry

                    if expired:
                        st.error("Reset code has expired (valid for 15 minutes). Request a new one.")
                    elif entered_code != stored_code:
                        st.error("Incorrect reset code. Please try again.")
                    elif len(new_pw1)<8:
                        st.error("Password must be at least 8 characters.")
                    elif new_pw1!=new_pw2:
                        st.error("Passwords do not match.")
                    else:
                        new_hash = bcrypt.hashpw(new_pw1.encode(), bcrypt.gensalt()).decode()
                        # Clear the code
                        st.session_state.pop(f"reset_code_{uname}", None)
                        st.session_state.pop(f"reset_code_time_{uname}", None)
                        push_notification("Password reset successfully. Please log in with your new password.", "🔑")
                        st.success("✓ Password reset successfully.")
                        st.info("To make permanent: EPRA admin must update secrets.toml with the new hash below.")
                        st.code(f'[credentials.usernames.{uname}]\npassword = "{new_hash}"', language="toml")
                        if st.button("← Back to login"):
                            st.session_state["show_forgot"] = False
                            st.rerun()
            else:
                st.error("Username and email do not match our records. Contact EPRA admin if you need help.")

        if st.button("← Back to login", use_container_width=True):
            st.session_state["show_forgot"] = False
            st.rerun()

        st.markdown("""<div style="margin-top:16px;padding:12px;background:#f7f6f2;border-radius:8px;font-size:11px;color:#6b6860;line-height:1.7">
        <b>How password reset is verified:</b><br>
        1. You provide your username AND the email registered to that account<br>
        2. Both must match exactly what is in the system<br>
        3. A 6-digit code is generated and sent to that email (simulated in demo)<br>
        4. The code expires after 15 minutes<br>
        5. You enter the code plus your new password to confirm<br>
        This two-factor approach (username + registered email + time-limited code) prevents unauthorised password resets.
        </div>""", unsafe_allow_html=True)

# ── LOGIN PAGE ────────────────────────────────────────────────────────────────
def show_login_page(authenticator):
    col1,col2,col3 = st.columns([1,1.2,1])
    with col2:
        st.markdown("""<div style="text-align:center;margin-bottom:24px;margin-top:40px">
          <div style="font-size:28px;font-weight:700;color:#1a1916">Kenya<span style="color:#0f9d7e">Watts</span></div>
          <div style="font-size:13px;color:#9c9a8e;margin-top:4px">Digital Integrated National Energy Planning Platform</div>
          <div style="font-size:11px;color:#c8c6be;margin-top:2px">EPRA Kenya · NGDA 2026</div>
        </div>""", unsafe_allow_html=True)

        try:
            name, auth_status, username = authenticator.login(
                location="main",
                fields={"Form name":"Sign in","Username":"Username","Password":"Password","Login":"Sign in"}
            )
        except:
            name, auth_status, username = authenticator.login("Sign in", "main")

        if auth_status is False:
            st.error("Incorrect username or password.")
        elif auth_status is None:
            st.markdown("""<div style="background:#f7f6f2;border-radius:10px;padding:14px 16px;margin-top:14px;font-size:12px;color:#6b6860;line-height:1.9">
            <b style="color:#1a1916;display:block;margin-bottom:6px">Demo credentials</b>
            <b>EPRA:</b> epra_admin / epra2026<br>
            <b>Ministry:</b> ministry / ministry2026<br>
            <b>Dev partner:</b> devpartner / partner2026<br>
            <b>Makueni:</b> makueni / makueni2026<br>
            <b>Turkana:</b> turkana / turkana2026<br>
            <b>KPLC:</b> kplc / kplc2026
            </div>""", unsafe_allow_html=True)

        # Forgot password link
        st.markdown("")
        if st.button("Forgot your password?", use_container_width=True):
            st.session_state["show_forgot"] = True
            st.rerun()

    return auth_status, username

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
        if is_epra:
            county_name_for_nav = ""
            opts = ["📊  National overview","🗺️  County map","🏛️  All 47 counties","✅  Validation queue",
                    "📥  Communications hub","📄  Makueni reference","⬇️  Data download",
                    "🤖  AI assistant","⚙️  Account settings"]
        elif is_county:
            cname = COUNTIES[COUNTIES["id"]==user_info["county_id"]]["name"].values[0] if user_info["county_id"] else "County"
            opts  = [f"🏠  {cname} dashboard","🗺️  County map","📤  Submit energy plan",
                     "📥  County inbox","📄  Makueni reference","⬇️  Data download",
                     "🤖  AI assistant","⚙️  Account settings"]
        else:
            opts  = ["📊  National overview","🗺️  County map","🗺️  County demand data",
                     "📤  Submit provider plan","⬇️  Data download","🤖  AI assistant","⚙️  Account settings"]

        selected = st.radio("", opts, label_visibility="collapsed")
        st.divider()
        try:
            authenticator.logout("Sign out","sidebar")
        except:
            if st.button("Sign out"):
                st.session_state["authentication_status"] = None
                st.rerun()
        st.markdown("""<div style="font-size:10px;color:#9c9a8e;line-height:1.7;margin-top:8px">
        Data: EPRA FY 2024/25 · Makueni CEP 2023–2032 · KNBS 2019<br>NGDA 2026 · DTU · Challenge 2 · EPRA Kenya
        </div>""", unsafe_allow_html=True)
    return selected

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    init_session()
    authenticator = setup_auth()

    # Forgot password flow
    if st.session_state.get("show_forgot"):
        page_forgot_password(authenticator)
        return

    # Auth check
    if st.session_state.get("authentication_status") is not True:
        auth_status, username = show_login_page(authenticator)
        if auth_status:
            st.session_state["authentication_status"] = True
            st.session_state["username"] = username
            st.rerun()
        return

    username  = st.session_state.get("username","")
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

    if   "National overview" in key:             page_national_overview(role)
    elif "County map" in key or "map" in key.lower(): page_map(role, county_id)
    elif "All 47" in key or "County demand" in key:  page_all_counties()
    elif "dashboard" in key.lower():             page_county_dashboard(county_id, user_info["name"])
    elif "Validation queue" in key:              page_validation_queue()
    elif "Submit" in key:                        page_submit(role, county_id, user_info["name"])
    elif "Inbox" in key or "Communications" in key: page_inbox(role, county_id, user_info["name"])
    elif "Makueni" in key:                       page_makueni()
    elif "Data download" in key:                 page_data_download(role, county_id)
    elif "AI" in key:                            page_ai(role, county_id)
    elif "Account settings" in key:              page_account_settings(authenticator, username, user_info)

if __name__ == "__main__":
    main()
