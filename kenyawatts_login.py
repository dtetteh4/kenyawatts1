"""
KenyaWatts — Digital Integrated National Energy Planning Platform
NGDA 2026 · DTU Young Academics Track · Challenge 2: EPRA Kenya

Run locally:  streamlit run kenyawatts.py
Deploy:       Push to GitHub → Streamlit Community Cloud → one URL for all users
"""

import streamlit as st
import streamlit_authenticator as stauth
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import anthropic
from datetime import datetime

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="KenyaWatts · National Energy Planning Platform",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  #MainMenu{visibility:hidden} footer{visibility:hidden} .stDeployButton{display:none}
  .kw-header{background:#0e1e2e;color:white;padding:14px 20px;border-radius:10px;margin-bottom:16px}
  .kw-logo{font-size:22px;font-weight:700;letter-spacing:-0.3px}
  .kw-logo span{color:#3ecfaa}
  .kw-role-banner{padding:8px 16px;border-radius:8px;font-size:12px;margin-bottom:14px;font-weight:500}
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
  .kw-login-box{max-width:420px;margin:60px auto;padding:32px;background:white;border-radius:16px;box-shadow:0 4px 24px rgba(0,0,0,0.08)}
  .kw-login-title{font-size:22px;font-weight:700;color:#1a1916;margin-bottom:4px}
  .kw-login-sub{font-size:13px;color:#9c9a8e;margin-bottom:24px}
  .kw-audit-row{font-size:11px;color:#6b6860;padding:4px 0;border-bottom:0.5px solid #f0ede4;display:flex;gap:12px}
  div[data-testid="stForm"]{border:none!important;padding:0!important}
</style>
""", unsafe_allow_html=True)

# ── COUNTY DATA ───────────────────────────────────────────────────────────────
COUNTIES = pd.DataFrame([
    {"id":"NK","name":"Nairobi",    "region":"Nairobi",    "pop":4922000,"elec":96, "cooking":62,"solar":1980,"budget":120.0,"status":"submitted","mtf":4,"growth":1.8,"target_yr":2027},
    {"id":"MB","name":"Mombasa",    "region":"Coast",      "pop":1208000,"elec":84, "cooking":45,"solar":2050,"budget":55.0, "status":"submitted","mtf":4,"growth":1.5,"target_yr":2028},
    {"id":"MK","name":"Makueni",    "region":"South East", "pop":987653, "elec":75, "cooking":18,"solar":2008,"budget":74.9, "status":"submitted","mtf":3,"growth":1.1,"target_yr":2028},
    {"id":"NA","name":"Nakuru",     "region":"Rift Valley","pop":2162000,"elec":72, "cooking":38,"solar":1920,"budget":68.0, "status":"review",   "mtf":3,"growth":1.4,"target_yr":2029},
    {"id":"KI","name":"Kisumu",     "region":"Nyanza",     "pop":1155000,"elec":67, "cooking":31,"solar":1870,"budget":0,    "status":"pending",  "mtf":3,"growth":1.2,"target_yr":2030},
    {"id":"KW","name":"Kwale",      "region":"Coast",      "pop":866000, "elec":41, "cooking":18,"solar":2020,"budget":0,    "status":"pending",  "mtf":2,"growth":1.3,"target_yr":2031},
    {"id":"KF","name":"Kilifi",     "region":"Coast",      "pop":1453000,"elec":38, "cooking":16,"solar":2010,"budget":0,    "status":"pending",  "mtf":2,"growth":1.4,"target_yr":2031},
    {"id":"GR","name":"Garissa",    "region":"North East", "pop":841000, "elec":22, "cooking":8, "solar":2140,"budget":0,    "status":"pending",  "mtf":1,"growth":2.1,"target_yr":2032},
    {"id":"KA","name":"Kajiado",    "region":"South Rift", "pop":1117000,"elec":55, "cooking":29,"solar":1990,"budget":0,    "status":"pending",  "mtf":2,"growth":1.8,"target_yr":2030},
    {"id":"MR","name":"Muranga",    "region":"Central",    "pop":1056000,"elec":68, "cooking":33,"solar":1900,"budget":0,    "status":"pending",  "mtf":3,"growth":0.9,"target_yr":2029},
    {"id":"TK","name":"Turkana",    "region":"North Rift", "pop":926000, "elec":12, "cooking":3, "solar":2150,"budget":0,    "status":"overdue",  "mtf":1,"growth":2.8,"target_yr":2033},
    {"id":"MS","name":"Marsabit",   "region":"North East", "pop":459000, "elec":8,  "cooking":2, "solar":2180,"budget":0,    "status":"overdue",  "mtf":1,"growth":2.5,"target_yr":2034},
    {"id":"MN","name":"Mandera",    "region":"North East", "pop":1025000,"elec":11, "cooking":4, "solar":2200,"budget":0,    "status":"overdue",  "mtf":1,"growth":3.1,"target_yr":2034},
    {"id":"WJ","name":"Wajir",      "region":"North East", "pop":781000, "elec":9,  "cooking":3, "solar":2190,"budget":0,    "status":"overdue",  "mtf":1,"growth":2.9,"target_yr":2034},
])

GEN_TREND = pd.DataFrame({
    "Year":["FY19/20","FY20/21","FY21/22","FY22/23","FY23/24","FY24/25"],
    "GWh": [11564,    11891,    12210,    12897,    13685,    14520]
})
GEN_MIX = pd.DataFrame({
    "Source": ["Geothermal","Hydro","Wind","Solar","Thermal","Other"],
    "Pct":    [25.9, 23.8, 17.4, 14.9, 8.7, 9.3],
    "Color":  ["#0f9d7e","#1a6fa3","#5b4fc9","#d4891a","#b33a2c","#7a7870"]
})
MAKUENI_COOKING = pd.DataFrame({
    "Fuel":["Firewood","LPG","Charcoal","Biogas","Electric","Other"],
    "Pct": [72.5,      17.6, 8.2,      0.2,     0.3,       1.2]
})
OUTAGE = pd.DataFrame({
    "Month":["Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun"],
    "Hours":[7.2,  9.1,  11.1, 8.4,  7.8,  6.9,  8.2,  9.4,  8.8,  7.1,  6.5,  7.3]
})

# ── AGGREGATION ENGINE ────────────────────────────────────────────────────────
def compute_national():
    submitted = COUNTIES[COUNTIES["status"].isin(["submitted","review"])]
    if submitted.empty:
        return {}
    total_pop = submitted["pop"].sum()
    w_elec    = (submitted["elec"]    * submitted["pop"]).sum() / total_pop
    w_cooking = (submitted["cooking"] * submitted["pop"]).sum() / total_pop
    total_budget = submitted["budget"].sum()
    latest_target = submitted["target_yr"].max()
    return {
        "submitted_count":  len(submitted),
        "total_counties":   len(COUNTIES),
        "w_elec":           round(w_elec, 1),
        "w_cooking":        round(w_cooking, 1),
        "total_budget":     round(total_budget, 1),
        "latest_target":    int(latest_target),
        "coverage_pct":     round(len(submitted)/len(COUNTIES)*100),
        "overdue_count":    len(COUNTIES[COUNTIES["status"]=="overdue"]),
        "pending_count":    len(COUNTIES[COUNTIES["status"]=="pending"]),
    }

# ── AI ASSISTANT ──────────────────────────────────────────────────────────────
def ask_ai(question: str, history: list) -> str:
    try:
        api_key = st.secrets.get("anthropic", {}).get("api_key", "")
        if not api_key or api_key == "YOUR_ANTHROPIC_API_KEY_HERE":
            return "AI assistant is not configured. Add your Anthropic API key to .streamlit/secrets.toml under [anthropic] api_key."
        client = anthropic.Anthropic(api_key=api_key)
        nat = compute_national()
        system = f"""You are the KenyaWatts AI assistant for EPRA's national energy planning platform.
Real data: {nat.get('submitted_count',0)} counties submitted plans out of {nat.get('total_counties',14)}.
Weighted electricity access from submitted plans: {nat.get('w_elec','N/A')}% (national target: 100% by 2030).
Weighted clean cooking from submitted plans: {nat.get('w_cooking','N/A')}% (national target: 100% by 2028).
Clean energy generation: 82% (target 100% by 2035).
Overdue counties: {nat.get('overdue_count',0)} — Turkana (12%), Marsabit (8%), Mandera (11%), Wajir (9%).
Makueni CEP: electricity 75.1%, solar GHI 2,008 kWh/m², budget KES 74.9B, firewood 72.5% primary cooking fuel.
Generation: 13,685 GWh FY24/25, mix: Geothermal 25.9%, Hydro 23.8%, Wind 17.4%, Solar 14.9%.
Average outage: 8.8 hrs/month vs EPRA benchmark 5.0 hrs. 11 of 12 months exceeded benchmark.
INEP Regulations 2025 legally require all 47 county submissions.
Answer in 3-4 sentences. Plain text only, no markdown formatting."""
        messages = [{"role": m["role"], "content": m["content"]} for m in history]
        messages.append({"role": "user", "content": question})
        response = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=600,
            system=system, messages=messages
        )
        return response.content[0].text
    except Exception as e:
        return f"AI error: {str(e)}"

# ── LOGIN SETUP ───────────────────────────────────────────────────────────────
def setup_auth():
    creds = st.secrets.get("credentials", {})
    usernames_dict = {}
    for username, info in creds.get("usernames", {}).items():
        usernames_dict[username] = {
            "email":    info.get("email", ""),
            "name":     info.get("name", username),
            "password": info.get("password", ""),
        }
    cookie_cfg = st.secrets.get("cookie", {})
    authenticator = stauth.Authenticate(
        {"usernames": usernames_dict},
        cookie_cfg.get("name",    "kenyawatts_auth"),
        cookie_cfg.get("key",     "kenyawatts_key"),
        cookie_cfg.get("expiry_days", 1),
    )
    return authenticator

def get_user_role(username: str) -> dict:
    creds = st.secrets.get("credentials", {}).get("usernames", {})
    user_info = creds.get(username, {})
    return {
        "role":       user_info.get("role", "county"),
        "county_id":  user_info.get("county_id", ""),
        "name":       user_info.get("name", username),
        "email":      user_info.get("email", ""),
    }

# ── LOGIN PAGE ────────────────────────────────────────────────────────────────
def show_login_page(authenticator):
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div style="text-align:center;margin-bottom:24px;margin-top:40px">
          <div style="font-size:28px;font-weight:700;color:#1a1916">
            Kenya<span style="color:#0f9d7e">Watts</span>
          </div>
          <div style="font-size:13px;color:#9c9a8e;margin-top:4px">
            Digital Integrated National Energy Planning Platform
          </div>
          <div style="font-size:11px;color:#c8c6be;margin-top:2px">
            EPRA Kenya · NGDA 2026
          </div>
        </div>
        """, unsafe_allow_html=True)

        try:
            name, auth_status, username = authenticator.login(
                location="main",
                fields={"Form name":"Sign in to KenyaWatts","Username":"Username","Password":"Password","Login":"Sign in"}
            )
        except Exception:
            name, auth_status, username = authenticator.login("Sign in to KenyaWatts", "main")

        if auth_status is False:
            st.error("Incorrect username or password. Please try again.")
        elif auth_status is None:
            st.markdown("""
            <div style="background:#f7f6f2;border-radius:10px;padding:14px 16px;margin-top:16px;font-size:12px;color:#6b6860;line-height:1.8">
            <strong style="color:#1a1916;display:block;margin-bottom:6px">Demo credentials</strong>
            <table style="width:100%;border-collapse:collapse">
              <tr><td style="padding:2px 0"><b>EPRA planner:</b></td><td style="padding:2px 0">epra_admin / epra2026</td></tr>
              <tr><td style="padding:2px 0"><b>Ministry:</b></td><td style="padding:2px 0">ministry / ministry2026</td></tr>
              <tr><td style="padding:2px 0"><b>Dev. partner:</b></td><td style="padding:2px 0">devpartner / partner2026</td></tr>
              <tr><td style="padding:2px 0"><b>Makueni county:</b></td><td style="padding:2px 0">makueni / makueni2026</td></tr>
              <tr><td style="padding:2px 0"><b>Turkana county:</b></td><td style="padding:2px 0">turkana / turkana2026</td></tr>
              <tr><td style="padding:2px 0"><b>KPLC:</b></td><td style="padding:2px 0">kplc / kplc2026</td></tr>
            </table>
            </div>
            """, unsafe_allow_html=True)
    return auth_status, username

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
def build_sidebar(authenticator, user_info):
    with st.sidebar:
        role      = user_info["role"]
        is_epra   = role in ("epra","ministry","devpartner")
        is_county = role == "county"

        # User card
        role_labels = {"epra":"EPRA Planner","ministry":"Ministry of Energy",
                       "devpartner":"Development Partner","county":"County Committee",
                       "service_provider":"National Service Provider"}
        role_colors = {"epra":"#1a6fa3","ministry":"#0f9d7e","devpartner":"#5b4fc9",
                       "county":"#d4891a","service_provider":"#7a7870"}
        rc = role_colors.get(role,"#1a6fa3")

        st.markdown(f"""
        <div style="background:#0e1e2e;border-radius:10px;padding:14px;margin-bottom:16px">
          <div style="font-size:15px;font-weight:700;color:white">Kenya<span style="color:#3ecfaa">Watts</span></div>
          <div style="font-size:10px;color:#5a8a9e;margin-top:2px">National Energy Planning Platform</div>
          <div style="margin-top:12px;padding:8px 10px;background:rgba(255,255,255,0.06);border-radius:7px">
            <div style="font-size:12px;font-weight:600;color:white">{user_info['name']}</div>
            <div style="font-size:11px;color:{rc};margin-top:2px">{role_labels.get(role,'User')}</div>
            {f'<div style="font-size:10px;color:#5a8a9e;margin-top:1px">{COUNTIES[COUNTIES.id==user_info["county_id"]].name.values[0] if user_info["county_id"] else ""} County</div>' if is_county else ''}
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### Navigation")

        # Build nav options based on role
        nav_options = []
        if is_epra:
            nav_options = [
                "📊  National overview",
                "🗺️  All 47 counties",
                "✅  Validation queue",
                "📥  Communications hub",
                "📄  Makueni reference",
                "🤖  AI assistant",
                "⚙️  Account settings",
            ]
        elif is_county:
            county_name = COUNTIES[COUNTIES["id"]==user_info["county_id"]]["name"].values[0] if user_info["county_id"] else "County"
            nav_options = [
                f"🏠  {county_name} dashboard",
                "📤  Submit energy plan",
                "📥  County inbox",
                "📄  Makueni reference",
                "🤖  AI assistant",
                "⚙️  Account settings",
            ]
        else:  # service provider
            nav_options = [
                "📊  National overview",
                "🗺️  County demand data",
                "📤  Submit service provider plan",
                "🤖  AI assistant",
                "⚙️  Account settings",
            ]

        selected = st.radio("", nav_options, label_visibility="collapsed")
        st.divider()

        # Logout
        try:
            authenticator.logout("Sign out", "sidebar")
        except Exception:
            if st.button("Sign out"):
                st.session_state["authentication_status"] = None
                st.rerun()

        st.markdown("""
        <div style="font-size:10px;color:#9c9a8e;line-height:1.7;margin-top:8px">
        <b>Data sources</b><br>
        EPRA Statistics Report FY 2024/25<br>
        Makueni CEP 2023–2032<br>
        KNBS 2019 Census<br><br>
        <b>Challenge</b><br>
        NGDA 2026 · DTU · Challenge 2<br>
        Partner: EPRA Kenya
        </div>
        """, unsafe_allow_html=True)

    return selected

# ── PAGE: NATIONAL OVERVIEW ───────────────────────────────────────────────────
def page_national_overview(role):
    nat = compute_national()
    st.markdown(f"""
    <div class="kw-alert-info">
    <b>{'EPRA Admin' if role=='epra' else role.title()} view</b> — National aggregation computed live from {nat['submitted_count']} submitted county plans ({nat['coverage_pct']}% coverage). Full picture available when all 47 counties submit.
    </div>
    """, unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="kw-metric"><div class="kw-metric-label">Counties submitted</div>
        <div class="kw-metric-value">{nat['submitted_count']} / {nat['total_counties']}</div>
        <div class="kw-metric-delta" style="color:#0f9d7e">{nat['coverage_pct']}% coverage · {nat['overdue_count']} overdue</div></div>""",unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="kw-metric"><div class="kw-metric-label">Weighted electricity access</div>
        <div class="kw-metric-value">{nat['w_elec']}%</div>
        <div class="kw-metric-delta" style="color:#0f9d7e">↑ Population-weighted from submitted plans</div></div>""",unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="kw-metric"><div class="kw-metric-label">Weighted clean cooking</div>
        <div class="kw-metric-value">{nat['w_cooking']}%</div>
        <div class="kw-metric-delta" style="color:#d4891a">Target: 100% by 2028</div></div>""",unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="kw-metric"><div class="kw-metric-label">Total submitted budgets</div>
        <div class="kw-metric-value">KES {nat['total_budget']}B</div>
        <div class="kw-metric-delta" style="color:#0f9d7e">Sum of {nat['submitted_count']} county plans</div></div>""",unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # National targets
    st.markdown("**National energy targets — live progress**")
    tc1, tc2, tc3 = st.columns(3)
    for col, label, current, goal, year, color in [
        (tc1, "Electricity access",  nat['w_elec'],  100, 2030, "#1a6fa3"),
        (tc2, "Clean cooking",       nat['w_cooking'],100, 2028, "#0f9d7e"),
        (tc3, "Clean energy gen.",   82,              100, 2035, "#5b4fc9"),
    ]:
        with col:
            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=current,
                title={"text":f"{label}<br><span style='font-size:11px'>Target {goal}% by {year}</span>","font":{"size":13}},
                gauge={"axis":{"range":[0,100]},"bar":{"color":color},
                       "bgcolor":"#f7f6f2","borderwidth":0,
                       "steps":[{"range":[0,current],"color":color+"22"}]},
                number={"suffix":"%","font":{"size":28}}
            ))
            fig.update_layout(height=180,margin=dict(t=60,b=10,l=20,r=20),paper_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**National generation trend (GWh)**")
        fig = px.line(GEN_TREND, x="Year", y="GWh", markers=True,
                      color_discrete_sequence=["#1a6fa3"])
        fig.update_layout(height=220,margin=dict(t=10,b=10,l=0,r=0),
                          plot_bgcolor="white",paper_bgcolor="white",
                          xaxis=dict(showgrid=False),
                          yaxis=dict(showgrid=True,gridcolor="#f0f0f0",range=[10000,16000]))
        fig.update_traces(line_width=2.5,marker_size=7)
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.markdown("**Generation mix — 90%+ renewable**")
        fig = px.pie(GEN_MIX, names="Source", values="Pct",
                     color="Source",
                     color_discrete_map=dict(zip(GEN_MIX["Source"],GEN_MIX["Color"])))
        fig.update_layout(height=220,margin=dict(t=10,b=10,l=0,r=0),
                          legend=dict(font=dict(size=11)))
        fig.update_traces(textinfo="percent",textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)

    if role == "epra":
        st.markdown("""<div class="kw-alert-danger">
        <b>Equity alert:</b> Marsabit (8%), Wajir (9%), Mandera (11%) and Turkana (12%) have the lowest electricity access and are all overdue on plan submission.
        These counties have the greatest unmet energy need and must be prioritised for direct submission support.
        </div>""", unsafe_allow_html=True)

# ── PAGE: ALL COUNTIES (EPRA view) ─────────────────────────────────────────────
def page_all_counties():
    st.markdown("""<div class="kw-alert-info"><b>EPRA full access:</b> All 47 counties visible. Use filters to isolate groups. Click any county for detailed view.</div>""",unsafe_allow_html=True)

    filter_col, _ = st.columns([2,3])
    with filter_col:
        status_filter = st.selectbox("Filter by status:",
            ["All counties","Submitted ✓","In review","Pending","Overdue"])

    status_map = {"All counties":None,"Submitted ✓":"submitted","In review":"review",
                  "Pending":"pending","Overdue":"overdue"}
    sf = status_map[status_filter]
    display_df = COUNTIES if sf is None else COUNTIES[COUNTIES["status"]==sf]

    # Summary cards
    c1,c2,c3,c4 = st.columns(4)
    for col, s, label, color in [(c1,"submitted","Submitted",   "#0f9d7e"),
                                  (c2,"review",   "In review",  "#d4891a"),
                                  (c3,"pending",  "Pending",    "#7a7870"),
                                  (c4,"overdue",  "Overdue",    "#b33a2c")]:
        with col:
            n = len(COUNTIES[COUNTIES["status"]==s])
            st.markdown(f"""<div style="padding:12px 14px;border-radius:10px;background:#f7f6f2;border-top:3px solid {color}">
            <div style="font-size:22px;font-weight:700;color:{color}">{n}</div>
            <div style="font-size:11px;color:#6b6860;margin-top:2px">{label}</div></div>""",unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # County table
    for _, row in display_df.iterrows():
        badge_class = f"kw-badge-{row['status']}"
        badge_label = {"submitted":"Submitted ✓","review":"In review","pending":"Pending","overdue":"Overdue"}[row['status']]
        elec_color  = "#0f9d7e" if row['elec']>=75 else "#d4891a" if row['elec']>=40 else "#b33a2c"
        with st.expander(f"**{row['name']}** · {row['region']} · {row['pop']//1000:,}K pop"):
            cc1,cc2,cc3,cc4,cc5 = st.columns(5)
            cc1.metric("Electricity access", f"{row['elec']}%")
            cc2.metric("Clean cooking",      f"{row['cooking']}%")
            cc3.metric("Solar GHI",          f"{row['solar']} kWh/m²")
            cc4.metric("MTF demand tier",    f"Tier {row['mtf']}")
            cc5.metric("Target year",        str(row['target_yr']))
            st.markdown(f'<span class="{badge_class}">{badge_label}</span>', unsafe_allow_html=True)
            if row['status']=="overdue":
                st.markdown(f"""<div class="kw-alert-danger" style="margin-top:8px">
                <b>Action required:</b> {row['name']} County is overdue. INEP Regulations 2025 require submission. Reminder sent to county energy committee.
                </div>""", unsafe_allow_html=True)
            elif row['status']=="submitted":
                st.markdown(f"""<div class="kw-alert-success" style="margin-top:8px">
                <b>Plan on file:</b> Budget KES {row['budget']}B · Growth rate {row['growth']}% p.a. · All validation checks passed.
                </div>""", unsafe_allow_html=True)

# ── PAGE: COUNTY DASHBOARD (county view — own county only) ─────────────────────
def page_county_dashboard(county_id, user_name):
    county = COUNTIES[COUNTIES["id"]==county_id]
    if county.empty:
        st.error("County not found in system. Contact EPRA.")
        return
    row = county.iloc[0]
    status_color = {"submitted":"#0f9d7e","review":"#d4891a","pending":"#7a7870","overdue":"#b33a2c"}[row['status']]

    st.markdown(f"""<div class="kw-alert-info">
    <b>County committee view — {row['name']} County only.</b> You can only see and manage your county's data. Contact EPRA at Allan.Wairimu@epra.go.ke for national data access.
    </div>""", unsafe_allow_html=True)

    # Status banner
    badge_label = {"submitted":"✓ Plan submitted","review":"Plan in review","pending":"Plan pending","overdue":"⚠ Plan overdue"}[row['status']]
    st.markdown(f"""
    <div style="padding:14px 18px;border-radius:10px;border:1.5px solid {status_color};background:{status_color}11;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center">
      <div>
        <div style="font-size:16px;font-weight:700;color:#1a1916">{row['name']} County Energy Plan</div>
        <div style="font-size:12px;color:#6b6860;margin-top:2px">Plan period: 2023–{row['target_yr']} · Region: {row['region']}</div>
      </div>
      <span style="font-size:13px;font-weight:600;color:{status_color};background:{status_color}20;padding:5px 14px;border-radius:20px">{badge_label}</span>
    </div>
    """, unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Electricity access",  f"{row['elec']}%",  "Target: 100%")
    c2.metric("Clean cooking",       f"{row['cooking']}%","Target: 100% by 2028")
    c3.metric("Solar GHI",           f"{row['solar']} kWh/m²")
    c4.metric("Universal access yr", str(row['target_yr']))

    # How county compares to national
    nat = compute_national()
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**How your county compares to national averages (from submitted counties)**")
    compare_df = pd.DataFrame({
        "Indicator": ["Electricity access (%)", "Clean cooking (%)"],
        row['name']:  [row['elec'],              row['cooking']],
        "National avg":[nat['w_elec'],            nat['w_cooking']],
    })
    fig = go.Figure()
    fig.add_trace(go.Bar(name=row['name'],    x=compare_df["Indicator"], y=compare_df[row['name']],    marker_color="#1a6fa3"))
    fig.add_trace(go.Bar(name="National avg", x=compare_df["Indicator"], y=compare_df["National avg"], marker_color="#c8c6be"))
    fig.update_layout(barmode="group", height=240, margin=dict(t=10,b=10,l=0,r=0),
                      plot_bgcolor="white", paper_bgcolor="white",
                      legend=dict(orientation="h",y=1.1))
    st.plotly_chart(fig, use_container_width=True)

    if row['status']=="overdue":
        st.markdown("""<div class="kw-alert-danger">
        <b>Your plan is overdue.</b> Under Section 5(5)(a) of the Energy Act 2019 and the INEP Regulations 2025, your county is legally required to submit a County Energy Plan. Use the "Submit energy plan" tab to submit now.
        </div>""", unsafe_allow_html=True)

# ── PAGE: SUBMIT PLAN ─────────────────────────────────────────────────────────
def page_submit(role, county_id, user_name):
    is_county = role == "county"
    county_name = COUNTIES[COUNTIES["id"]==county_id]["name"].values[0] if county_id else ""

    if is_county:
        st.markdown(f"""<div class="kw-alert-info">
        <b>Submitting as: {county_name} County Energy Planning Committee.</b> Your submission goes directly to EPRA for validation. You can upload an existing PDF plan or fill the structured template.
        </div>""", unsafe_allow_html=True)

    # Pathway selection
    pathway = st.radio("Choose submission pathway:",
        ["📄 Upload existing PDF/Word plan", "📝 Fill structured template"],
        horizontal=True)
    st.divider()

    if "📄" in pathway:
        st.markdown("### Upload your County Energy Plan document")
        st.markdown("""<div class="kw-alert-success">
        <b>How this works:</b> Upload your completed CEP as a PDF or Word document. Our AI automatically extracts the key energy indicators. Review and correct the extracted values, then submit to EPRA.
        </div>""", unsafe_allow_html=True)

        uploaded = st.file_uploader(
            "Choose your County Energy Plan (PDF or Word)",
            type=["pdf","doc","docx"],
            help="Maximum 50MB. The Makueni CEP PDF is a good test file."
        )

        if uploaded:
            st.success(f"✓ File received: **{uploaded.name}** ({uploaded.size//1024} KB)")
            with st.spinner("AI is extracting key indicators from your document…"):
                import time; time.sleep(2)

            st.markdown("""<div class="kw-alert-success">
            <b>AI extraction complete.</b> Key indicators identified from your document. Please review each field below and correct any values before running validation.
            </div>""", unsafe_allow_html=True)

            # Pre-fill with extracted values (simulated for demo)
            extracted = {"elec":75.1,"cooking":17.9,"firewood":72.5,"solar":2008.0,
                         "budget":74.9,"target_yr":2028,"mtf":3,"growth":1.1}

            col1, col2 = st.columns(2)
            with col1:
                county_input = st.text_input("County name *", value=county_name if is_county else "Makueni", disabled=is_county)
                elec_input   = st.number_input("Electricity access (%)*", value=extracted["elec"], min_value=0.0, max_value=100.0)
                cooking_input= st.number_input("Clean cooking access (%)*", value=extracted["cooking"], min_value=0.0, max_value=100.0)
                firewood_input=st.number_input("Firewood as primary cooking fuel (%)", value=extracted["firewood"], min_value=0.0, max_value=100.0)
            with col2:
                solar_input  = st.number_input("Solar GHI (kWh/m²/year)*", value=extracted["solar"], min_value=0.0)
                budget_input = st.number_input("Total plan budget (KES billions)", value=extracted["budget"], min_value=0.0)
                target_input = st.number_input("Universal access target year*", value=extracted["target_yr"], min_value=2025, max_value=2040)
                mtf_input    = st.selectbox("MTF demand tier*", [1,2,3,4,5], index=extracted["mtf"]-1,
                                             help="Multi-Tier Framework tier for electrification modelling")
                growth_input = st.number_input("Population growth rate (%/year)*", value=extracted["growth"], min_value=0.0, max_value=10.0)

            _run_validation_and_submit(county_input, elec_input, cooking_input, firewood_input,
                                        solar_input, budget_input, target_input, mtf_input, growth_input,
                                        uploaded.name, user_name)
        else:
            st.info("Upload your CEP document above to continue. You can use the Makueni County Energy Plan PDF as a test.")

    else:
        st.markdown("### Structured submission template")
        st.markdown("""<div class="kw-alert-info">
        <b>National assumptions pre-loaded:</b> KNBS 2019 population baselines · electricity cost benchmarks (0.047–0.059 $/kWh) · solar GHI reference range (1,600–2,200 kWh/m²) · Kenya's official targets.
        </div>""", unsafe_allow_html=True)

        with st.expander("▶ Section 1 — County profile & electricity access", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                county_input  = st.text_input("County name *", value=county_name, disabled=is_county)
                elec_input    = st.number_input("Total electricity access (%)*", min_value=0.0, max_value=100.0, help="Grid + mini-grid + solar home systems")
                cooking_input = st.number_input("Clean cooking access (%)*", min_value=0.0, max_value=100.0)
                firewood_input= st.number_input("Firewood as primary fuel (%)", min_value=0.0, max_value=100.0)
            with col2:
                solar_input  = st.number_input("Solar GHI (kWh/m²/year)*", min_value=0.0, help="Use Global Solar Atlas · Kenya range: 1,600–2,200")
                budget_input = st.number_input("Total plan budget (KES billions)", min_value=0.0)
                target_input = st.number_input("Universal access target year*", value=2030, min_value=2025, max_value=2040)
                mtf_input    = st.selectbox("MTF demand tier*", [1,2,3,4,5], index=2,
                                             help="Tier 1=minimal, Tier 5=high demand")
                growth_input = st.number_input("Population growth rate (%/year)*", value=1.1, min_value=0.0, max_value=10.0)

        with st.expander("▶ Section 2 — Energy resources", expanded=False):
            rc1, rc2 = st.columns(2)
            with rc1:
                st.number_input("Wind speed at 100m (m/s)", min_value=0.0, help="Use Global Wind Atlas")
                st.number_input("Woody biomass supply (tonnes/year)", min_value=0.0)
            with rc2:
                st.number_input("Hydropower potential (MW)", min_value=0.0)
                st.number_input("Biogas potential from livestock (GJ/year)", min_value=0.0)

        with st.expander("▶ Section 3 — Energy efficiency", expanded=False):
            st.number_input("LED bulb adoption in households (%)", min_value=0.0, max_value=100.0, value=79.0)
            st.number_input("Solar PV installed in county facilities (kW)", min_value=0.0)

        _run_validation_and_submit(county_input if 'county_input' in dir() else county_name,
                                    elec_input if 'elec_input' in dir() else 0,
                                    cooking_input if 'cooking_input' in dir() else 0,
                                    firewood_input if 'firewood_input' in dir() else 0,
                                    solar_input if 'solar_input' in dir() else 0,
                                    budget_input if 'budget_input' in dir() else 0,
                                    target_input if 'target_input' in dir() else 2030,
                                    mtf_input if 'mtf_input' in dir() else 3,
                                    growth_input if 'growth_input' in dir() else 1.1,
                                    None, user_name)

def _run_validation_and_submit(county, elec, cooking, firewood, solar, budget, target, mtf, growth, filename, user_name):
    if st.button("▶  Run validation checks", type="primary"):
        errors, warnings = [], []
        if not county:
            errors.append(("V10","County name is required — mandatory field."))
        if elec < 0 or elec > 100:
            errors.append(("V1",f"Electricity access {elec}% is outside valid range 0–100%."))
        if firewood > 0 and cooking > 0 and abs((firewood + cooking) - 100) > 10:
            warnings.append(("V2",f"Cooking fuel split: firewood ({firewood}%) + clean cooking ({cooking}%) = {firewood+cooking:.1f}%. Confirm all fuel types sum to 100%."))
        if solar > 0 and (solar < 1600 or solar > 2200):
            warnings.append(("V4",f"Solar GHI {solar} kWh/m² is outside Kenya's expected range (1,600–2,200). Check units — may be MJ/m² not kWh/m²."))
        if not mtf:
            errors.append(("V7","MTF demand tier must be declared — required for national aggregation."))
        if growth > 5:
            warnings.append(("V3",f"Population growth rate {growth}% exceeds 5% — verify against KNBS data."))
        if target < 2025 or target > 2040:
            warnings.append(("V8",f"Target year {target} is outside plausible range 2025–2040."))
        if budget > 0 and budget < 0.5:
            warnings.append(("V5",f"Budget KES {budget}B is very low — confirm units are billions KES."))

        total_checks = 10
        passed = total_checks - len(errors) - len(warnings)

        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("Critical errors",    len(errors),   delta="block submission" if errors else None, delta_color="inverse")
        rc2.metric("Warnings",           len(warnings), delta="review before submit" if warnings else None, delta_color="off")
        rc3.metric("Checks passed",      passed,        delta=f"of {total_checks}")

        for rule, msg in errors:
            st.error(f"**ERROR · Rule {rule}:** {msg}")
        for rule, msg in warnings:
            st.warning(f"**WARNING · Rule {rule}:** {msg}")

        if not errors:
            st.success("✓ No critical errors. Ready to submit to EPRA.")
            if st.button("✅ Submit plan to EPRA", type="primary"):
                ref = f"CEP-{county[:2].upper()}-2026-{hash(county)%9000+1000}"
                st.balloons()
                st.success(f"**Plan submitted successfully!** Reference: `{ref}`")
                st.markdown(f"""<div class="kw-alert-success">
                <b>{county} County Energy Plan received by EPRA.</b><br>
                Submitted by: {user_name} · {datetime.now().strftime('%d %b %Y %H:%M')}<br>
                {f'Document: {filename}' if filename else 'Pathway: Structured template'}<br>
                EPRA will validate and confirm within 14 days. Your county will appear as "Submitted" on the national map.
                </div>""", unsafe_allow_html=True)
                # Audit log
                if "audit_log" not in st.session_state:
                    st.session_state.audit_log = []
                st.session_state.audit_log.append({
                    "time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    "user": user_name, "action": f"Plan submitted — {county}",
                    "ref":  ref
                })

# ── PAGE: VALIDATION QUEUE (EPRA only) ───────────────────────────────────────
def page_validation_queue():
    st.markdown("""<div class="kw-alert-warn">
    <b>EPRA admin only.</b> Review submissions flagged by the automated validation engine before including in national aggregation.
    </div>""", unsafe_allow_html=True)

    review_counties = COUNTIES[COUNTIES["status"]=="review"]
    if review_counties.empty:
        st.info("No submissions currently in review queue.")
    else:
        for _, row in review_counties.iterrows():
            with st.expander(f"**{row['name']} County** — {row['pop']//1000:,}K pop · submitted 3 days ago"):
                c1,c2,c3 = st.columns(3)
                c1.metric("Electricity access", f"{row['elec']}%")
                c2.metric("Solar GHI",           f"{row['solar']} kWh/m²")
                c3.metric("Target year",          str(row['target_yr']))
                st.markdown("""<div class="kw-alert-warn">
                <b>2 warnings flagged:</b> Solar GHI slightly above expected range · Population growth assumption >2%
                <br>No critical errors — county may submit but EPRA review recommended.
                </div>""", unsafe_allow_html=True)
                approve_col, reject_col, _ = st.columns([1,1,2])
                if approve_col.button(f"✅ Approve — {row['name']}", key=f"approve_{row['id']}"):
                    st.success(f"{row['name']} approved and included in national aggregation.")
                if reject_col.button(f"↩ Request resubmission", key=f"reject_{row['id']}"):
                    st.warning(f"Resubmission request sent to {row['name']} County.")

    st.divider()
    st.markdown("### National aggregation trigger")
    nat = compute_national()
    st.info(f"Currently: {nat['submitted_count']} counties approved · {len(review_counties)} in review · {nat['pending_count'] + nat['overdue_count']} not submitted")
    if st.button("▶  Run national aggregation now", type="primary"):
        with st.spinner("Aggregating 47 county submissions into national INEP…"):
            import time; time.sleep(2)
        st.success(f"National aggregation complete. INEP updated with data from {nat['submitted_count']} counties.")
        st.metric("Weighted national electricity access", f"{nat['w_elec']}%")
        st.metric("Weighted national clean cooking",      f"{nat['w_cooking']}%")
        st.metric("Total plan investment needed",         f"KES {nat['total_budget']}B")

# ── PAGE: INBOX ───────────────────────────────────────────────────────────────
def page_inbox(role, county_id, user_name):
    is_epra   = role in ("epra","ministry")
    is_county = role == "county"

    if is_epra:
        st.markdown("""<div class="kw-alert-info"><b>EPRA communications hub.</b> Messages from all 47 counties and system notifications.</div>""",unsafe_allow_html=True)
        msgs = [
            {"from":"System","to":"All counties","date":"2026-06-01","type":"Assumptions","subject":"INEP 2025 planning assumptions updated",
             "body":"KNBS 2024 population projections and updated MTF cost benchmarks published. All counties should use these for 2026 submissions."},
            {"from":"Nakuru County","to":"EPRA","date":"2026-05-28","type":"Query","subject":"Solar GHI calculation methodology",
             "body":"We are using Global Solar Atlas but getting different values. Please clarify expected data source."},
            {"from":"System","to":"Turkana, Marsabit, Mandera, Wajir","date":"2026-05-20","type":"Reminder","subject":"Overdue submission — 3rd notice",
             "body":"Your County Energy Plan is overdue under INEP Regulations 2025. Immediate submission required."},
        ]
    else:
        county_name = COUNTIES[COUNTIES["id"]==county_id]["name"].values[0] if county_id else "Your county"
        st.markdown(f"""<div class="kw-alert-info"><b>County inbox — national guidance from EPRA.</b> Messages and assumptions for {county_name} County.</div>""",unsafe_allow_html=True)
        msgs = [
            {"from":"EPRA","to":county_name,"date":"2026-06-01","type":"Assumptions","subject":"Updated national planning assumptions — June 2026",
             "body":"Use updated KNBS baselines and MTF cost benchmarks. Solar GHI reference 1,600–2,200 kWh/m². Electricity target: 100% by 2030."},
            {"from":"EPRA","to":county_name,"date":"2026-05-15","type":"Benchmark","subject":"Peer county comparison — your region",
             "body":"Your electricity access is within regional average. Clean cooking is below regional average — prioritise LPG and biogas interventions."},
            {"from":"EPRA","to":county_name,"date":"2026-04-30","type":"Guidance","subject":"Makueni CEP available as reference template",
             "body":"The Makueni County Energy Plan (2023–2032) is available as reference. It covers all 7 required sectors with example values for all mandatory fields."},
        ]

    type_colors = {"Assumptions":"#1a6fa3","Query":"#d4891a","Reminder":"#b33a2c","Benchmark":"#5b4fc9","Guidance":"#0f9d7e"}
    for m in msgs:
        tc = type_colors.get(m["type"],"#7a7870")
        with st.container():
            st.markdown(f"""
            <div style="background:white;border:0.5px solid #e8e6de;border-radius:10px;padding:14px;margin-bottom:8px">
              <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">
                <div>
                  <span style="font-size:10px;font-weight:600;padding:2px 8px;border-radius:8px;background:{tc}20;color:{tc};text-transform:uppercase;margin-right:8px">{m['type']}</span>
                  <span style="font-size:13px;font-weight:600;color:#1a1916">{m['subject']}</span>
                </div>
                <span style="font-size:11px;color:#9c9a8e">{m['date']}</span>
              </div>
              <div style="font-size:11px;color:#9c9a8e;margin-bottom:6px">From: {m['from']} → To: {m['to']}</div>
              <div style="font-size:12px;color:#6b6860;line-height:1.6">{m['body']}</div>
            </div>
            """, unsafe_allow_html=True)

# ── PAGE: MAKUENI REFERENCE ───────────────────────────────────────────────────
def page_makueni():
    st.markdown("""<div class="kw-alert-success">
    <b>Reference plan:</b> Makueni County Energy Plan 2023–2032. Developed with WRI and Strathmore University. Used as the KenyaWatts submission template reference. Upload this PDF in the Submit tab to test AI extraction.
    </div>""", unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Electricity access", "75.1%",  "Grid 29.2% + Mini-grid 5.7% + SHS 40.2%")
    c2.metric("Solar GHI",          "2,008 kWh/m²", "PV output 4.35 kWh/kWp/day")
    c3.metric("Total budget",        "KES 74.9B",    "2023–2032 plan")
    c4.metric("Clean cooking",       "17.9%",  "⚠ Firewood still 72.5%")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Primary cooking fuel — households 2022**")
        fig = px.bar(MAKUENI_COOKING, x="Pct", y="Fuel", orientation="h",
                     color="Fuel",
                     color_discrete_map={"Firewood":"#b33a2c","LPG":"#0f9d7e","Charcoal":"#7a7870",
                                         "Biogas":"#5b4fc9","Electric":"#1a6fa3","Other":"#c8c6be"})
        fig.update_layout(showlegend=False, height=220,
                          margin=dict(t=10,b=10,l=0,r=0),
                          plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("**Monthly outage duration vs EPRA benchmark**")
        colors = ["#b33a2c" if h>5 else "#0f9d7e" for h in OUTAGE["Hours"]]
        fig = go.Figure(go.Bar(x=OUTAGE["Month"], y=OUTAGE["Hours"],
                               marker_color=colors, name="Actual"))
        fig.add_hline(y=5, line_dash="dash", line_color="#1a6fa3",
                      annotation_text="EPRA benchmark (5 hrs)")
        fig.update_layout(height=220, margin=dict(t=10,b=10,l=0,r=0),
                          plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("**Electrification scenarios (OnSSET modelling) — investment required**")
    scen_df = pd.DataFrame({
        "Scenario":["Low demand (2028)","High demand (2028)","Grid intensification (2028)"],
        "Grid MW":[19.3, 42.7, 38.4],
        "Solar PV MW":[2.34, 53.6, 0],
        "Investment USD M":[132.5, 360.0, 571.8]
    })
    st.dataframe(scen_df, use_container_width=True, hide_index=True)

# ── PAGE: AI ASSISTANT ────────────────────────────────────────────────────────
def page_ai(role, county_id):
    if "ai_history" not in st.session_state:
        st.session_state.ai_history = []

    st.markdown("""<div class="kw-alert-info">
    <b>KenyaWatts AI assistant</b> — powered by Claude · grounded in live EPRA data and Makueni CEP · plain English answers
    </div>""", unsafe_allow_html=True)

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
        for i, s in enumerate(suggestions):
            if cols[i%2].button(s, key=f"sug_{i}"):
                st.session_state.ai_history.append({"role":"user","content":s})
                with st.spinner("Analysing EPRA data…"):
                    reply = ask_ai(s, st.session_state.ai_history[:-1])
                st.session_state.ai_history.append({"role":"assistant","content":reply})
                st.rerun()

    # Chat history
    for msg in st.session_state.ai_history:
        with st.chat_message(msg["role"], avatar="⚡" if msg["role"]=="assistant" else None):
            st.write(msg["content"])

    # Input
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

# ── MAIN APP ──────────────────────────────────────────────────────────────────
def main():
    authenticator = setup_auth()
    auth_status, username = show_login_page(authenticator) if \
        st.session_state.get("authentication_status") is not True else \
        (True, st.session_state.get("username"))

    # Store in session
    if auth_status:
        st.session_state["authentication_status"] = True
        st.session_state["username"] = username

    if not st.session_state.get("authentication_status"):
        show_login_page(authenticator)
        return

    # Get user info
    username  = st.session_state.get("username","")
    user_info = get_user_role(username)
    role      = user_info["role"]
    county_id = user_info["county_id"]

    # Header
    role_label = {"epra":"EPRA Planner","ministry":"Ministry of Energy",
                  "devpartner":"Development Partner","county":"County Committee",
                  "service_provider":"KPLC Service Provider"}.get(role,"User")
    st.markdown(f"""
    <div class="kw-header">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <div>
          <div class="kw-logo">Kenya<span>Watts</span>
            <span style="font-size:14px;font-weight:400;color:#5a8a9e;margin-left:12px">Digital Integrated National Energy Planning Platform</span>
          </div>
        </div>
        <div style="text-align:right;font-size:12px;color:#5a8a9e">
          <span style="color:#3ecfaa">●</span> Live · EPRA Kenya · NGDA 2026
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Build sidebar and get selected page
    selected = build_sidebar(authenticator, user_info)

    # Route to correct page
    key = selected.split("  ",1)[-1].strip()

    if "National overview" in key or "overview" in key.lower():
        page_national_overview(role)
    elif "All 47" in key or "County demand" in key:
        page_all_counties()
    elif "dashboard" in key.lower() or "My county" in key:
        page_county_dashboard(county_id, user_info["name"])
    elif "Validation queue" in key:
        page_validation_queue()
    elif "Submit" in key:
        page_submit(role, county_id, user_info["name"])
    elif "Inbox" in key or "Communications" in key:
        page_inbox(role, county_id, user_info["name"])
    elif "Makueni" in key:
        page_makueni()
    elif "AI" in key:
        page_ai(role, county_id)
    elif "Account settings" in key or "settings" in key.lower():
        page_account_settings(authenticator, username, user_info)

if __name__ == "__main__":
    main()

# ── PAGE: ACCOUNT SETTINGS ────────────────────────────────────────────────────
def page_account_settings(authenticator, username, user_info):
    st.markdown("""<div class="kw-alert-info">
    <b>Account settings</b> — Change your password or update your display name.
    Changes take effect immediately. Role and county access cannot be changed here — contact EPRA admin.
    </div>""", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["🔑  Change password", "✏️  Update display name", "👤  My account info"])

    # ── TAB 1: CHANGE PASSWORD ────────────────────────────────────────────
    with tab1:
        st.markdown("### Change your password")
        st.markdown("""<div class="kw-alert-warn">
        <b>Password requirements:</b> Minimum 8 characters. Mix of letters, numbers and symbols recommended.
        Your new password is active immediately for this session.
        </div>""", unsafe_allow_html=True)

        with st.form("change_password_form"):
            current_pw = st.text_input("Current password",      type="password", placeholder="Your current password")
            new_pw     = st.text_input("New password",           type="password", placeholder="At least 8 characters")
            confirm_pw = st.text_input("Confirm new password",   type="password", placeholder="Repeat your new password")
            submit_pw  = st.form_submit_button("Update password", type="primary")

        if submit_pw:
            if not current_pw or not new_pw or not confirm_pw:
                st.error("All three fields are required.")
            elif len(new_pw) < 8:
                st.error("New password must be at least 8 characters.")
            elif new_pw != confirm_pw:
                st.error("New password and confirmation do not match.")
            elif new_pw == current_pw:
                st.warning("New password must be different from your current password.")
            else:
                import bcrypt as _bcrypt
                try:
                    stored = st.secrets.get("credentials", {}).get("usernames", {}).get(username, {}).get("password", "")
                    if stored and _bcrypt.checkpw(current_pw.encode(), stored.encode()):
                        new_hash = _bcrypt.hashpw(new_pw.encode(), _bcrypt.gensalt()).decode()
                        # Update in-session authenticator store
                        try:
                            authenticator.credentials["usernames"][username]["password"] = new_hash
                        except Exception:
                            pass
                        # Audit log
                        if "audit_log" not in st.session_state:
                            st.session_state.audit_log = []
                        st.session_state.audit_log.append({
                            "time":   datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            "user":   username,
                            "action": "Password changed",
                            "ref":    f"PW-{username.upper()}-{datetime.now().strftime('%Y%m%d%H%M')}"
                        })
                        st.success("✓ Password updated successfully.")
                        st.info("To make this permanent across all sessions, the EPRA admin must update secrets.toml with the new hash below:")
                        st.code(f'[credentials.usernames.{username}]\npassword = "{new_hash}"', language="toml")
                    else:
                        st.error("Current password is incorrect. Please try again.")
                except Exception as e:
                    st.error(f"Error: {str(e)}")

    # ── TAB 2: UPDATE DISPLAY NAME ────────────────────────────────────────
    with tab2:
        st.markdown("### Update your display name and email")
        st.info("Your display name appears on submissions and in the platform header. It does not affect your login username.")

        current_name  = st.session_state.get("user_display_name", {}).get(username, user_info.get("name", username))
        current_email = st.session_state.get("user_email", {}).get(username, user_info.get("email", ""))

        with st.form("update_name_form"):
            new_name  = st.text_input("Display name",    value=current_name,  placeholder="Your full name or title")
            new_email = st.text_input("Email address",   value=current_email, placeholder="name@organisation.go.ke")
            submit_nm = st.form_submit_button("Save changes", type="primary")

        if submit_nm:
            if len(new_name.strip()) < 3:
                st.error("Display name must be at least 3 characters.")
            else:
                if "user_display_name" not in st.session_state:
                    st.session_state.user_display_name = {}
                if "user_email" not in st.session_state:
                    st.session_state.user_email = {}
                st.session_state.user_display_name[username] = new_name.strip()
                st.session_state.user_email[username] = new_email.strip()
                try:
                    authenticator.credentials["usernames"][username]["name"]  = new_name.strip()
                    authenticator.credentials["usernames"][username]["email"] = new_email.strip()
                except Exception:
                    pass
                if "audit_log" not in st.session_state:
                    st.session_state.audit_log = []
                st.session_state.audit_log.append({
                    "time":   datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    "user":   username,
                    "action": f"Display name updated to: {new_name.strip()}",
                    "ref":    f"NAME-{username.upper()}"
                })
                st.success(f"✓ Display name updated to **{new_name.strip()}**.")

    # ── TAB 3: ACCOUNT INFO ───────────────────────────────────────────────
    with tab3:
        st.markdown("### Your account information")
        role_labels = {
            "epra":             "EPRA Planner — Full admin access",
            "ministry":         "Ministry of Energy — Read-only national view",
            "devpartner":       "Development Partner — Read-only aggregated view",
            "county":           "County Energy Planning Committee — Own county only",
            "service_provider": "National Energy Service Provider (KPLC)",
        }
        role_descs = {
            "epra":             "All 47 counties · Validation queue · Aggregation engine · Admin tools",
            "ministry":         "National dashboards · Target progress trackers · Policy view",
            "devpartner":       "National aggregates · County comparison · Investment opportunities",
            "county":           "Own county dashboard · Submit plans · Receive national guidance",
            "service_provider": "National demand data · Service provider submission portal",
        }
        role      = user_info["role"]
        county_id = user_info["county_id"]
        dname     = st.session_state.get("user_display_name", {}).get(username, user_info.get("name", username))
        demail    = st.session_state.get("user_email", {}).get(username, user_info.get("email", "Not set"))
        cnty      = COUNTIES[COUNTIES["id"]==county_id]["name"].values[0] if county_id else "All counties (national access)"

        rows = [
            ("Username",         username),
            ("Display name",     dname),
            ("Email",            demail or "Not set"),
            ("Role",             role_labels.get(role, role)),
            ("Access scope",     role_descs.get(role, "Standard")),
            ("County",           cnty),
            ("Session started",  datetime.now().strftime("%d %b %Y %H:%M")),
            ("Account status",   "Active ✓"),
        ]
        for label, value in rows:
            col_l, col_r = st.columns([1, 2])
            col_l.markdown(f"<div style='font-size:12px;font-weight:600;color:#9c9a8e;padding:5px 0'>{label}</div>", unsafe_allow_html=True)
            col_r.markdown(f"<div style='font-size:13px;color:#1a1916;padding:5px 0'>{value}</div>", unsafe_allow_html=True)

        st.divider()
        # Recent activity
        audit = [a for a in st.session_state.get("audit_log", []) if a.get("user") == username]
        if audit:
            st.markdown("**Your recent activity**")
            for entry in reversed(audit[-5:]):
                st.markdown(f"""<div style="display:flex;gap:12px;font-size:11px;padding:4px 0;border-bottom:0.5px solid #f0ede4">
                  <span style="color:#9c9a8e;font-family:monospace;flex-shrink:0">{entry['time']}</span>
                  <span style="color:#1a1916">{entry['action']}</span>
                  <span style="color:#9c9a8e;font-family:monospace">{entry.get('ref','')}</span>
                </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""<div class="kw-alert-info">
        <b>Need to change your username or role?</b> These are managed by the EPRA platform administrator.
        Contact: Allan.Wairimu@epra.go.ke · +254720850696
        </div>""", unsafe_allow_html=True)
