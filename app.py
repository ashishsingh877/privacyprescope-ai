import streamlit as st
import requests
import json
import re
import os
import time
from docx_generator import generate_questionnaire_docx

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PrivacyScope AI",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.stApp{background:linear-gradient(160deg,#0B1120 0%,#0D1829 60%,#0B1120 100%);}
[data-testid="stSidebar"]{background:rgba(13,20,38,0.98)!important;border-right:1px solid rgba(255,255,255,0.07);}
[data-testid="stSidebar"] *{color:#CBD5E1!important;}

.hero{background:linear-gradient(135deg,rgba(30,58,138,0.35) 0%,rgba(46,117,182,0.2) 100%);
      border:1px solid rgba(46,117,182,0.35);border-radius:16px;padding:32px 36px;
      text-align:center;margin-bottom:22px;position:relative;overflow:hidden;}
.hero::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
              background:linear-gradient(90deg,#C8973A,#2E75B6,#C8973A);}

.stat-card{background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.09);
           border-radius:12px;padding:16px 18px;text-align:center;}
.stat-num{font-size:28px;font-weight:800;color:#2E75B6;}
.stat-lbl{font-size:11px;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin-top:4px;}

.option-card{background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);
             border-radius:10px;padding:14px 16px;margin-bottom:12px;}
.option-title{font-size:11px;font-weight:700;color:#94A3B8;
              text-transform:uppercase;letter-spacing:0.08em;margin-bottom:8px;}
.option-item{font-size:12px;color:#CBD5E1;padding:3px 0;
             border-bottom:1px solid rgba(255,255,255,0.04);display:flex;align-items:center;gap:8px;}
.option-other{font-size:11px;color:#475569;padding:3px 0;font-style:italic;}

.badge{display:inline-block;background:rgba(16,185,129,0.15);color:#10B981;
       border:1px solid rgba(16,185,129,0.3);border-radius:100px;padding:3px 12px;
       font-size:11px;font-weight:600;}
.badge-blue{display:inline-block;background:rgba(46,117,182,0.15);color:#60A5FA;
            border:1px solid rgba(46,117,182,0.3);border-radius:100px;padding:3px 12px;
            font-size:11px;font-weight:600;}

.key-box{background:rgba(16,185,129,0.05);border:1px solid rgba(16,185,129,0.18);
         border-radius:8px;padding:12px 14px;font-size:12px;color:#94A3B8;line-height:1.85;}
.key-ok {background:rgba(16,185,129,0.08);border:1px solid rgba(16,185,129,0.25);
         border-radius:8px;padding:10px 14px;font-size:12px;color:#10B981;margin-top:4px;}

.progress-box{background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);
              border-radius:12px;padding:24px 28px;}

.download-area{background:linear-gradient(135deg,rgba(30,58,138,0.2),rgba(46,117,182,0.15));
               border:1px solid rgba(46,117,182,0.3);border-radius:14px;
               padding:24px 28px;text-align:center;margin:20px 0;}

.stButton>button{background:linear-gradient(135deg,#1E3A8A,#2E75B6)!important;
                 color:white!important;border:none!important;border-radius:8px!important;
                 font-weight:600!important;letter-spacing:0.02em!important;}
.stButton>button:hover{opacity:.92!important;transform:translateY(-1px)!important;}
/* Input fields — always white bg + dark text, visible in both light & dark mode */
.stTextInput>div>div>input,
.stTextInput>div>div>input:focus,
.stTextInput>div>div>input:active,
.stTextInput>div>div>input:hover,
[data-baseweb="input"] input,
[data-baseweb="base-input"] input {
    background:#FFFFFF!important;
    color:#111827!important;
    -webkit-text-fill-color:#111827!important;
    border:1.5px solid #CBD5E1!important;
    border-radius:8px!important;
    font-size:14px!important;
    caret-color:#2E75B6!important;
    opacity:1!important;
}
.stTextInput label{color:#94A3B8!important;font-size:12px!important;font-weight:500!important;}
hr{border-color:rgba(255,255,255,0.07)!important;}
h1,h2,h3{color:#F1F5F9!important;}
</style>
""", unsafe_allow_html=True)


# ─── Groq AI ──────────────────────────────────────────────────────────────────
GROQ_MODELS = [
    "llama-3.3-70b-versatile",
    "llama3-70b-8192",
    "mixtral-8x7b-32768",
    "llama3-8b-8192",
]

SYSTEM = ("You are a senior privacy consultant. "
          "Respond with ONLY valid JSON — no markdown, no extra text.")

PROMPT = """Analyse this organisation for a privacy pre-scoping questionnaire:

Organisation: {org_name}
Website: {website}

Return ONLY this JSON:
{{
  "short_name": "Short abbreviation used in documents (e.g. TCS, HDFC, Infosys)",
  "sector": "Primary industry sector",
  "business_lines": ["Line 1", "Line 2", "Line 3", "Line 4", "Line 5"],
  "stakeholder_teams": ["HR & People Operations", "IT & Cybersecurity", "Legal & Compliance", "Team 4", "Team 5", "Team 6"],
  "customer_interfaces": ["Interface 1", "Interface 2", "Interface 3", "Interface 4", "Interface 5"],
  "core_systems": ["HRIS (Human Resource Information System)", "CRM (Customer Relationship Management)", "ERP (Enterprise Resource Planning)", "System 4", "System 5", "System 6", "System 7"],
  "data_subjects": ["Category 1 (detail)", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6"],
  "data_types": ["Identity Data (ID proofs, Aadhaar, PAN)", "Contact & Demographic Data", "Financial Data (salary, transactions)", "Type 4 (detail)", "Type 5", "Type 6", "Type 7"]
}}

RULES:
- short_name: common abbreviation or first word
- business_lines: 4-6 SPECIFIC to this exact company/sector
- stakeholder_teams: 5-7 teams, always include HR, IT, Legal
- customer_interfaces: 4-6 channels this org type uses
- core_systems: 5-8 systems specific to this sector (e.g. Finacle for banks, SAP for manufacturing)
- data_subjects: 5-7 categories specific to this org's operations
- data_types: 6-8 types with descriptions in brackets where helpful
- Return ONLY raw JSON, nothing else"""


def get_ai_options(org: str, site: str, key: str) -> dict:
    hdrs = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    body = PROMPT.format(
        org_name=org,
        website=site.strip() if site.strip() else "not provided",
    )
    last = None
    for model in GROQ_MODELS:
        try:
            r = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers=hdrs,
                json={
                    "model": model,
                    "messages": [
                        {"role": "system", "content": SYSTEM},
                        {"role": "user",   "content": body},
                    ],
                    "temperature": 0.2,
                    "max_tokens":  2048,
                    "response_format": {"type": "json_object"},
                },
                timeout=45,
            )
            if r.status_code == 429: last = "Rate limit"; continue
            if r.status_code == 401: raise ValueError("Invalid API key — check your Groq key.")
            r.raise_for_status()
            txt = r.json()["choices"][0]["message"]["content"].strip()
            txt = re.sub(r"^```(?:json)?", "", txt).strip()
            txt = re.sub(r"```$", "", txt).strip()
            m = re.search(r"\{[\s\S]*\}", txt)
            if m:
                return json.loads(m.group(0))
        except ValueError: raise
        except Exception as e: last = e; continue
    raise ValueError(f"Could not get AI response. Error: {last}")


# ─── Session state ────────────────────────────────────────────────────────────
for k, v in {"phase": "landing", "ai": None, "org": "", "site": "", "key": ""}.items():
    if k not in st.session_state:
        st.session_state[k] = v

if not st.session_state.key:
    try:    st.session_state.key = st.secrets.get("GROQ_API_KEY", "")
    except: st.session_state.key = os.environ.get("GROQ_API_KEY", "")


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🛡️ PrivacyScope AI")
    st.markdown("*Pre-Scoping Questionnaire Generator*")
    inp = st.text_input("🔑 Groq API Key", type="password",
                        value=st.session_state.key, placeholder="gsk_...",
                        help="Free from console.groq.com")
    if inp:
        st.session_state.key = inp

    if st.session_state.key:
        st.markdown("<div class='key-ok'>✅ API key connected — ready to generate</div>",
                    unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class='key-box'>
        <b style='color:#10B981'>Get FREE key (2 min):</b><br>
        1. Go to <b>console.groq.com</b><br>
        2. Sign up with Google / email<br>
        3. API Keys → <b>Create API Key</b><br>
        4. Copy key (starts <b>gsk_</b>)<br>
        5. Paste above ⬆️<br><br>
        ✅ No credit card · Works in India
        </div>""", unsafe_allow_html=True)

    if st.session_state.phase == "done" and st.session_state.ai:
        st.divider()
        ai = st.session_state.ai
        st.markdown(f"""
        <div style='font-size:12px;color:#94A3B8;line-height:2'>
        <b style='color:#F1F5F9'>Org:</b> {st.session_state.org}<br>
        <b style='color:#F1F5F9'>Sector:</b> {ai.get('sector','—')}<br>
        <b style='color:#F1F5F9'>Business lines:</b> {len(ai.get('business_lines',[]))}<br>
        <b style='color:#F1F5F9'>IT systems:</b> {len(ai.get('core_systems',[]))}
        </div>""", unsafe_allow_html=True)
        st.divider()
        if st.button("🔄 New Questionnaire", use_container_width=True):
            st.session_state.update({"phase":"landing","ai":None,"org":"","site":""})
            st.rerun()


# ─── LANDING ──────────────────────────────────────────────────────────────────
if st.session_state.phase == "landing":

    st.markdown("""
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    for col, icon, num, lbl in zip([c1,c2,c3,c4],
        ["🏢","🤖","☑️","📄"],
        ["01","02","03","04"],
        ["Enter org\n& website","AI tailors\noptions","Review options\nin app","Download\nprofessional .docx"]):
        with col:
            st.markdown(f"""
            <div class='stat-card'>
              <div style='font-size:22px'>{icon}</div>
              <div style='font-size:10px;color:#1E3A8A;font-weight:700;margin:4px 0 2px'>{num}</div>
              <div style='font-size:11px;color:#94A3B8;white-space:pre-line'>{lbl}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")
    L, R = st.columns([1.5, 1])

    with L:
        st.markdown("### 🏢 Organisation Details")
        org  = st.text_input("Organisation Name *", placeholder="e.g. Infosys Limited / HDFC Bank",
                             value=st.session_state.org)
        site = st.text_input("Company Website (recommended)",
                             placeholder="e.g. https://www.infosys.com",
                             value=st.session_state.site)
        if not st.session_state.key:
            st.warning("⚠️ Please enter your Groq API key in the sidebar.")

        go = st.button("⚡  Generate Tailored Questionnaire",
                       use_container_width=True, type="primary")

        if go:
            if not org.strip():
                st.error("Please enter the organisation name.")
            elif not st.session_state.key:
                st.error("Please enter your Groq API key in the sidebar.")
            else:
                st.session_state.org  = org
                st.session_state.site = site

                steps = [
                    ("🔍", "Researching organisation profile…"),
                    ("🎯", "Identifying sector & business lines…"),
                    ("🖥️", "Mapping IT systems & interfaces…"),
                    ("👥", "Profiling data subjects & types…"),
                    ("📋", "Tailoring questionnaire options…"),
                    ("✅", "Ready!"),
                ]
                box = st.empty()

                def render(active):
                    rows = ""
                    for i, (ic, lb) in enumerate(steps):
                        if   i < active:  c_, s_ = "#10B981", "✓"
                        elif i == active: c_, s_ = "#60A5FA", "⟳"
                        else:             c_, s_ = "#334155", "○"
                        rows += (f"<div style='color:{c_};font-size:13px;margin:7px 0;font-weight:600'>"
                                 f"{ic}  {s_}  {lb}</div>")
                    pct = int((active + 1) / len(steps) * 100)
                    box.markdown(f"""
                    <div class='progress-box'>
                      <div style='color:#F1F5F9;font-weight:700;font-size:15px;margin-bottom:14px'>
                        🤖 Analysing: {org}
                      </div>
                      {rows}
                      <div style='margin-top:16px;background:rgba(255,255,255,0.08);
                           border-radius:100px;height:5px'>
                        <div style='width:{pct}%;height:100%;border-radius:100px;
                             background:linear-gradient(90deg,#1E3A8A,#2E75B6,#C8973A);
                             transition:width .3s'></div>
                      </div>
                    </div>""", unsafe_allow_html=True)

                for i in range(len(steps) - 1):
                    render(i); time.sleep(0.4)
                render(len(steps) - 1)

                try:
                    ai = get_ai_options(org, site, st.session_state.key)
                    st.session_state.ai    = ai
                    st.session_state.phase = "done"
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")

    with R:
        st.markdown("### 📄 Document features")
        st.markdown("""
        <div class='option-card'>
        <div style='font-size:13px;color:#94A3B8;line-height:2.2'>
        ☑️ <b style='color:#F1F5F9'>Clickable checkboxes</b> in Word<br>
        🔤 <b style='color:#F1F5F9'>Aptos 11pt font</b> throughout<br>
        🎨 <b style='color:#F1F5F9'>Professional template</b> matching original<br>
        🏷️ <b style='color:#F1F5F9'>Org name</b> replaced in all questions<br>
        📊 <b style='color:#F1F5F9'>AI-tailored options</b> per section<br>
        🔵 <b style='color:#F1F5F9'>Alternating row shading</b><br>
        📌 <b style='color:#F1F5F9'>Header &amp; footer</b> on every page<br>
        🔒 <b style='color:#F1F5F9'>Confidential footer</b> with date<br>
        📝 <b style='color:#F1F5F9'>All fields empty</b> — org fills them
        </div></div>""", unsafe_allow_html=True)

        st.markdown("### 🎯 Example output")
        st.markdown("""
        <div class='option-card'>
        <div style='font-size:12px;color:#94A3B8;line-height:1.9'>
        Input: <b style='color:#60A5FA'>HDFC Bank</b><br><br>
        Business lines: Retail Banking, Corporate Banking, Insurance, Loans, Cards…<br>
        Systems: Finacle Core Banking, Salesforce, Workday, Oracle…<br>
        Data subjects: Account Holders, Loan Applicants, KYC…<br>
        Interfaces: NetBanking Portal, Mobile App, Branches…
        </div></div>""", unsafe_allow_html=True)


# ─── DONE ─────────────────────────────────────────────────────────────────────
elif st.session_state.phase == "done":
    ai    = st.session_state.ai
    org   = st.session_state.org
    short = ai.get("short_name", org.split()[0])

    # Banner
    st.markdown(f"""
    <div style='background:linear-gradient(135deg,rgba(30,58,138,0.25),rgba(46,117,182,0.15));
         border:1px solid rgba(46,117,182,0.35);border-radius:14px;padding:20px 26px;
         margin-bottom:20px;position:relative;overflow:hidden;'>
      <div style='position:absolute;top:0;left:0;right:0;height:3px;
           background:linear-gradient(90deg,#C8973A,#2E75B6,#C8973A)'></div>
      <div style='display:flex;align-items:center;gap:16px'>
        <span style='font-size:36px'>✅</span>
        <div>
          <div style='font-size:18px;font-weight:700;color:#F1F5F9'>{org}</div>
          <div style='font-size:12px;color:#64748B;margin-top:4px'>
          </div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Stats row
    s1, s2, s3, s4, s5 = st.columns(5)
    for col, num, lbl in zip(
        [s1, s2, s3, s4, s5],
        [len(ai.get("business_lines",[])),
         len(ai.get("stakeholder_teams",[])),
         len(ai.get("core_systems",[])),
         len(ai.get("data_subjects",[])),
         len(ai.get("data_types",[]))],
        ["Business\nLines","Stakeholder\nTeams","IT\nSystems","Data\nSubjects","Data\nTypes"],
    ):
        with col:
            st.markdown(f"""
            <div class='stat-card'>
              <div class='stat-num'>{num}</div>
              <div class='stat-lbl'>{lbl}</div>
            </div>""", unsafe_allow_html=True)

    st.divider()

    # Options preview
    st.markdown("### 📋 AI-Generated Options Preview")
    colA, colB = st.columns(2)

    def render_card(col, title, key):
        items = ai.get(key, [])
        rows  = "".join(
            f"<div class='option-item'>☐&nbsp; {it}</div>" for it in items
        )
        rows += "<div class='option-other'>☐&nbsp; Other – Specify: ___________</div>"
        with col:
            st.markdown(f"""
            <div class='option-card'>
              <div class='option-title'>{title}</div>
              {rows}
            </div>""", unsafe_allow_html=True)

    render_card(colA, "📊 Business Lines",       "business_lines")
    render_card(colB, "💻 Core IT Systems",       "core_systems")
    render_card(colA, "👥 Data Subjects",          "data_subjects")
    render_card(colB, "📁 Data Types",             "data_types")
    render_card(colA, "🖥️ Customer Interfaces",   "customer_interfaces")
    render_card(colB, "🏢 Stakeholder Teams",      "stakeholder_teams")

    st.divider()

    # Download
    st.markdown("""
    <div class='download-area'>
      <div style='font-size:22px;margin-bottom:8px'>📄</div>
      <div style='font-size:16px;font-weight:700;color:#F1F5F9;margin-bottom:6px'>
        Professional Word Document Ready
      </div>
    </div>
    """, unsafe_allow_html=True)

    docx_bytes = generate_questionnaire_docx(org, ai)

    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        st.download_button(
            label="⬇️  Download Questionnaire (.docx)",
            data=docx_bytes,
            file_name=f"Pre-Scoping_Privacy_Questionnaire_{org.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    _, mid2, _ = st.columns([2, 1, 2])
    with mid2:
        if st.button("🔄 New Organisation", use_container_width=True):
            st.session_state.update({"phase": "landing", "ai": None})
            st.rerun()
