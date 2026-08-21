"""
Automated Certificate Generation and Email Sending System
DV Analytics

Run with:  streamlit run app.py
"""

import os
import io
import zipfile

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from certificate_generator import (
    load_template_as_image,
    generate_certificate,
    render_certificate_image,
    suggest_text_style,
    suggest_name_position,
)
from email_sender import EmailSender, SMTPConfig, is_valid_email
from utils import (
    validate_excel_columns,
    normalize_records,
    build_report_dataframe,
    log_event,
    now_str,
    LOG_PATH,
)

load_dotenv()

OUTPUT_DIR = "output"
CERT_DIR = os.path.join(OUTPUT_DIR, "certificates")
REPORT_PATH = os.path.join(OUTPUT_DIR, "Email_Sending_Report.xlsx")
ERROR_REPORT_PATH = os.path.join(OUTPUT_DIR, "Error_Report.xlsx")

os.makedirs(CERT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Page config & modern theme
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="DV Analytics — Certificate & Email Suite",
    page_icon="🎓",
    layout="wide",
)


# ---------------------------------------------------------------------------
# Login gate — must run before any other UI renders
# ---------------------------------------------------------------------------
def login():
    st.markdown(
        """
        <style>
        [data-testid="stHeader"] { background: transparent; }
        [data-testid="stToolbar"] { display: none; }
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }

        @keyframes auroraDrift {
            0%   { transform: translate(0, 0) scale(1); }
            50%  { transform: translate(-3%, 2%) scale(1.08); }
            100% { transform: translate(0, 0) scale(1); }
        }
        @keyframes floatY {
            0%, 100% { transform: translateY(0px); }
            50%      { transform: translateY(-8px); }
        }
        @keyframes glowPulse {
            0%, 100% { box-shadow: 0 0 0 10px rgba(45,212,191,.08), 0 12px 35px rgba(16,185,129,.35); }
            50%      { box-shadow: 0 0 0 16px rgba(45,212,191,.14), 0 16px 44px rgba(16,185,129,.5); }
        }
        @keyframes shimmer {
            0%   { background-position: -300px 0; }
            100% { background-position: 300px 0; }
        }

        .stApp {
            background:
                radial-gradient(circle at 8% 82%, rgba(20, 184, 166, .26), transparent 30%),
                radial-gradient(circle at 92% 78%, rgba(16, 185, 129, .22), transparent 32%),
                radial-gradient(circle at 55% 8%, rgba(45, 212, 191, .16), transparent 40%),
                linear-gradient(135deg, #01110d 0%, #042e24 40%, #063d2f 75%, #011a14 100%);
            min-height: 100vh;
            position: relative;
            overflow-x: hidden;
        }
        .stApp::before {
            content: "";
            position: fixed; inset: -10%;
            background:
                radial-gradient(circle at 20% 30%, rgba(20,184,166,.10), transparent 35%),
                radial-gradient(circle at 80% 70%, rgba(16,185,129,.10), transparent 35%);
            animation: auroraDrift 16s ease-in-out infinite;
            pointer-events: none;
            z-index: 0;
        }
        .block-container { max-width: 620px !important; padding-top: 3rem !important; padding-bottom: 1rem !important; position: relative; z-index: 1; }

        .login-top { display: flex; justify-content: center; margin-bottom: 22px; }
        .login-top-badge {
            display: inline-flex; align-items: center; gap: 9px; padding: 8px 18px;
            border: 1px solid rgba(45,212,191,.35); border-radius: 999px; color: #99f6e4;
            background: linear-gradient(90deg, rgba(45,212,191,.10), rgba(255,255,255,.04));
            font-size: 12.5px; font-weight: 600; letter-spacing: .4px; text-transform: uppercase;
            backdrop-filter: blur(12px);
        }
        .login-top-badge .dot { width: 6px; height: 6px; border-radius: 50%; background: #34d399; box-shadow: 0 0 0 3px rgba(52,211,153,.25); display: inline-block; }

        .brand-wrap { text-align: center; margin: 5px auto 30px; }
        .dv-logo {
            width: 78px; height: 78px; margin: 0 auto 20px; border-radius: 20px;
            display: flex; align-items: center; justify-content: center; color: white;
            font-size: 32px; font-weight: 800; letter-spacing: -2px;
            background: linear-gradient(150deg, #2dd4bf 0%, #10b981 55%, #065f46 100%);
            border: 1px solid rgba(255,255,255,.28);
            box-shadow: 0 18px 45px rgba(16,185,129,.35), inset 0 1px 0 rgba(255,255,255,.25);
            position: relative;
            animation: floatY 5s ease-in-out infinite;
        }
        .dv-logo:after {
            content: ""; position: absolute; right: -2px; top: -2px; width: 26px; height: 26px;
            border-radius: 0 20px 0 20px; background: #ea1313;
            box-shadow: 0 2px 8px rgba(234,19,19,.5);
        }
        .brand-title {
            font-size: clamp(30px, 4.2vw, 46px); line-height: 1.1; font-weight: 800; letter-spacing: -1.6px; margin: 0;
            background: linear-gradient(90deg, #ffffff 20%, #99f6e4 50%, #34d399 80%);
            -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent;
        }
        .brand-subtitle { color: rgba(220,252,241,.72); font-size: 15.5px; margin-top: 11px; letter-spacing: .1px; }

        .login-card {
            max-width: 560px; margin: 0 auto; padding: 38px 38px 32px; border-radius: 26px;
            background: linear-gradient(155deg, rgba(255,255,255,.11), rgba(255,255,255,.03));
            border: 1px solid rgba(255,255,255,.16);
            box-shadow: 0 32px 90px rgba(0,0,0,.45), inset 0 1px 0 rgba(255,255,255,.10);
            backdrop-filter: blur(26px);
            position: relative;
            overflow: hidden;
        }
        .login-card:before {
            content: ""; position: absolute; top: 0; left: -100%; width: 60%; height: 100%;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,.06), transparent);
            transform: skewX(-20deg);
        }

        .lock-circle {
            width: 72px; height: 72px; margin: 0 auto 18px; border-radius: 50%;
            display: flex; align-items: center; justify-content: center; font-size: 30px;
            background: linear-gradient(135deg, #2dd4bf, #10b981 60%, #065f46);
            animation: glowPulse 3s ease-in-out infinite;
        }
        .welcome-title {
            text-align: center; font-size: 27px; font-weight: 750; margin: 0;
            background: linear-gradient(90deg, #ffffff, #d1fae5);
            -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent;
        }
        .welcome-text { text-align: center; color: rgba(232,238,255,.68); margin: 8px 0 26px; font-size: 14.5px; }

        .stTextInput > label { color: #6ee7d1 !important; font-weight: 600 !important; font-size: 13.5px !important; letter-spacing: .2px; text-transform: uppercase; }
        .stTextInput > div > div {
            background: rgba(2,20,16,.45) !important; border: 1px solid rgba(255,255,255,.16) !important;
            border-radius: 13px !important; min-height: 54px !important; transition: .2s ease;
        }
        .stTextInput > div > div:focus-within {
            border-color: #2dd4bf !important; box-shadow: 0 0 0 3px rgba(45,212,191,.18) !important;
        }
        .stTextInput input { color: #ffffff !important; font-size: 15px !important; }
        .stTextInput input::placeholder { color: rgba(255,255,255,.38) !important; }

        .login-actions { display: flex; justify-content: space-between; align-items: center; color: rgba(238,243,255,.70); font-size: 13px; margin: 4px 2px 20px; }
        .remember { display: flex; gap: 7px; align-items: center; }
        .remember .tick { color: #34d399; font-weight: 800; }
        .login-actions .secure-access { display: flex; align-items: center; gap: 6px; color: #6ee7d1; font-weight: 600; }

        .secure-line { display: flex; align-items: center; gap: 14px; color: rgba(236,242,255,.75); margin-top: 26px; font-size: 12.5px; text-transform: uppercase; letter-spacing: .6px; font-weight: 600; }
        .secure-line:before, .secure-line:after { content: ""; height: 1px; flex: 1; background: linear-gradient(90deg, transparent, rgba(255,255,255,.18), transparent); }
        .secure-content {
            display: flex; align-items: center; gap: 13px; margin-top: 20px; padding: 15px 16px; border-radius: 14px;
            background: linear-gradient(120deg, rgba(52,211,153,.08), rgba(4,19,55,.30));
            border: 1px solid rgba(52,211,153,.18);
            color: rgba(235,241,255,.75); font-size: 13px;
        }
        .secure-icon { font-size: 24px; }
        .login-footer { text-align: center; color: rgba(225,233,255,.48); font-size: 12px; margin: 28px auto 6px; }
        .help-footer { max-width: 560px; margin: 18px auto 0; padding: 16px 18px; border-top: 1px solid rgba(255,255,255,.10); display: flex; justify-content: space-between; color: rgba(230,237,255,.55); font-size: 13px; }

        .stButton > button {
            width: 100% !important; min-height: 56px !important; border: 0 !important; border-radius: 13px !important;
            color: white !important; font-size: 16px !important; font-weight: 750 !important; letter-spacing: .2px;
            background: linear-gradient(90deg, #065f46 0%, #10b981 45%, #2dd4bf 75%, #5eead4 100%) !important;
            background-size: 200% auto !important;
            box-shadow: 0 14px 34px rgba(16,185,129,.38) !important;
            transition: transform .18s ease, box-shadow .18s ease, background-position .4s ease !important;
        }
        .stButton > button:hover {
            transform: translateY(-2px); background-position: right center !important;
            box-shadow: 0 18px 44px rgba(45,212,191,.35) !important;
        }
        .stButton > button:active { transform: translateY(0px) scale(.99); }

        @media (max-width: 700px) {
            .block-container { padding: 1rem !important; }
            .login-card { padding: 26px 22px 24px; }
            .brand-title { font-size: 30px; }
            .help-footer { flex-direction: column; gap: 8px; text-align: center; }
        }
        </style>

        <div class="login-top"><div class="login-top-badge"><span class="dot"></span> Secure Login</div></div>
        <div class="brand-wrap">
            <div class="dv-logo">DV</div>
            <div class="brand-title">Certificate Email Automation</div>
            <div class="brand-subtitle">Automated certificate generation and email delivery</div>
        </div>
        <div class="login-card">
            <div class="lock-circle">🔒</div>
            <div class="welcome-title">Welcome Back!</div>
            <div class="welcome-text">Please sign in to continue to your dashboard</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    username = st.text_input("Username", value="", placeholder="Enter your email address", key="login_username")
    password = st.text_input("Password", type="password", placeholder="Enter your password", key="login_password")

    st.markdown(
        '<div class="login-actions"><span class="remember"><span class="tick">✓</span> Keep me signed in</span>'
        '<span class="secure-access">🔐 Secure access</span></div>',
        unsafe_allow_html=True,
    )

    if st.button("🚀  Login", key="login_button"):
        # Credentials come from Streamlit secrets / environment variables only —
        # never hardcoded in source, since this repo is public on GitHub.
        valid_username = os.environ.get("APP_USERNAME", "")
        valid_password = os.environ.get("APP_PASSWORD", "")

        if not valid_username or not valid_password:
            st.error("⚠️ Login is not configured. Set APP_USERNAME and APP_PASSWORD in secrets.")
        elif username == valid_username and password == valid_password:
            st.session_state["logged_in"] = True
            st.rerun()
        else:
            st.error("❌ Invalid username or password")

    st.markdown(
        """
        <div class="secure-line">Secure Access</div>
        <div class="secure-content">
            <span class="secure-icon">🛡️</span>
            <div>
                <strong style="color:#eef4ff;">Your connection is protected</strong><br>
                <span>Authorized access only · DV Analytics Certificate & Email Suite</span>
            </div>
        </div>
        <div class="login-footer">© 2026 DV Analytics · Certificate &amp; Email Automation Suite</div>
        <div class="help-footer"><span>DV Analytics</span><span>Need help? Contact your administrator</span></div>
        """,
        unsafe_allow_html=True,
    )

    return st.session_state.get("logged_in", False)


if not st.session_state.get("logged_in", False):
    login()
    st.stop()


NAVY = "#0B1B4D"
NAVY_DEEP = "#03143f"
RED = "#EF233C"
ACCENT = "#3b8cff"
ACCENT_2 = "#a832ff"
GOLD = "#D4AF37"
BG = "#03143f"
CARD = "rgba(255,255,255,.06)"
BORDER = "rgba(255,255,255,.14)"
MUTED = "rgba(232,238,255,.68)"
TEXT = "#EAF0FF"

st.markdown(
    f"""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@500;600;700;800&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
        .stApp {{
            background:
                radial-gradient(circle at 8% 82%, rgba(0, 183, 255, .20), transparent 27%),
                radial-gradient(circle at 92% 78%, rgba(210, 40, 255, .19), transparent 30%),
                radial-gradient(circle at 50% 0%, rgba(65, 93, 255, .14), transparent 38%),
                linear-gradient(135deg, #03143f 0%, #071b4f 45%, #13052f 100%);
            min-height: 100vh;
        }}
        #MainMenu, footer, header[data-testid="stHeader"] {{ visibility: hidden; }}

        /* ---------- Global text colors for dark theme ---------- */
        .stApp, .stApp p, .stApp span, .stApp label, .stApp li,
        .stApp .stMarkdown, .stApp .stCaption, [data-testid="stCaptionContainer"] {{
            color: {TEXT};
        }}
        .stApp h1, .stApp h2, .stApp h3, .stApp h4 {{ color: #ffffff; }}

        /* ---------- Hero ---------- */
        .dv-hero {{
            position: relative; overflow: hidden;
            background: radial-gradient(130% 180% at 0% 0%, {ACCENT_2}3d 0%, transparent 45%),
                        radial-gradient(120% 160% at 100% 100%, {RED}26 0%, transparent 40%),
                        linear-gradient(145deg, rgba(255,255,255,.10), rgba(255,255,255,.03));
            padding: 32px 36px; border-radius: 22px; margin-bottom: 28px;
            display: flex; align-items: center; justify-content: space-between; gap: 18px;
            box-shadow: 0 20px 45px -18px rgba(0,0,0,.55), inset 0 1px 0 rgba(255,255,255,.08);
            border: 1px solid rgba(255,255,255,.14);
            backdrop-filter: blur(18px);
        }}
        .dv-hero-left {{ display: flex; align-items: center; gap: 18px; }}
        .dv-hero .mark {{
            width: 54px; height: 54px; border-radius: 15px; flex-shrink: 0;
            background: linear-gradient(135deg, {RED}, #ff7a7a);
            display: flex; align-items: center; justify-content: center;
            color: white; font-weight: 800; font-size: 20px; letter-spacing: -1px;
            font-family: 'Poppins', sans-serif;
            box-shadow: 0 8px 20px -4px {RED}aa, inset 0 1px 0 rgba(255,255,255,.25);
        }}
        .dv-hero h1 {{
            color: white; font-size: 24px; margin: 0; font-family: 'Poppins', sans-serif; font-weight: 700;
            letter-spacing: -.3px;
        }}
        .dv-hero p {{ color: #B7C2F5; font-size: 13px; margin: 4px 0 0; }}
        .dv-hero-pill {{
            display: flex; align-items: center; gap: 8px; color: #EAF0FF; font-size: 12.5px;
            font-weight: 600; padding: 8px 14px; border-radius: 999px;
            background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.14);
            backdrop-filter: blur(8px); white-space: nowrap;
        }}
        .dv-hero-pill .dot {{
            width: 7px; height: 7px; border-radius: 50%; background: #34d399;
            box-shadow: 0 0 0 3px rgba(52,211,153,.25);
        }}

        /* ---------- Glass cards ---------- */
        .dv-card {{
            background: linear-gradient(145deg, rgba(255,255,255,.09), rgba(255,255,255,.035));
            border: 1px solid {BORDER}; border-radius: 20px;
            padding: 24px 26px; margin-bottom: 20px;
            box-shadow: 0 18px 45px -22px rgba(0,0,0,.55), inset 0 1px 0 rgba(255,255,255,.06);
            backdrop-filter: blur(18px);
            transition: box-shadow .2s ease, border-color .2s ease;
        }}
        .dv-card:hover {{ box-shadow: 0 22px 55px -22px rgba(0,0,0,.65); border-color: rgba(255,255,255,.22); }}
        .dv-section-title {{
            font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 17.5px;
            color: #ffffff; margin-bottom: 2px; display: flex; align-items: center; gap: 8px;
        }}
        .dv-section-title:before {{
            content: ""; display: inline-block; width: 6px; height: 18px; border-radius: 4px;
            background: linear-gradient(180deg, {ACCENT}, {RED});
        }}
        .dv-section-sub {{ color: {MUTED}; font-size: 13px; margin: 4px 0 16px 14px; }}

        /* ---------- Tabs as pill nav ---------- */
        div[data-testid="stTabs"] div[data-baseweb="tab-list"] {{
            gap: 6px; background: rgba(255,255,255,.05); padding: 6px; border-radius: 14px;
            border: 1px solid {BORDER};
        }}
        div[data-testid="stTabs"] button[data-baseweb="tab"] {{
            font-family: 'Poppins', sans-serif; font-weight: 600; font-size: 13.5px;
            color: {MUTED}; padding: 10px 18px; border-radius: 10px; transition: all .15s ease;
        }}
        div[data-testid="stTabs"] button[data-baseweb="tab"] p {{ color: inherit; }}
        div[data-testid="stTabs"] button[aria-selected="true"] {{
            color: white !important;
            background: linear-gradient(135deg, #096dff 0%, #7140ff 55%, #a82cff 100%);
            box-shadow: 0 6px 18px -6px rgba(120,80,255,.55);
        }}
        div[data-testid="stTabs"] button[aria-selected="true"] p {{ color: white !important; }}
        div[data-testid="stTabs"] div[data-baseweb="tab-highlight"] {{ display: none; }}
        div[data-testid="stTabs"] div[data-baseweb="tab-border"] {{ display: none; }}

        /* ---------- Metrics ---------- */
        div[data-testid="stMetric"] {{
            background: linear-gradient(145deg, rgba(255,255,255,.09), rgba(255,255,255,.03));
            border: 1px solid {BORDER}; border-radius: 16px;
            padding: 16px 18px; box-shadow: 0 6px 18px -12px rgba(0,0,0,.4);
            border-top: 3px solid {ACCENT};
            backdrop-filter: blur(14px);
        }}
        div[data-testid="stMetricLabel"] {{ color: {MUTED} !important; font-weight: 600; font-size: 12.5px; text-transform: uppercase; letter-spacing: .3px; }}
        div[data-testid="stMetricValue"] {{ color: #ffffff !important; font-family: 'Poppins', sans-serif; font-weight: 700; }}

        /* ---------- Advanced buttons ---------- */
        .stButton>button {{
            position: relative; overflow: hidden;
            background: linear-gradient(135deg, #096dff 0%, #7140ff 55%, #a82cff 100%);
            background-size: 200% auto;
            color: white; border-radius: 12px; font-weight: 700; border: none;
            padding: 0.68em 1.5em; font-family: 'Poppins', sans-serif; font-size: 14.5px;
            box-shadow: 0 10px 26px -8px rgba(62,78,255,.45), inset 0 1px 0 rgba(255,255,255,.18);
            transition: all .22s ease; letter-spacing: .1px;
        }}
        .stButton>button:hover {{
            transform: translateY(-2px); background-position: right center;
            box-shadow: 0 16px 34px -10px rgba(120,80,255,.55), inset 0 1px 0 rgba(255,255,255,.25); color: white;
        }}
        .stButton>button:active {{ transform: translateY(0px) scale(.99); }}
        .stButton>button:disabled {{
            background: rgba(255,255,255,.08); color: rgba(234,240,255,.35); box-shadow: none; transform: none;
        }}
        .stDownloadButton>button {{
            background: rgba(255,255,255,.06); color: #EAF0FF; border: 1.5px solid rgba(255,255,255,.20); border-radius: 12px;
            font-weight: 700; font-family: 'Poppins', sans-serif; font-size: 13.5px;
            transition: all .18s ease; padding: 0.6em 1.2em; backdrop-filter: blur(10px);
        }}
        .stDownloadButton>button:hover {{
            border-color: {ACCENT}; color: {ACCENT}; transform: translateY(-1px);
            box-shadow: 0 8px 20px -10px {ACCENT}aa;
        }}

        /* Primary CTA buttons (generate / send) get an extra glow */
        div[data-testid="stTabs"] .stButton>button[kind="secondary"],
        button[kind="primary"] {{
            background: linear-gradient(135deg, {RED} 0%, {ACCENT_2} 100%) !important;
        }}

        .dv-badge {{
            display: inline-block; padding: 4px 13px; border-radius: 999px;
            font-size: 12px; font-weight: 700; font-family: 'Poppins', sans-serif;
            letter-spacing: .2px;
        }}
        .dv-badge-ok {{ background: rgba(52,211,153,.16); color: #34d399; border: 1px solid rgba(52,211,153,.3); }}
        .dv-badge-warn {{ background: rgba(239,35,60,.16); color: #ff8398; border: 1px solid rgba(239,35,60,.3); }}

        div[data-testid="stProgress"] > div > div {{
            background: linear-gradient(90deg, {ACCENT}, {RED}); border-radius: 999px;
        }}
        div[data-testid="stProgress"] {{ background: rgba(255,255,255,.08); border-radius: 999px; }}

        /* ---------- File uploader ---------- */
        [data-testid="stFileUploaderDropzone"] {{
            background: linear-gradient(180deg, rgba(255,255,255,.05), rgba(255,255,255,.02)) !important;
            border: 1.5px dashed rgba(255,255,255,.25) !important; border-radius: 14px !important;
        }}
        [data-testid="stFileUploaderDropzone"] * {{ color: {TEXT} !important; }}
        [data-testid="stFileUploaderDropzone"] small {{ color: {MUTED} !important; }}
        [data-testid="stFileUploaderDropzone"] button {{
            background: rgba(255,255,255,.08) !important; color: #EAF0FF !important;
            border: 1px solid rgba(255,255,255,.2) !important;
        }}

        /* ---------- Inputs, sliders, checkboxes, color picker ---------- */
        .stTextInput > div > div, .stTextArea > div > div {{
            background: rgba(255,255,255,.05) !important; border: 1px solid rgba(255,255,255,.16) !important;
            border-radius: 12px !important;
        }}
        .stTextInput > div > div:focus-within, .stTextArea > div > div:focus-within {{
            border-color: {ACCENT} !important; box-shadow: 0 0 0 3px rgba(59,140,255,.16) !important;
        }}
        .stTextInput input, .stTextArea textarea {{ color: #ffffff !important; }}
        .stTextInput input::placeholder, .stTextArea textarea::placeholder {{ color: rgba(255,255,255,.35) !important; }}
        .stSlider [data-baseweb="slider"] > div > div {{ background: rgba(255,255,255,.12) !important; }}
        .stSlider [role="slider"] {{ background: {ACCENT} !important; box-shadow: 0 0 0 6px rgba(59,140,255,.16) !important; }}
        .stCheckbox label span, .stCheckbox p {{ color: {TEXT} !important; }}
        [data-testid="stColorPickerBlock"] {{ border: 1px solid {BORDER}; border-radius: 10px; overflow: hidden; }}

        /* ---------- Expander ---------- */
        details[data-testid="stExpander"] {{
            background: linear-gradient(145deg, rgba(255,255,255,.06), rgba(255,255,255,.02));
            border: 1px solid {BORDER}; border-radius: 14px; backdrop-filter: blur(10px);
        }}
        details[data-testid="stExpander"] summary {{ color: {TEXT} !important; font-weight: 600; }}

        /* ---------- Dataframe ---------- */
        [data-testid="stDataFrame"] {{
            border: 1px solid {BORDER}; border-radius: 14px; overflow: hidden;
        }}

        /* ---------- Alerts (info / warning / success / error) ---------- */
        div[data-testid="stAlert"] {{
            background: rgba(255,255,255,.06) !important; border: 1px solid {BORDER} !important;
            border-radius: 12px !important; backdrop-filter: blur(10px);
        }}
        div[data-testid="stAlert"] p {{ color: {TEXT} !important; }}

        /* ---------- Sidebar ---------- */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, {NAVY_DEEP}, {NAVY}) !important;
            border-right: 1px solid rgba(255,255,255,.08);
        }}
        section[data-testid="stSidebar"] * {{ color: #EAF0FF !important; }}
        section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,.12) !important; }}
    </style>

    <div class="dv-hero">
        <div class="dv-hero-left">
            <div class="mark">DV</div>
            <div>
                <h1>Certificate &amp; Email Suite</h1>
                <p>Automated certificate generation and personalized email delivery</p>
            </div>
        </div>
        <div class="dv-hero-pill"><span class="dot"></span> DV Analytics Workspace</div>
    </div>
    """,
    unsafe_allow_html=True,
)


def card_start(title: str, subtitle: str = ""):
    sub_html = f'<div class="dv-section-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="dv-card"><div class="dv-section-title">{title}</div>{sub_html}',
        unsafe_allow_html=True,
    )


def card_end():
    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
defaults = {
    "records": [],
    "column_map": {},
    "template_path": None,
    "template_fingerprint": None,
    "cert_paths": {},
    "results": [],
    "font_size": 60,
    "y_pos_pct": 50,
    "text_color_hex": "#0B1B4D",
    "confirmed_params": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------------------------------------------------------------------------
# Sidebar — SMTP status
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("#### ✉️ SMTP status")
    cfg = SMTPConfig()
    if cfg.is_configured():
        st.markdown('<span class="dv-badge dv-badge-ok">● Connected</span>', unsafe_allow_html=True)
        st.caption(f"{cfg.username}\nvia {cfg.host}:{cfg.port}")
    else:
        st.markdown('<span class="dv-badge dv-badge-warn">● Not configured</span>', unsafe_allow_html=True)
        st.caption(
            "Set **SMTP_EMAIL** and **SMTP_PASSWORD** as environment variables / "
            "Streamlit secrets. Gmail requires an **App Password**, not your normal login."
        )
    st.divider()
    st.caption("DV Analytics · Certificate & Email Suite")

tab1, tab2, tab3, tab4 = st.tabs(
    ["①  Upload", "②  Generate", "③  Send", "④  Dashboard"]
)

# ---------------------------------------------------------------------------
# TAB 1 — Upload
# ---------------------------------------------------------------------------
with tab1:
    card_start("Participant list", "Required fields: Name, Mobile Number, Email ID — headers are matched flexibly.")
    excel_file = st.file_uploader("Excel file (.xlsx)", type=["xlsx"], label_visibility="collapsed")

    if excel_file:
        try:
            df = pd.read_excel(excel_file)
            is_valid, column_map, missing = validate_excel_columns(df)
            if not is_valid:
                st.error(f"Missing required column(s): {', '.join(missing)}")
            else:
                records = normalize_records(df, column_map)
                st.session_state.records = records
                st.session_state.column_map = column_map
                mapping_str = " · ".join(f"{k} → `{v}`" for k, v in column_map.items())
                st.markdown(f'<span class="dv-badge dv-badge-ok">✓ {len(records)} records validated</span>', unsafe_allow_html=True)
                st.caption(mapping_str)
                st.dataframe(pd.DataFrame(records), use_container_width=True, height=220)
        except Exception as e:
            st.error(f"Could not read Excel file: {e}")
    card_end()

    card_start("Certificate template", "PNG, JPG, or PDF — the name is drawn centered on top of this image.")
    template_file = st.file_uploader("Certificate template", type=["png", "jpg", "jpeg", "pdf"], label_visibility="collapsed")

    if template_file:
        template_path = os.path.join(OUTPUT_DIR, f"template{os.path.splitext(template_file.name)[1]}")
        with open(template_path, "wb") as f:
            f.write(template_file.getbuffer())
        st.session_state.template_path = template_path

        fingerprint = (template_file.name, template_file.size)
        is_new_template = fingerprint != st.session_state.template_fingerprint

        template_img = load_template_as_image(template_path)

        if is_new_template:
            st.session_state.template_fingerprint = fingerprint
            # Auto-detect a blank band on the template first (fixes the name
            # landing on top of printed text like "has successfully
            # participated in..."), then tune font size/color for that spot.
            suggested_y = suggest_name_position(template_img)
            suggested_size, suggested_color = suggest_text_style(template_img, suggested_y)
            st.session_state.y_pos_pct = int(round(suggested_y * 100))
            st.session_state.font_size = suggested_size
            st.session_state.text_color_hex = "#%02x%02x%02x" % suggested_color
            st.session_state.confirmed_params = None
            st.markdown('<span class="dv-badge dv-badge-ok">✓ Template uploaded — name placement auto-detected</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="dv-badge dv-badge-ok">✓ Template uploaded</span>', unsafe_allow_html=True)

        with st.expander("🎨 Customize name placement, size & color", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                font_size = st.slider("Font size (px)", 20, 300, key="font_size")
                y_pos = st.slider("Vertical position (% down)", 0, 100, key="y_pos_pct") / 100.0
            with c2:
                color_hex = st.color_picker("Text color", key="text_color_hex")
                if st.button("↺ Auto-fit to this template"):
                    s_y = suggest_name_position(template_img)
                    s_size, s_color = suggest_text_style(template_img, s_y)
                    st.session_state.y_pos_pct = int(round(s_y * 100))
                    st.session_state.font_size = s_size
                    st.session_state.text_color_hex = "#%02x%02x%02x" % s_color
                    st.session_state.confirmed_params = None
                    st.rerun()

        font_size = st.session_state.font_size
        y_pos = st.session_state.y_pos_pct / 100.0
        color_hex = st.session_state.text_color_hex
        text_color = tuple(int(color_hex.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))

        if st.session_state.records:
            sample_name = st.session_state.records[0]["Name"]
            preview_img = render_certificate_image(
                template_img, sample_name, None, font_size, text_color, y_pos
            )
            st.image(preview_img, caption=f"Live preview — {sample_name}", use_container_width=True)

            current_params = (fingerprint, font_size, round(y_pos, 3), color_hex)
            confirmed = st.checkbox(
                "✅ I can clearly see the name above on the certificate",
                value=(st.session_state.confirmed_params == current_params),
            )
            if confirmed:
                st.session_state.confirmed_params = current_params
            elif st.session_state.confirmed_params == current_params:
                st.session_state.confirmed_params = None
        else:
            st.info("Upload the participant list above to preview a sample name on this template.")
    card_end()

# ---------------------------------------------------------------------------
# TAB 2 — Generate Certificates
# ---------------------------------------------------------------------------
with tab2:
    card_start("Generate certificates")

    records = st.session_state.records
    template_path = st.session_state.template_path

    c1, c2, c3 = st.columns(3)
    c1.metric("Total records", len(records))
    c2.metric("Template ready", "Yes" if template_path else "No")
    c3.metric("Certificates generated", len(st.session_state.cert_paths))

    fingerprint = st.session_state.template_fingerprint
    font_size = st.session_state.font_size
    y_pos = st.session_state.y_pos_pct / 100.0
    color_hex = st.session_state.text_color_hex
    text_color = tuple(int(color_hex.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    current_params = (fingerprint, font_size, round(y_pos, 3), color_hex)
    preview_confirmed = st.session_state.confirmed_params == current_params and fingerprint is not None

    disabled = not (records and template_path)
    if disabled:
        st.info("Upload both the participant list and the certificate template in Step ① first.")
    elif not preview_confirmed:
        st.warning("Go back to Step ① and confirm the name preview looks correct before generating in bulk.")

    if st.button("🎓  Generate Certificates", disabled=disabled or not preview_confirmed, key="btn_generate"):
        template_img = load_template_as_image(template_path)
        progress = st.progress(0, text="Starting...")
        used_names = {}
        cert_paths = {}
        for i, rec in enumerate(records):
            name = rec["Name"]
            try:
                path = generate_certificate(
                    template_img, name, CERT_DIR,
                    font_path=None, font_size=font_size,
                    text_color=text_color, y_position_pct=y_pos,
                    used_names=used_names,
                )
                cert_paths[name] = path
            except Exception as e:
                cert_paths[name] = None
                log_event(rec.get("Email ID", ""), "CERT_GENERATION_FAILED", str(e))
            progress.progress((i + 1) / len(records), text=f"Generating {i+1} of {len(records)} — {name}")
        st.session_state.cert_paths = cert_paths
        progress.empty()
        ok_count = sum(1 for v in cert_paths.values() if v)
        st.success(f"Done — {ok_count} of {len(records)} certificates generated.")

    if st.session_state.cert_paths:
        st.markdown("**Review**")
        review_df = pd.DataFrame(
            [{"Name": n, "Certificate Generated": "Yes" if p else "No"}
             for n, p in st.session_state.cert_paths.items()]
        )
        st.dataframe(review_df, use_container_width=True, height=220)

        sample_paths = [p for p in st.session_state.cert_paths.values() if p]
        if sample_paths:
            with st.expander("🔍 Preview a generated certificate"):
                st.image(load_template_as_image(sample_paths[0]), use_container_width=True)
    card_end()

# ---------------------------------------------------------------------------
# TAB 3 — Send Certificates
# ---------------------------------------------------------------------------
with tab3:
    card_start("Compose email")

    subject = st.text_input("Subject", "Congratulations! Your Certificate is Ready")
    body_template = st.text_area(
        "Body (use {{Name}} to insert the participant's name)",
        value=(
            "Dear {{Name}},\n\n"
            "Thank you for participating in our program.\n"
            "Please find your certificate attached to this email.\n\n"
            "We appreciate your participation and wish you all the best for your "
            "future endeavors.\n\n"
            "Regards,\nDV Analytics Team"
        ),
        height=200,
    )
    card_end()

    card_start("Send")
    records = st.session_state.records
    cert_paths = st.session_state.cert_paths

    total = len(records)
    sent_count = sum(1 for r in st.session_state.results if r.get("Email Sent") == "Yes")
    failed_count = sum(1 for r in st.session_state.results if r.get("Email Sent") == "No")
    pending_count = total - sent_count - failed_count

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Records", total)
    m2.metric("Sent Successfully", sent_count)
    m3.metric("Failed", failed_count)
    m4.metric("Pending", max(pending_count, 0))

    can_send = bool(records) and bool(cert_paths) and SMTPConfig().is_configured()
    if not records or not cert_paths:
        st.info("Generate certificates in Step ② before sending.")
    elif not SMTPConfig().is_configured():
        st.warning("SMTP is not configured — see the sidebar.")

    if st.button("✉️  Send Certificates", disabled=not can_send, key="btn_send"):
        progress = st.progress(0, text="Starting...")
        results = []
        sender = EmailSender()
        try:
            sender.connect()
        except Exception as e:
            st.error(f"Could not connect to SMTP server: {e}")
            sender = None

        for i, rec in enumerate(records):
            name = rec["Name"]
            mobile = rec.get("Mobile Number", "")
            email = rec.get("Email ID", "")
            cert_path = cert_paths.get(name)
            row = {
                "Name": name,
                "Mobile Number": mobile,
                "Email ID": email,
                "Certificate Generated": "Yes" if cert_path else "No",
                "Email Sent": "No",
                "Sent Date & Time": "",
                "Error Message": "",
            }

            progress.progress((i + 1) / total, text=f"Sending {i+1} of {total} — {name}")

            if not cert_path:
                row["Error Message"] = "Certificate not generated"
                log_event(email, "SKIPPED", "Certificate not generated")
                results.append(row)
                continue
            if not is_valid_email(email):
                row["Error Message"] = "Invalid email address"
                log_event(email, "FAILED", "Invalid email address")
                results.append(row)
                continue
            if sender is None:
                row["Error Message"] = "SMTP connection unavailable"
                results.append(row)
                continue

            personalized_body = body_template.replace("{{Name}}", name)
            try:
                sender.send(email, subject, personalized_body, cert_path)
                row["Email Sent"] = "Yes"
                row["Sent Date & Time"] = now_str()
                log_event(email, "SENT")
            except Exception as e:
                row["Error Message"] = str(e)
                log_event(email, "FAILED", str(e))

            results.append(row)

        if sender is not None:
            sender.close()

        st.session_state.results = results
        progress.empty()

        report_df = build_report_dataframe(results)
        report_df.to_excel(REPORT_PATH, index=False)
        errors_df = report_df[report_df["Error Message"] != ""]
        errors_df.to_excel(ERROR_REPORT_PATH, index=False)

        n_sent = sum(1 for r in results if r["Email Sent"] == "Yes")
        st.success(f"Done — {n_sent} of {total} emails sent successfully.")
        st.rerun()
    card_end()

# ---------------------------------------------------------------------------
# TAB 4 — Report & Dashboard
# ---------------------------------------------------------------------------
with tab4:
    card_start("Dashboard")

    records = st.session_state.records
    cert_paths = st.session_state.cert_paths
    results = st.session_state.results

    total = len(records)
    certs_generated = sum(1 for v in cert_paths.values() if v)
    emails_sent = sum(1 for r in results if r.get("Email Sent") == "Yes")
    emails_failed = sum(1 for r in results if r.get("Email Sent") == "No")
    success_rate = f"{(emails_sent / total * 100):.1f}%" if total else "0.0%"

    d1, d2, d3, d4, d5 = st.columns(5)
    d1.metric("Total Participants", total)
    d2.metric("Certificates Generated", certs_generated)
    d3.metric("Emails Sent", emails_sent)
    d4.metric("Failed Emails", emails_failed)
    d5.metric("Success Rate", success_rate)
    card_end()

    card_start("Downloads")
    dl1, dl2, dl3, dl4 = st.columns(4)

    with dl1:
        if cert_paths and any(cert_paths.values()):
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, path in cert_paths.items():
                    if path and os.path.exists(path):
                        zf.write(path, arcname=os.path.basename(path))
            st.download_button("⬇️ Certificates (.zip)", buf.getvalue(), file_name="Certificates.zip", mime="application/zip")
        else:
            st.button("⬇️ Certificates (.zip)", disabled=True)

    with dl2:
        if os.path.exists(REPORT_PATH):
            with open(REPORT_PATH, "rb") as f:
                st.download_button("⬇️ Email Report", f.read(), file_name="Email_Sending_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.button("⬇️ Email Report", disabled=True)

    with dl3:
        if os.path.exists(ERROR_REPORT_PATH):
            with open(ERROR_REPORT_PATH, "rb") as f:
                st.download_button("⬇️ Error Report", f.read(), file_name="Error_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.button("⬇️ Error Report", disabled=True)

    with dl4:
        if os.path.exists(LOG_PATH):
            with open(LOG_PATH, "rb") as f:
                st.download_button("⬇️ Log File", f.read(), file_name="email_log.txt")
        else:
            st.button("⬇️ Log File", disabled=True)
    card_end()

    if results:
        card_start("Full report")
        st.dataframe(build_report_dataframe(results), use_container_width=True, height=320)
        card_end()
