"""
Skytecher Cold Email Automation - Unified Streamlit Application
================================================================
Single-file production dashboard for uploading Excel client files,
previewing AI-personalized emails, running test sends, and managing
live outreach campaigns — all powered by a Python-native Streamlit UI.

Run with:  python main.py
"""

import os
import sys

# Ensure UTF-8 console encoding on Windows
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

import re
import time
import json
import random
import logging
import threading
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

import streamlit as st

# When run directly with `python main.py`, hand off to streamlit CLI once on port 8501 and exit
if __name__ == "__main__" and not st.runtime.exists():
    import subprocess
    cmd = [sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve()), "--server.port", "8501"] + sys.argv[1:]
    sys.exit(subprocess.call(cmd))

import config
from excel_handler import (
    ExcelHandler,
    is_valid_email,
    FIELD_STATUS,
    FIELD_EMAIL,
    FIELD_COMPANY_NAME,
    FIELD_CLIENT_NAME,
)
from email_writer import EmailWriter
from email_sender import EmailSender

# ---------------------------------------------------------------------------
# Constants & Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    filename=str(config.LOG_FILE),
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding="utf-8",
)
logger = logging.getLogger("SkytecherApp")


# ---------------------------------------------------------------------------
# Campaign Controller  (Thread-safe singleton stored in session_state)
# ---------------------------------------------------------------------------
class CampaignController:
    """Thread-safe campaign state tracker for background email sending."""

    def __init__(self):
        self.lock = threading.Lock()
        self.is_running = False
        self.stop_requested = False
        self.current_file: Optional[Path] = None
        self.total = 0
        self.current_index = 0
        self.success_count = 0
        self.failed_count = 0
        self.skipped_count = 0
        self.current_lead = ""
        self.current_status_msg = "Idle"
        self.delay_remaining = 0
        self.logs: List[Dict[str, Any]] = []

    def log(self, text: str, level: str = "info"):
        timestamp = time.strftime("%H:%M:%S")
        entry = {"time": timestamp, "text": text, "level": level}
        with self.lock:
            self.logs.append(entry)
            if len(self.logs) > 200:
                self.logs.pop(0)

    def reset(self, total: int, file_path: Path):
        with self.lock:
            self.is_running = True
            self.stop_requested = False
            self.current_file = file_path
            self.total = total
            self.current_index = 0
            self.success_count = 0
            self.failed_count = 0
            self.skipped_count = 0
            self.current_lead = ""
            self.current_status_msg = "Starting campaign..."
            self.delay_remaining = 0
            self.logs = []
        self.log(
            f"Campaign started for {file_path.name} with {total} pending leads", "info"
        )

    def get_status(self) -> Dict[str, Any]:
        with self.lock:
            return {
                "is_running": self.is_running,
                "total": self.total,
                "current_index": self.current_index,
                "success_count": self.success_count,
                "failed_count": self.failed_count,
                "skipped_count": self.skipped_count,
                "current_lead": self.current_lead,
                "current_status_msg": self.current_status_msg,
                "delay_remaining": self.delay_remaining,
                "logs": list(self.logs),
            }


def _get_campaign() -> CampaignController:
    """Returns the singleton CampaignController, creating it once per session."""
    if "campaign" not in st.session_state:
        st.session_state["campaign"] = CampaignController()
    return st.session_state["campaign"]


# ---------------------------------------------------------------------------
# Background Campaign Worker
# ---------------------------------------------------------------------------
def background_campaign_worker(
    file_path: Path, min_delay: int, max_delay: int, campaign: CampaignController
):
    """Background thread: generates + sends emails and updates Excel."""
    handler = ExcelHandler(file_path)
    ok, _ = handler.load_file()
    if not ok:
        campaign.log(f"Error opening Excel file: {file_path}", "error")
        with campaign.lock:
            campaign.is_running = False
        return

    clients = handler.get_clients()
    pending = [
        c
        for c in clients
        if str(c.get("status", "")).strip().lower() != "sent"
        and is_valid_email(c.get("email", ""))
    ]

    sender = EmailSender()
    can_send, count_today, limit = sender.can_send_today()
    remaining = max(0, limit - count_today)
    to_send = pending[:remaining]
    writer = EmailWriter()

    for idx, client in enumerate(to_send, 1):
        if campaign.stop_requested:
            campaign.log("Campaign stopped by user.", "warning")
            break

        row_idx = client["_row_index"]
        company = client.get("company_name", "Company")
        email = client.get("email", "")
        name = client.get("client_name", "Business Leader")

        with campaign.lock:
            campaign.current_index = idx
            campaign.current_lead = f"{company} ({email})"
            campaign.current_status_msg = (
                f"Writing personalized email for {company}..."
            )

        campaign.log(
            f"[{idx}/{len(to_send)}] Preparing outreach for {company} <{email}>",
            "info",
        )

        try:
            gen = writer.generate_email(client)
        except Exception as e:
            err = f"AI copy generation error: {e}"
            campaign.log(err, "error")
            handler.update_row_status(row_idx, status="Failed", error_msg=err)
            with campaign.lock:
                campaign.failed_count += 1
            continue

        with campaign.lock:
            campaign.current_status_msg = (
                f"Connecting to SMTP to send email to {email}..."
            )

        send_ok, err_msg = sender.send_email(
            recipient_email=email,
            subject=gen["subject"],
            plain_body=gen["plain_body"],
            html_body=gen["html_body"],
            recipient_name=name,
        )

        if send_ok:
            handler.update_row_status(row_idx, status="Sent", error_msg="")
            campaign.log(
                f"[✓] Sent successfully to {email} | Subject: '{gen['subject']}'",
                "success",
            )
            with campaign.lock:
                campaign.success_count += 1
        else:
            handler.update_row_status(row_idx, status="Failed", error_msg=err_msg)
            campaign.log(f"[✗] Send failed for {email}: {err_msg}", "error")
            with campaign.lock:
                campaign.failed_count += 1

        # Anti-spam delay
        if idx < len(to_send) and not campaign.stop_requested:
            delay = random.randint(min_delay, max_delay)
            campaign.log(f"Anti-spam delay: pausing for {delay} seconds...", "info")
            for rem in range(delay, 0, -1):
                if campaign.stop_requested:
                    break
                with campaign.lock:
                    campaign.delay_remaining = rem
                    campaign.current_status_msg = (
                        f"Antispam cooldown: {rem}s remaining..."
                    )
                time.sleep(1)
            with campaign.lock:
                campaign.delay_remaining = 0

    with campaign.lock:
        campaign.is_running = False
        campaign.current_status_msg = "Completed"
    campaign.log(
        f"Campaign batch finished! Sent: {campaign.success_count}, Failed: {campaign.failed_count}",
        "info",
    )


# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------
def _load_excel(file_path: Path):
    """Load an Excel file, store handler + clients in session_state."""
    handler = ExcelHandler(file_path)
    ok, msg = handler.load_file()
    if not ok:
        st.error(f"❌ {msg}")
        return

    clients = handler.get_clients()
    st.session_state["active_file"] = file_path
    st.session_state["clients"] = clients
    st.session_state["column_mapping"] = handler.get_column_mapping_summary()

    total = len(clients)
    valid = sum(1 for c in clients if is_valid_email(c.get("email", "")))
    sent = sum(
        1
        for c in clients
        if str(c.get("status", "")).strip().lower() == "sent"
    )
    pending = sum(
        1
        for c in clients
        if str(c.get("status", "")).strip().lower() != "sent"
        and is_valid_email(c.get("email", ""))
    )

    st.session_state["stats"] = {
        "total": total,
        "valid": valid,
        "sent": sent,
        "pending": pending,
    }
    st.success(f"✅ {msg}")


def start_campaign(
    file_path: Optional[Path], min_delay: int = 5, max_delay: int = 15
) -> Tuple[bool, str]:
    """Validates parameters, SMTP connection, and starts outreach campaign in background."""
    if not file_path or not Path(file_path).exists():
        return False, "No Excel file loaded. Please upload an Excel file first."

    # Upfront SMTP verification
    sender = EmailSender()
    conn_ok, conn_err = sender.test_connection()
    if not conn_ok:
        return False, f"SMTP Connection Failed: {conn_err}"

    can_send, count_today, lim = sender.can_send_today()
    if not can_send:
        return False, f"Daily limit reached ({count_today}/{lim} sent today). Cannot send more today."

    handler = ExcelHandler(file_path)
    ok, msg = handler.load_file()
    if not ok:
        return False, msg

    clients = handler.get_clients()
    pending = [
        c
        for c in clients
        if str(c.get("status", "")).strip().lower() != "sent"
        and is_valid_email(c.get("email", ""))
    ]
    if not pending:
        return False, "All leads in this file have already been contacted (no pending leads)."

    campaign = _get_campaign()
    min_d = max(1, int(min_delay))
    max_d = max(min_d, int(max_delay))
    campaign.reset(len(pending), Path(file_path))

    worker = threading.Thread(
        target=background_campaign_worker,
        args=(Path(file_path), min_d, max_d, campaign),
        daemon=True,
    )
    worker.start()
    return True, f"Outreach campaign launched for {len(pending)} pending leads!"


def render_campaign_monitor(campaign: CampaignController, key_prefix: str = "mon"):
    """Unified live campaign monitor for active progress, metrics, and logs."""
    status = campaign.get_status()
    if status["is_running"]:
        st.markdown(
            f"""
            <div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 18px 22px; border-radius: 12px; border: 1px solid #38bdf8; margin-bottom: 16px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size: 1.15rem; font-weight: 700; color: #38bdf8;">⚡ Live Outreach in Progress</span>
                    <span style="font-size: 0.95rem; color: #94a3b8; font-weight: 600;">{status['current_index']} / {status['total']} Leads</span>
                </div>
                <p style="margin: 8px 0 0 0; color: #f1f5f9; font-size: 0.98rem; font-weight: 500;">
                    {status['current_status_msg']}
                </p>
                <p style="margin: 4px 0 0 0; color: #94a3b8; font-size: 0.88rem;">
                    Target: <strong style="color: #67e8f9;">{status['current_lead']}</strong>
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        progress_val = (
            status["current_index"] / status["total"]
            if status["total"] > 0
            else 0.0
        )
        st.progress(progress_val, text=f"Campaign Progress: {status['current_index']}/{status['total']} Leads Processed")

        c1, c2, c3 = st.columns(3)
        c1.metric("✅ Successfully Sent", status["success_count"])
        c2.metric("❌ Failed Deliveries", status["failed_count"])
        c3.metric("⏳ Anti-Spam Cooldown", f"{status['delay_remaining']}s")

        if st.button("⛔ Stop Outreach Campaign", type="secondary", width="stretch", key=f"{key_prefix}_stop_btn"):
            with campaign.lock:
                campaign.stop_requested = True
            campaign.log("Stop requested by user.", "warning")
            st.warning("Halting outreach... completing current email.")

        with st.expander("📋 Live Activity & Delivery Logs", expanded=True):
            log_container = st.container(height=260)
            with log_container:
                for entry in reversed(status["logs"]):
                    lvl = entry.get("level", "info")
                    css_cls = f"log-{lvl}"
                    st.markdown(
                        f'<span class="{css_cls}">[{entry["time"]}] {entry["text"]}</span>',
                        unsafe_allow_html=True,
                    )

        time.sleep(1.5)
        st.rerun()

    elif status["total"] > 0 and status["current_status_msg"] == "Completed":
        # Synchronize Excel into session_state once completed
        active_f = st.session_state.get("active_file")
        sync_token = f"sync_{status['success_count']}_{status['failed_count']}"
        if active_f and Path(active_f).exists() and st.session_state.get("last_sync_token") != sync_token:
            _load_excel(Path(active_f))
            st.session_state["last_sync_token"] = sync_token

        st.success(
            f"🎉 **Campaign Complete!** Successfully delivered **{status['success_count']}** cold emails ({status['failed_count']} failed)."
        )


def _save_ai_settings(provider: str, nvidia_key: str, nvidia_model: str, anthropic_key: str):
    """Persist AI settings to config module and .env file."""
    if provider in ("nvidia", "anthropic", "template"):
        config.AI_PROVIDER = provider
    if nvidia_key is not None and not nvidia_key.startswith("your_"):
        config.NVIDIA_API_KEY = nvidia_key.strip()
    if nvidia_model:
        config.NVIDIA_MODEL = nvidia_model.strip()
    if anthropic_key is not None and not anthropic_key.startswith("your_"):
        config.ANTHROPIC_API_KEY = anthropic_key.strip()

    env_file = config.ENV_PATH
    if env_file.exists():
        try:
            content = env_file.read_text(encoding="utf-8")
            updates = {
                "AI_PROVIDER": config.AI_PROVIDER,
                "NVIDIA_API_KEY": config.NVIDIA_API_KEY,
                "NVIDIA_MODEL": config.NVIDIA_MODEL,
                "ANTHROPIC_API_KEY": config.ANTHROPIC_API_KEY,
            }
            for k, v in updates.items():
                pattern = rf"^{k}=.*$"
                if re.search(pattern, content, flags=re.M):
                    content = re.sub(pattern, f"{k}={v}", content, flags=re.M)
                else:
                    content += f"\n{k}={v}"
            env_file.write_text(content, encoding="utf-8")
        except Exception as e:
            logger.warning(f"Could not persist settings to .env: {e}")


# ═══════════════════════════════════════════════════════════════════════════
#  STREAMLIT PAGE CONFIG
# ═══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Skytecher Email Automation",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Custom CSS for Premium Look
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
/* ---- Global ---- */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* ---- Sidebar ---- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
}
section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stTextInput label { color: #94a3b8 !important; }

/* ---- Metric cards ---- */
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    border: 1px solid #475569;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.15);
}
div[data-testid="stMetric"] label { color: #94a3b8 !important; }
div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #f1f5f9 !important; font-weight: 700; }

/* ---- Tabs ---- */
.stTabs [data-baseweb="tab-list"] { gap: 8px; }
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    padding: 10px 24px;
    font-weight: 600;
}

/* ---- Log area ---- */
.log-success { color: #22c55e; font-weight: 500; }
.log-error   { color: #ef4444; font-weight: 500; }
.log-warning { color: #f59e0b; }
.log-info    { color: #94a3b8; }

/* ---- Header ---- */
.hero-header {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #0f172a 100%);
    border-radius: 16px;
    padding: 28px 36px;
    margin-bottom: 24px;
    border: 1px solid #334155;
}
.hero-header h1 { color: #f8fafc; margin: 0; font-size: 1.8rem; }
.hero-header p  { color: #94a3b8; margin: 4px 0 0 0; font-size: 0.95rem; }

/* ---- Dataframe ---- */
.stDataFrame { border-radius: 8px; overflow: hidden; }
</style>
""",
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════════
#  SIDEBAR — Configuration & AI Settings
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🚀 Skytecher")
    st.caption("Cold Email Automation System v2.0")
    st.divider()

    # ---- Connection status ----
    st.markdown("#### ⚙️ System Status")
    has_password = bool(
        config.GMAIL_APP_PASSWORD and not config.GMAIL_APP_PASSWORD.startswith("your_")
    )
    st.markdown(
        f"**SMTP:** `{config.SMTP_SERVER}:{config.SMTP_PORT}` — "
        + ("🟢 Password Set" if has_password else "🔴 Password Missing")
    )

    engine_label = "Smart Template (Offline)"
    if config.AI_PROVIDER == "nvidia" and config.NVIDIA_API_KEY and not config.NVIDIA_API_KEY.startswith("your_"):
        engine_label = f"NVIDIA NIM ({config.NVIDIA_MODEL})"
    elif config.AI_PROVIDER == "anthropic" and config.ANTHROPIC_API_KEY and not config.ANTHROPIC_API_KEY.startswith("your_"):
        engine_label = f"Claude ({config.ANTHROPIC_MODEL})"
    st.markdown(f"**AI Engine:** {engine_label}")

    sender_obj = EmailSender()
    can_send, sent_today, limit = sender_obj.can_send_today()
    st.markdown(f"**Quota:** {sent_today} / {limit} emails today")
    st.divider()

    # ---- AI Settings ----
    st.markdown("#### 🤖 AI Provider Settings")
    provider_options = ["nvidia", "anthropic", "template"]
    current_idx = provider_options.index(config.AI_PROVIDER) if config.AI_PROVIDER in provider_options else 0
    ai_provider = st.selectbox("Provider", provider_options, index=current_idx, key="sb_provider")

    if ai_provider == "nvidia":
        nvidia_key_input = st.text_input(
            "NVIDIA API Key",
            value=config.NVIDIA_API_KEY if config.NVIDIA_API_KEY and not config.NVIDIA_API_KEY.startswith("your_") else "",
            type="password",
            placeholder="nvapi-...",
            help="Get your API key at https://build.nvidia.com",
            key="sb_nvidia_key",
        )

        catalog = list(config.NVIDIA_MODELS_CATALOG)
        custom_tag = "✏️ Custom Model ID..."
        model_options = catalog + [custom_tag]

        if config.NVIDIA_MODEL in catalog:
            default_model_idx = catalog.index(config.NVIDIA_MODEL)
        else:
            default_model_idx = len(catalog)

        chosen_model_option = st.selectbox(
            "NVIDIA NIM Model",
            model_options,
            index=default_model_idx,
            help="Select any NIM model from build.nvidia.com or enter a custom one",
            key="sb_nvidia_model_select",
        )

        if chosen_model_option == custom_tag:
            nvidia_model_input = st.text_input(
                "Enter Model ID",
                value=config.NVIDIA_MODEL if config.NVIDIA_MODEL not in catalog else "",
                placeholder="e.g. openai/gpt-oss-20b",
                key="sb_nvidia_model_custom",
            )
        else:
            nvidia_model_input = chosen_model_option

        col_save, col_test = st.columns(2)
        with col_save:
            if st.button("💾 Save", width="stretch", key="btn_save_ai"):
                _save_ai_settings("nvidia", nvidia_key_input, nvidia_model_input, "")
                st.success("Saved!")
                st.rerun()

        with col_test:
            if st.button("⚡ Test", width="stretch", key="btn_test_nvidia_sb"):
                from email_writer import test_nvidia_connection
                with st.spinner("Testing NIM..."):
                    ok, msg = test_nvidia_connection(
                        nvidia_key_input,
                        nvidia_model_input,
                        config.NVIDIA_BASE_URL,
                    )
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)

    elif ai_provider == "anthropic":
        anthropic_key_input = st.text_input(
            "Anthropic API Key",
            value=config.ANTHROPIC_API_KEY if config.ANTHROPIC_API_KEY and not config.ANTHROPIC_API_KEY.startswith("your_") else "",
            type="password",
            key="sb_anthropic_key",
        )
        anthropic_model_input = st.text_input(
            "Claude Model", value=config.ANTHROPIC_MODEL, key="sb_anthropic_model"
        )
        if st.button("💾 Save AI Settings", width="stretch", key="btn_save_anthropic"):
            _save_ai_settings("anthropic", "", "", anthropic_key_input)
            st.success("AI settings saved!")
            st.rerun()

    else:
        st.info("ℹ️ Using Smart Template Engine (100% offline, zero API costs).")
        if st.button("💾 Save as Default", width="stretch", key="btn_save_tpl"):
            _save_ai_settings("template", "", "", "")
            st.success("Saved!")
            st.rerun()

    st.divider()

    # ---- Services ----
    st.markdown("#### 🛠️ Company Services")
    for svc in config.COMPANY_SERVICES:
        st.markdown(f"- {svc}")


# ═══════════════════════════════════════════════════════════════════════════
#  HERO HEADER
# ═══════════════════════════════════════════════════════════════════════════
st.markdown(
    f"""
<div class="hero-header">
    <h1>🚀 Skytecher Cold Email Automation</h1>
    <p>AI-Powered B2B Personalization &amp; High-Deliverability Outreach &nbsp;|&nbsp;
       Sender: <strong>{config.SENDER_EMAIL}</strong> &nbsp;|&nbsp;
       <a href="https://{config.COMPANY_WEBSITE}" target="_blank" style="color:#60a5fa;">{config.COMPANY_WEBSITE}</a>
    </p>
</div>
""",
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════════
#  TABS
# ═══════════════════════════════════════════════════════════════════════════
tab_upload, tab_preview, tab_campaign, tab_diagnostics = st.tabs(
    ["📁 Upload & Data", "✉️ Preview & Test", "🚀 Live Campaign", "🩺 Diagnostics"]
)


# ═══════════════════════════════════════════════════════════════════════════
#  TAB 1 — Upload & Data
# ═══════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.subheader("Upload Client Excel File")

    # ---- Template Download ----
    st.markdown(
        """
> **📋 Required Excel Format:**
> `Company` · `Category` · `Country` · `City` · `Email(s)` · `Website` · `Has Website?` · `Lead Type` · `Suggestion` · `LinkedIn` · `Instagram` · `Maps Link`
>
> Download the template below, fill it in, and upload it.
"""
    )

    template_path = BASE_DIR / "upload_template.xlsx"
    if not template_path.exists():
        ExcelHandler.generate_template_excel(template_path)

    with open(template_path, "rb") as tpl_file:
        st.download_button(
            "📥 Download Excel Template (.xlsx)",
            data=tpl_file,
            file_name="Skytecher_Upload_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            key="btn_download_template",
        )

    st.divider()

    campaign = _get_campaign()
    camp_status = campaign.get_status()

    col_up, col_sample = st.columns([3, 1])
    newly_uploaded = False
    with col_up:
        uploaded = st.file_uploader(
            "Drag & drop your .xlsx / .xls file",
            type=["xlsx", "xls", "xlsm"],
            key="file_uploader",
        )
        if uploaded is not None:
            upload_id = f"{uploaded.name}_{uploaded.size}"
            if st.session_state.get("last_uploaded_id") != upload_id:
                filename = uploaded.name.replace(" ", "_")
                if not filename:
                    filename = "uploaded_clients.xlsx"
                save_path = UPLOAD_DIR / filename
                save_path.write_bytes(uploaded.getvalue())
                st.session_state["last_uploaded_id"] = upload_id
                _load_excel(save_path)
                newly_uploaded = True

    with col_sample:
        st.markdown("&nbsp;")  # spacer
        if st.button("📂 Load Sample Data", width="stretch", key="btn_load_sample"):
            sample_path = config.DEFAULT_INPUT_FILE
            ExcelHandler.generate_sample_excel(sample_path)
            st.session_state["last_uploaded_id"] = "sample_data_fixed"
            _load_excel(sample_path)
            st.rerun()

    # ---- Stats Cards ----
    if "stats" in st.session_state:
        stats = st.session_state["stats"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Rows", stats["total"])
        c2.metric("Valid Emails", stats["valid"])
        c3.metric("Already Sent", stats["sent"])
        c4.metric("Pending Leads", stats["pending"])

        # Auto-launch if enabled
        if newly_uploaded and st.session_state.get("auto_send_on_upload") and stats["pending"] > 0:
            active_file = st.session_state.get("active_file")
            min_d = st.session_state.get("t1_min_delay", config.MIN_DELAY_SECONDS)
            max_d = st.session_state.get("t1_max_delay", config.MAX_DELAY_SECONDS)
            ok, msg = start_campaign(active_file, min_d, max_d)
            if ok:
                st.success(f"⚡ Auto-Launch: {msg}")
                st.rerun()

    # ---- Live Campaign Monitor (If outreach is active) ----
    if camp_status["is_running"]:
        st.markdown("---")
        render_campaign_monitor(campaign, key_prefix="t1_active")
    else:
        # ---- Outreach Action Center (Direct Send Button) ----
        if "stats" in st.session_state and st.session_state["stats"]["pending"] > 0:
            st.markdown(
                """
                <div style="background: linear-gradient(135deg, #1e3a5f 0%, #0f172a 100%); border: 1px solid #38bdf8; border-radius: 12px; padding: 20px 24px; margin: 18px 0 10px 0;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h3 style="color: #f8fafc; margin: 0; font-size: 1.25rem;">🚀 Ready to Launch Cold Outreach</h3>
                        <span style="background: #0284c7; color: white; padding: 4px 12px; border-radius: 9999px; font-weight: 600; font-size: 0.85rem;">Ready to Send</span>
                    </div>
                    <p style="color: #cbd5e1; margin: 8px 0 0 0; font-size: 0.95rem;">
                        Excel data loaded successfully. Click the button below to generate AI-personalized emails and deliver them via SMTP.
                    </p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            col_btn, col_d1, col_d2 = st.columns([3, 1, 1])
            with col_btn:
                pending_count = st.session_state["stats"]["pending"]
                if st.button(
                    f"🚀 START OUTREACH: GENERATE & SEND TO ALL ({pending_count} LEADS)",
                    type="primary",
                    width="stretch",
                    key="btn_launch_from_tab1",
                ):
                    active_f = st.session_state.get("active_file")
                    min_d = st.session_state.get("t1_min_delay", 5)
                    max_d = st.session_state.get("t1_max_delay", 15)
                    ok, msg = start_campaign(active_f, min_d, max_d)
                    if ok:
                        st.success(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")

            with col_d1:
                st.number_input("Min Delay (s)", value=5, min_value=1, max_value=60, key="t1_min_delay")
            with col_d2:
                st.number_input("Max Delay (s)", value=15, min_value=2, max_value=120, key="t1_max_delay")

            st.checkbox(
                "⚡ Automatically start sending cold emails immediately whenever a new Excel file is uploaded",
                value=st.session_state.get("auto_send_on_upload", False),
                key="auto_send_on_upload",
            )
        elif "stats" in st.session_state and st.session_state["stats"]["total"] > 0:
            st.success("🎉 All leads in this file have been contacted! (0 pending leads).")

    if "column_mapping" in st.session_state:
        with st.expander("🗂️ Column Mapping", expanded=False):
            mapping = st.session_state["column_mapping"]
            for internal, actual in mapping.items():
                icon = "✅" if actual != "Not Found" else "⚠️"
                st.markdown(f"{icon} **{internal}** → `{actual}`")

    if "clients" in st.session_state and st.session_state["clients"]:
        import pandas as pd

        display_keys = [
            "company_name",
            "email",
            "industry",
            "city",
            "website",
            "lead_type",
            "suggested_service",
            "status",
        ]
        display_labels = {
            "company_name": "Company",
            "email": "Email",
            "industry": "Category",
            "city": "Location",
            "website": "Website",
            "lead_type": "Lead Type",
            "suggested_service": "Suggestion",
            "status": "Status",
        }
        rows = [
            {display_labels.get(k, k): c.get(k, "") for k in display_keys}
            for c in st.session_state["clients"]
        ]
        st.dataframe(pd.DataFrame(rows), width="stretch", height=380)


# ═══════════════════════════════════════════════════════════════════════════
#  TAB 2 — Preview & Test Send
# ═══════════════════════════════════════════════════════════════════════════
with tab_preview:
    if "clients" not in st.session_state or not st.session_state["clients"]:
        st.info("📤 Please upload an Excel file first in the **Upload & Data** tab.")
    else:
        clients = st.session_state["clients"]
        pending_clients = [
            c
            for c in clients
            if str(c.get("status", "")).strip().lower() != "sent"
            and is_valid_email(c.get("email", ""))
        ]

        st.subheader("Single Email Preview")
        if not pending_clients:
            st.warning("No pending leads to preview.")
        else:
            options = [
                f"{c.get('company_name', 'N/A')} — {c.get('email', '')}"
                for c in pending_clients
            ]
            col_lead, col_model = st.columns([3, 2])
            with col_lead:
                selected_idx = st.selectbox(
                    "Select a lead to preview",
                    range(len(options)),
                    format_func=lambda i: options[i],
                    key="sb_lead_select",
                )
            with col_model:
                preview_models = [config.NVIDIA_MODEL] + [
                    m for m in config.NVIDIA_MODELS_CATALOG if m != config.NVIDIA_MODEL
                ]
                preview_model = st.selectbox(
                    "AI Model for preview",
                    preview_models,
                    key="sb_preview_model",
                    help="Select which NVIDIA NIM model should write this preview",
                )
            selected_client = pending_clients[selected_idx]

            if st.button("🔍 Generate Email Preview", key="btn_preview_single"):
                with st.spinner(f"Generating personalized email with {preview_model}..."):
                    writer = EmailWriter()
                    try:
                        gen = writer.generate_email(selected_client, model_override=preview_model)
                        st.session_state["preview_result"] = gen
                    except Exception as e:
                        st.error(f"Generation failed: {e}")

            if "preview_result" in st.session_state:
                gen = st.session_state["preview_result"]
                st.markdown(f"**Subject:** {gen['subject']}")
                st.markdown(f"**Service:** {gen['service']}  |  **Engine:** {gen['source']}")
                with st.expander("📝 Plain Text Body", expanded=True):
                    st.text(gen["plain_body"])
                with st.expander("🌐 HTML Preview", expanded=False):
                    st.components.v1.html(gen["html_body"], height=500, scrolling=True)

                col_snd_client, col_snd_spacer = st.columns([2, 1])
                with col_snd_client:
                    client_email_target = selected_client.get("email", "").strip()
                    btn_label = f"✉️ Send This Email Directly to {selected_client.get('company_name', 'Client')} ({client_email_target})"
                    if st.button(btn_label, type="primary", width="stretch", key="btn_send_single_client"):
                        if not is_valid_email(client_email_target):
                            st.error(f"Invalid recipient email: '{client_email_target}'")
                        else:
                            with st.spinner(f"Sending cold email to {client_email_target}..."):
                                sender = EmailSender()
                                ok, err = sender.send_email(
                                    recipient_email=client_email_target,
                                    subject=gen["subject"],
                                    plain_body=gen["plain_body"],
                                    html_body=gen["html_body"],
                                    recipient_name=selected_client.get("client_name") or selected_client.get("company_name"),
                                )
                                if ok:
                                    active_f = st.session_state.get("active_file")
                                    if active_f and Path(active_f).exists():
                                        handler = ExcelHandler(active_f)
                                        handler.load_file()
                                        handler.update_row_status(selected_client["_row_index"], status="Sent", error_msg="")
                                        _load_excel(Path(active_f))
                                    st.success(f"🎉 Email successfully sent to **{client_email_target}**!")
                                    st.rerun()
                                else:
                                    st.error(f"❌ Sending failed: {err}")

        st.divider()

        # ---- Test Send ----
        st.subheader("📧 Test Send")
        test_email = st.text_input(
            "Test recipient email",
            value=config.TEST_RECIPIENT_EMAIL or config.SENDER_EMAIL,
            key="test_email_input",
        )

        if st.button("📤 Send Test Email", key="btn_test_send"):
            if not is_valid_email(test_email):
                st.error(f"Invalid email: {test_email}")
            elif not pending_clients:
                st.error("No pending client to generate a test for.")
            else:
                with st.spinner("Generating & sending test email..."):
                    writer = EmailWriter()
                    try:
                        gen = writer.generate_email(pending_clients[0])
                    except Exception as e:
                        st.error(f"Generation failed: {e}")
                        gen = None

                    if gen:
                        sender = EmailSender()
                        ok, err = sender.send_email(
                            recipient_email=test_email,
                            subject=f"[TEST] {gen['subject']}",
                            plain_body=gen["plain_body"],
                            html_body=gen["html_body"],
                            recipient_name="Tester",
                        )
                        if ok:
                            st.success(f"✅ Test email sent to **{test_email}**!")
                        else:
                            st.error(f"❌ Send failed: {err}")

        st.divider()

        # ---- Dry Run (Batch Preview) ----
        st.subheader("📋 Dry Run — Batch Preview")
        if st.button("▶️ Generate All Previews (Dry Run)", key="btn_dry_run"):
            active_file = st.session_state.get("active_file")
            if not active_file or not Path(active_file).exists():
                st.error("No Excel file loaded.")
            else:
                handler = ExcelHandler(active_file)
                ok, msg = handler.load_file()
                if not ok:
                    st.error(msg)
                else:
                    all_clients = handler.get_clients()
                    writer = EmailWriter()
                    previews = []
                    progress_bar = st.progress(0, text="Generating emails...")
                    total_pending = len([
                        c for c in all_clients
                        if str(c.get("status", "")).strip().lower() != "sent"
                        and is_valid_email(c.get("email", ""))
                    ])
                    done = 0

                    for c in all_clients:
                        if str(c.get("status", "")).strip().lower() == "sent":
                            continue
                        if not is_valid_email(c.get("email", "")):
                            continue
                        try:
                            gen = writer.generate_email(c)
                            previews.append({
                                "Client": c.get("client_name", ""),
                                "Company": c.get("company_name", ""),
                                "Email": c.get("email", ""),
                                "Service": gen["service"],
                                "Engine": gen["source"],
                                "Subject": gen["subject"],
                            })
                        except Exception as e:
                            logger.error(f"Preview error for {c.get('company_name')}: {e}")
                        done += 1
                        if total_pending > 0:
                            progress_bar.progress(
                                done / total_pending,
                                text=f"Generated {done}/{total_pending}...",
                            )

                    progress_bar.progress(1.0, text="Done!")
                    if previews:
                        import pandas as pd

                        st.dataframe(pd.DataFrame(previews), width="stretch")
                        handler.export_preview(previews, config.PREVIEW_FILE)
                        st.success(
                            f"✅ Generated {len(previews)} email previews. Exported to `{config.PREVIEW_FILE.name}`."
                        )
                    else:
                        st.warning("No pending emails to preview.")


# ═══════════════════════════════════════════════════════════════════════════
#  TAB 3 — Live Campaign
# ═══════════════════════════════════════════════════════════════════════════
with tab_campaign:
    campaign = _get_campaign()
    status = campaign.get_status()

    st.subheader("🚀 Live Email Campaign")

    if not status["is_running"]:
        # ---- Launch controls ----
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            min_delay = st.number_input(
                "Min Delay (sec)", value=config.MIN_DELAY_SECONDS, min_value=1, key="camp_min_delay"
            )
        with col_d2:
            max_delay = st.number_input(
                "Max Delay (sec)", value=config.MAX_DELAY_SECONDS, min_value=2, key="camp_max_delay"
            )

        if st.button("🚀 Launch Campaign", type="primary", width="stretch", key="btn_launch"):
            active_file = st.session_state.get("active_file")
            ok, msg = start_campaign(active_file, min_delay, max_delay)
            if ok:
                st.success(f"🚀 {msg}")
                st.rerun()
            else:
                st.error(f"❌ {msg}")

        # Render completion summary if campaign just completed
        render_campaign_monitor(campaign, key_prefix="tab3_done")

    else:
        # ---- Live progress ----
        render_campaign_monitor(campaign, key_prefix="tab3_active")

    # ---- Download buttons ----
    st.divider()
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        active_file = st.session_state.get("active_file")
        if active_file and Path(active_file).exists():
            with open(active_file, "rb") as f:
                st.download_button(
                    "📥 Download Updated Excel",
                    data=f,
                    file_name=Path(active_file).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="btn_dl_active_excel",
                )
    with col_dl2:
        if config.PREVIEW_FILE.exists():
            with open(config.PREVIEW_FILE, "rb") as f:
                st.download_button(
                    "📥 Download Preview Excel",
                    data=f,
                    file_name="Skytecher_Email_Preview.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    key="btn_dl_preview_excel",
                )


# ═══════════════════════════════════════════════════════════════════════════
#  TAB 4 — Diagnostics
# ═══════════════════════════════════════════════════════════════════════════
with tab_diagnostics:
    st.subheader("🩺 System Diagnostics")

    if st.button("🔄 Run Full Diagnostics", key="btn_diag"):
        with st.spinner("Running diagnostics..."):
            results = {}

            # SMTP
            sender = EmailSender()
            smtp_ok, smtp_msg = sender.test_connection()
            results["smtp"] = {"ok": smtp_ok, "msg": smtp_msg}

            # NVIDIA
            from email_writer import test_nvidia_connection
            nvidia_ok, nvidia_msg = test_nvidia_connection(
                config.NVIDIA_API_KEY, config.NVIDIA_MODEL, config.NVIDIA_BASE_URL
            )
            results["nvidia"] = {"ok": nvidia_ok, "msg": nvidia_msg}

            # Anthropic
            ai_ok = False
            ai_msg = "Anthropic API Key not set."
            if config.ANTHROPIC_API_KEY and not config.ANTHROPIC_API_KEY.startswith("your_"):
                try:
                    import anthropic

                    ai_client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
                    resp = ai_client.messages.create(
                        model=config.ANTHROPIC_MODEL,
                        max_tokens=15,
                        messages=[{"role": "user", "content": "Ping!"}],
                    )
                    ai_ok = True
                    ai_msg = f"Claude Active! Model: {config.ANTHROPIC_MODEL}"
                except Exception as e:
                    ai_msg = f"Claude Error: {e}"
            results["anthropic"] = {"ok": ai_ok, "msg": ai_msg}

            st.session_state["diag_results"] = results

    if "diag_results" in st.session_state:
        res = st.session_state["diag_results"]

        for label, key in [
            ("📧 Gmail SMTP", "smtp"),
            ("🟢 NVIDIA NIM API", "nvidia"),
            ("🟣 Anthropic Claude", "anthropic"),
        ]:
            data = res[key]
            icon = "✅" if data["ok"] else "❌"
            st.markdown(f"### {label} {icon}")
            if data["ok"]:
                st.success(data["msg"])
            else:
                st.error(data["msg"])

        st.markdown(f"**Active Provider:** `{config.AI_PROVIDER}`")


# ═══════════════════════════════════════════════════════════════════════════
#  FOOTER
# ═══════════════════════════════════════════════════════════════════════════
st.divider()
st.caption(
    f"© 2026 Skytecher • skytecher.com • Engine: {config.AI_PROVIDER.upper()} • "
    f"Quota: {sent_today}/{limit} today"
)

