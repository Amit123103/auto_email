# Skytecher Cold Email Automation System

A production-ready, AI-driven B2B cold email automation engine built in Python. Reads prospective client datasets directly from Excel, intelligently infers or matches the ideal Skytecher service offering, authors hyper-personalized outreach emails using **NVIDIA NIM Models** (e.g. LLaMA 3.1 70B, Nemotron, Mixtral) or **Anthropic Claude**, with an intelligent multi-variation template fallback, and delivers them via **Gmail SMTP** with real-time web dashboard monitoring and Excel status writeback.

---

## 🌟 Key Features

- **Web Dashboard & Real-Time Monitoring**: Interactive graphical UI launched automatically with `python main.py` at `http://127.0.0.1:5000`.
- **NVIDIA NIM & Claude AI Integration**: Support for cutting-edge NVIDIA AI Foundation Models (`meta/llama-3.1-70b-instruct`, `meta/llama-3.3-70b-instruct`, `nvidia/llama-3.1-nemotron-70b-instruct`) and Anthropic Claude.
- **Live AI Settings Switcher**: Switch providers or update your API keys directly from the Web UI without restarting the application.
- **Fail-Safe Template Fallback**: If API keys are unset or network issues occur, seamlessly uses rotating high-converting B2B templates without stopping.
- **Smart Column Detection**: Auto-detects columns regardless of variations (e.g. `Email ID`, `e-mail`, `Company Name`, `Firm`, `Suggested Service`, `Service`).
- **Auto Service Matching**: Automatically analyzes company descriptions, websites, and industries to select the best-matching Skytecher offering.
- **Gmail SMTP Delivery**: Sends authenticated emails via `smtp.gmail.com` using a secure 16-character Google App Password (port 587 TLS / 465 SSL).
- **Anti-Spam Throttling**: Random delays (30–90 seconds) between emails and configurable daily limits protect your sender score.
- **Live Excel Writeback**: Updates `Status` (`Sent` / `Failed` / `Skipped`), `Sent Date & Time`, and `Error Message` directly in your spreadsheet.
- **Three Operational Modes**:
  1. **DRY-RUN Mode**: Generates emails, previews them in terminal / browser, and exports them to `preview.xlsx`.
  2. **TEST Mode**: Generates an email and sends it only to your personal inbox to verify formatting and deliverability.
  3. **SEND Mode**: Executes the live outreach campaign with automated retries and delays.

---

## 🛠️ Step-by-Step Setup Guide

### 1. Prerequisites
- **Python 3.10+** installed on Windows, macOS, or Linux.
- A **Gmail or Google Workspace** account.
- *(Recommended)* An **NVIDIA NIM API Key** (free with 1,000 credits at [build.nvidia.com](https://build.nvidia.com/)).
- *(Optional)* An **Anthropic Claude API Key** from [console.anthropic.com](https://console.anthropic.com/).

---

### 2. Generate a Gmail App Password

> **Important**: You cannot use your regular Gmail password. Google requires an **App Password** for SMTP connections.

1. Go to your **Google Account**: [myaccount.google.com](https://myaccount.google.com/).
2. In the left navigation, click **Security**.
3. Under *"How you sign in to Google"*, verify that **2-Step Verification** is turned **ON**.
4. In the search bar at the top of your Google Account, search for **"App passwords"** (or visit directly: [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)).
5. Under **App name**, type `Skytecher Outreach` and click **Create**.
6. Google will display a **16-character code** (e.g. `abcd efgh ijkl mnop`).
7. Copy this 16-character code into your `.env` file as `GMAIL_APP_PASSWORD`.

---

### 3. Configure Your Environment (`.env`)

Open `.env` in any editor or configure it directly in the Web UI:

```ini
# --- Gmail SMTP Configuration ---
GMAIL_ADDRESS=skytechersolutions@gmail.com
GMAIL_APP_PASSWORD=abcdefghijklmnop
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
USE_SSL=False

# --- AI Provider Configuration ('nvidia', 'anthropic', or 'template') ---
AI_PROVIDER=nvidia
USE_AI=True

# --- NVIDIA NIM API (Free credits at https://build.nvidia.com/) ---
NVIDIA_API_KEY=nvapi-your_nvidia_api_key_here
NVIDIA_BASE_URL=https://integrate.api.nvidia.com/v1
NVIDIA_MODEL=meta/llama-3.1-70b-instruct

# --- Anthropic Claude API (Optional) ---
ANTHROPIC_API_KEY=sk-ant-api03-...
ANTHROPIC_MODEL=claude-3-5-sonnet-20241022

# --- Anti-Spam Delays & Quotas ---
MIN_DELAY_SECONDS=30
MAX_DELAY_SECONDS=90
DAILY_SEND_LIMIT=100
EMAIL_LANGUAGE=English
TEST_RECIPIENT_EMAIL=skytechersolutions@gmail.com
```

---

### 4. Install Dependencies

Install the required Python libraries using pip:

```powershell
pip install -r requirements.txt
```

---

## 🚀 Running the Application

Launch the application:

```powershell
python main.py
```

Running `python main.py` starts the local web server and **automatically opens the Skytecher Web Dashboard in your browser** at `http://127.0.0.1:5000`.

### What you can do in the Web Dashboard:
1. **Upload Excel Spreadsheet**: Drag & drop any `.xlsx` or `.xls` file. The system auto-detects columns and displays all client rows with validation badges.
2. **Quick Load Sample Dataset**: Click **"Load Sample Dataset"** to immediately test with 10 pre-loaded diverse B2B leads.
3. **Interactive Leads Table**: Search, filter, and click **"Preview Email"** on any client to see the personalized email rendered in full HTML and plain text.
4. **Dry-Run Preview**: Generate emails for all eligible leads and download the generated `preview.xlsx` file.
5. **Test Send**: Send a real test email to your own inbox to inspect Gmail deliverability and layout.
6. **Live Outreach Campaign**: Click **"Start Live Campaign"** with real-time progress bar, anti-spam delays (30–90s countdown), and live terminal logs.
7. **Download Updated Excel**: Download the spreadsheet anytime with real-time `Sent` / `Failed` statuses and timestamps recorded.

*(Tip: If you ever prefer the terminal menu, you can run `python main.py --cli`)*

---

## 📊 Modes of Operation

### 1. DRY-RUN Mode (Safe Preview)
- Reads the Excel file, verifies email validity, skips leads already marked as `Sent`.
- Personalizes every email and displays full plain-text previews in the terminal.
- Exports all generated emails, subjects, and HTML code directly to `preview.xlsx`.
- **Zero emails are sent.**

### 2. TEST Mode (Deliverability Check)
- Lets you pick any row from the Excel file.
- Writes a personalized email for that company.
- Sends it **only** to your own test email address (`TEST_RECIPIENT_EMAIL`).
- Allows you to inspect inbox placement and mobile formatting before bulk sending.

### 3. SEND Mode (Live Campaign)
- Performs pre-flight checks on SMTP credentials and daily limits.
- Confirms campaign parameters with you before starting.
- Automatically skips invalid emails and previously sent leads.
- Employs human-like random delays (30–90 seconds) between sends.
- Immediately records results (`Sent`, `Failed`, or `Skipped`) and timestamps back into the Excel file.

### 4. System Diagnostics
- Instantly tests your Gmail SMTP connection and authenticates your App Password.
- Pings the Anthropic Claude API to confirm your API key and token balance.

---

## 📋 Excel File Format & Column Mapping

The system includes fuzzy column matching, so column headers are flexible:

| Field | Supported Header Aliases | Example Value |
| :--- | :--- | :--- |
| **Client Name** | `Client Name`, `Name`, `Contact Person`, `Contact Name` | Dr. Sarah Jenkins |
| **Company Name** | `Company Name`, `Company`, `Organization`, `Business Name` | Apex Dental Clinic |
| **Email** | `Email`, `Email ID`, `E-mail`, `Work Email` | sjenkins@apexdentalcare.org |
| **Industry** | `Industry`, `Sector`, `Business Sector` | Healthcare |
| **Website** | `Website`, `URL`, `Site`, `Company URL` | https://apexdentalcare.org |
| **City** | `City`, `Location`, `State` | Austin |
| **Company Details** | `Company Details / Description`, `Description`, `About` | Private cosmetic dental clinic... |
| **Suggested Service** | `Suggested Service`, `Service`, `Offering` *(Can be blank)* | Website Design & Development |
| **Status** | `Status`, `Email Status` *(Managed by system)* | `Sent`, `Failed`, or blank |
| **Sent Date & Time** | `Sent Date & Time` *(Managed by system)* | 2026-09-30 22:05:00 |
| **Error Message** | `Error Message` *(Managed by system)* | Blank or failure reason |

---

## ✉️ 3 Sample Generated Cold Emails

Below are three examples of emails produced by the system across different industries:

### Sample 1: Healthcare / Dental Clinic
**Pitched Service**: Website Design & Development  
**Subject**: Quick question regarding Apex Dental Clinic's digital roadmap

> Dear Dr. Sarah Jenkins,
>
> I came across Apex Dental Clinic based in Austin and was very impressed by your dedication to patient care and cosmetic dentistry.
>
> In today's digital landscape, prospective patients expect a high-speed, mobile-first experience to evaluate treatments and schedule appointments seamlessly.
>
> Here is how Skytecher's Website Design & Development can support Apex Dental Clinic:
> - Modern, mobile-first responsive architecture built for maximum patient conversion
> - Blazing-fast load speeds, intuitive navigation, and direct online booking integration
> - Built-in local healthcare SEO foundations to dominate regional Austin searches
>
> At Skytecher, we've helped 50+ growing businesses scale their digital presence with agile, modern solutions.
>
> Would you be open to a brief 15-minute call this Thursday to share ideas, or simply reply to this email to see if there's a fit?
>
> Best regards,  
> **Skytecher Team**  
> [skytecher.com](https://skytecher.com) | skytechersolutions@gmail.com  
> 
> ---  
> *If you'd prefer not to receive emails from us, just reply 'Unsubscribe'.*

---

### Sample 2: E-Commerce & Retail (Streetwear Brand)
**Pitched Service**: Digital Marketing & SEO (Auto-matched from description)  
**Subject**: Scaling UrbanVibe Apparel's digital presence

> Dear Marcus Vance,
>
> While researching innovative teams in the e-commerce space, I noticed the strong community momentum UrbanVibe Apparel has been building around sustainable streetwear.
>
> With customer acquisition costs rising across paid channels, capturing high-intent organic search traffic and optimizing checkout conversions is the most sustainable way to drive repeat orders.
>
> Here is how Skytecher's Digital Marketing & SEO can help UrbanVibe Apparel:
> - High-intent e-commerce keyword optimization to outrank competing apparel brands
> - Conversion rate audit to reduce shopping cart abandonment and streamline checkout
> - Transparent weekly ROI reporting with actionable customer acquisition analytics
>
> Our team at Skytecher has spent the last 5 years helping modern consumer brands turn organic traffic into a dependable revenue engine.
>
> Are you free for a quick 15-minute introductory call next week, or feel free to reply directly here if you'd like more details.
>
> Best regards,  
> **Skytecher Team**  
> [skytecher.com](https://skytecher.com) | skytechersolutions@gmail.com  
> 
> ---  
> *If you'd prefer not to receive emails from us, just reply 'Unsubscribe'.*

---

### Sample 3: FinTech / SaaS Platform
**Pitched Service**: Mobile App Development  
**Subject**: Idea for PaySwift Financial | Mobile App Development

> Dear Elena Rostova,
>
> I've been following PaySwift Financial's footprint in the cross-border payments sector and wanted to reach out regarding your expansion initiatives.
>
> As global mobile transactions continue to outpace desktop, delivering a high-security, native mobile experience is critical for retaining remote teams and high-volume transactors.
>
> Here is how Skytecher's Mobile App Development can accelerate PaySwift Financial:
> - Intuitive iOS and Android applications built on high-performance native architectures
> - Bank-grade biometric authentication and real-time transaction push notifications
> - Rapid App Store and Google Play deployment with dedicated release engineering
>
> At Skytecher, we specialize in end-to-end digital engineering, partnering with venture-backed tech platforms to deliver scalable, zero-downtime products.
>
> Would you be against a 15-minute chat this week to explore ideas for PaySwift Financial? Alternatively, let me know if now isn't the right time.
>
> Best regards,  
> **Skytecher Team**  
> [skytecher.com](https://skytecher.com) | skytechersolutions@gmail.com  
> 
> ---  
> *If you'd prefer not to receive emails from us, just reply 'Unsubscribe'.*

---

## 🛡️ Deliverability & Anti-Spam Best Practices

1. **Warm Up Your Email**: If your Gmail address is new, start by sending 15–25 emails per day for the first two weeks before increasing to 50–100/day.
2. **Never Disable Delays in Live Campaigns**: Keep `MIN_DELAY_SECONDS=30` and `MAX_DELAY_SECONDS=90`. Sending emails in rapid bursts is the #1 trigger for Gmail automated spam filters.
3. **Respect Opt-Outs**: The included footer allows recipients to reply "Unsubscribe". If someone replies, remove them from your active lead sheets.
4. **Clean Your Lists**: The system automatically rejects malformed email strings, but you should regularly verify that your contact domains are active.

---

## 📄 License
Internal use for Skytecher B2B outreach campaigns.
