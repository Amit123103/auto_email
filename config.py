"""
Skytecher Cold Email Automation - Configuration Module
======================================================
Contains all environment settings, company profile, service offerings,
sending thresholds, and copywriting templates.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(dotenv_path=ENV_PATH)

# ==============================================================================
# SENDER & COMPANY PROFILE
# ==============================================================================
COMPANY_NAME = "Skytecher"
COMPANY_WEBSITE = "skytecher.com"
COMPANY_PHONE = os.getenv("COMPANY_PHONE", "+91-8960061745")
SENDER_NAME = "Skytecher Team"
SENDER_EMAIL = os.getenv("GMAIL_ADDRESS", "skytechersolutions@gmail.com")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "").strip()

# Gmail SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
USE_SSL = os.getenv("USE_SSL", "False").lower() in ("true", "1", "yes")

# ==============================================================================
# AI PROVIDER CONFIGURATION (NVIDIA NIM / ANTHROPIC CLAUDE)
# ==============================================================================
# Primary provider: 'nvidia', 'anthropic', or 'template'
AI_PROVIDER = os.getenv("AI_PROVIDER", "nvidia").lower()
USE_AI = os.getenv("USE_AI", "True").lower() in ("true", "1", "yes")

# NVIDIA NIM API Configuration (https://build.nvidia.com/)
NVIDIA_API_KEY = os.getenv("NVIDIA_API_KEY", "").strip()
NVIDIA_BASE_URL = os.getenv("NVIDIA_BASE_URL", "https://integrate.api.nvidia.com/v1").strip()
NVIDIA_MODEL = os.getenv("NVIDIA_MODEL", "openai/gpt-oss-20b").strip()

# Full catalog of supported NVIDIA NIM text/chat models
NVIDIA_MODELS_CATALOG = [
    "openai/gpt-oss-20b",
    "deepseek-ai/deepseek-v4.1-flash",
    "deepseek-ai/deepseek-r1",
    "deepseek-ai/deepseek-v3",
    "google/gemma-4-31b-it",
    "google/gemma-2-27b-it",
    "google/gemma-2-9b-it",
    "google/diffusiongemma-26b-a4b-it",
    "meta/llama-3.3-70b-instruct",
    "meta/llama-3.1-70b-instruct",
    "meta/llama-3.1-8b-instruct",
    "meta/llama-3.2-11b-vision-instruct",
    "meta/llama-3.2-90b-vision-instruct",
    "meta/muse-glimmer-30b",
    "qwen/qwen2.5-72b-instruct",
    "nvidia/llama-3.1-nemotron-70b-instruct",
    "nvidia/nemotron-4-340b-instruct",
    "mistralai/mixtral-8x22b-instruct-v0.1",
]

# Anthropic Claude API Configuration
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "").strip()
ANTHROPIC_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-20241022").strip()

# ==============================================================================
# CAMPAIGN & SENDING LIMITS
# ==============================================================================
EMAIL_LANGUAGE = os.getenv("EMAIL_LANGUAGE", "English")
TEST_RECIPIENT_EMAIL = os.getenv("TEST_RECIPIENT_EMAIL", SENDER_EMAIL)

# Anti-spam delay between emails (in seconds)
MIN_DELAY_SECONDS = int(os.getenv("MIN_DELAY_SECONDS", "30"))
MAX_DELAY_SECONDS = int(os.getenv("MAX_DELAY_SECONDS", "90"))

# Daily sending quota (Gmail free limits ~500/day, safe cold outreach is 50-100/day)
DAILY_SEND_LIMIT = int(os.getenv("DAILY_SEND_LIMIT", "100"))

# Retry settings for failed deliveries
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 5

# File paths
LOG_FILE = BASE_DIR / "email_log.txt"
DAILY_STATS_FILE = BASE_DIR / "daily_stats.json"
DEFAULT_INPUT_FILE = BASE_DIR / "sample_clients.xlsx"
PREVIEW_FILE = BASE_DIR / "preview.xlsx"

# ==============================================================================
# SKYTECHER SERVICES & VALUE PROPOSITIONS
# ==============================================================================
# Edit or add services here. These are used in auto-matching and AI prompts.
COMPANY_SERVICES = [
    "Website Design & Development",
    "Mobile App Development",
    "Software / Custom Development",
    "Digital Marketing & SEO",
    "Branding & UI/UX Design",
    "IT Consulting & Support",
]

# Service details for smart copywriting and pitch personalization
SERVICE_DETAILS = {
    "Website Design & Development": {
        "focus": "High-converting, ultra-fast websites & web platforms tailored for business growth",
        "bullets": [
            "Modern, mobile-first responsive architecture built for maximum conversion rates",
            "Blazing-fast load speeds, clean code, and intuitive user navigation",
            "Built-in on-page SEO foundations and seamless CRM/analytics integrations",
        ],
    },
    "Mobile App Development": {
        "focus": "Native and cross-platform iOS & Android mobile applications that users love",
        "bullets": [
            "Intuitive UI/UX paired with high-performance Flutter/React Native or Native codebases",
            "Scalable cloud backend, secure authentication, and real-time push notification workflows",
            "End-to-end App Store & Google Play launch support with continuous updates",
        ],
    },
    "Software / Custom Development": {
        "focus": "Bespoke SaaS applications, internal workflow automations, and enterprise software",
        "bullets": [
            "Tailored internal tools and custom portals that eliminate repetitive manual tasks",
            "Robust API integrations connecting your CRMs, billing, and operational systems",
            "Secure, scalable cloud-native architectures ready to support thousands of daily users",
        ],
    },
    "Digital Marketing & SEO": {
        "focus": "Data-driven organic search ranking, paid acquisition, and inbound lead pipelines",
        "bullets": [
            "High-intent keyword SEO optimization to outrank regional and national competitors",
            "Conversion-rate optimized landing pages and targeted customer acquisition funnels",
            "Transparent weekly ROI tracking with clear revenue and lead metrics",
        ],
    },
    "Branding & UI/UX Design": {
        "focus": "World-class visual identity, product UI/UX, and distinctive design systems",
        "bullets": [
            "Comprehensive brand guidelines, premium typography, and memorable design assets",
            "Figma wireframing and user journey mapping optimized for frictionless engagement",
            "Modern, trust-building aesthetic that positions you as the market leader in your sector",
        ],
    },
    "IT Consulting & Support": {
        "focus": "Enterprise-grade IT infrastructure, cloud migration, security, and continuous support",
        "bullets": [
            "Cloud architecture audit and cost optimization across AWS, Azure, and Google Cloud",
            "Proactive security hardening, automated backups, and 24/7 uptime monitoring",
            "Dedicated tech support team ensuring zero unplanned downtime for your operations",
        ],
    },
}

# Heuristic keyword matching when "Suggested Service" is blank in Excel
SERVICE_MATCHING_RULES = {
    "Website Design & Development": [
        "website", "web design", "site", "redesign", "online presence", "landing page",
        "no website", "outdated website", "slow website", "portfolio"
    ],
    "Mobile App Development": [
        "mobile", "ios", "android", "app", "on-demand", "booking", "food delivery",
        "ride", "fintech app", "fitness app", "social app"
    ],
    "Software / Custom Development": [
        "saas", "software", "crm", "erp", "portal", "automation", "dashboard",
        "internal tool", "logistics", "supply chain", "api", "cloud platform"
    ],
    "Digital Marketing & SEO": [
        "marketing", "seo", "traffic", "leads", "sales", "ecommerce", "retail",
        "shop", "store", "conversion", "advertising", "social media", "growth"
    ],
    "Branding & UI/UX Design": [
        "brand", "branding", "logo", "ui/ux", "ux", "ui", "design", "creative",
        "fashion", "luxury", "redesign", "identity"
    ],
    "IT Consulting & Support": [
        "it support", "consulting", "infrastructure", "security", "cloud", "server",
        "cybersecurity", "devops", "healthcare it", "legal", "finance compliance"
    ],
}
