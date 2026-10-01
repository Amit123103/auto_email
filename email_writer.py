"""
Skytecher Cold Email Automation - Email Writer Module
=====================================================
Generates high-converting, personalized cold outreach emails using
NVIDIA NIM Models (e.g. LLaMA 3.1 70B/8B, Nemotron) or Anthropic Claude API,
with an intelligent multi-variation template fallback.
"""

import json
import logging
import random
import re
from typing import Dict, Any, Tuple, Optional, List
import requests
import config

logger = logging.getLogger("SkytecherEmail")


def normalize_text(text: str) -> str:
    """Normalizes unusual unicode dashes and quotes to clean standard characters."""
    if not text:
        return ""
    replacements = {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2015": "-",
        "\u2018": "'",
        "\u2019": "'",
        "\u201a": "'",
        "\u201b": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u201e": '"',
        "\u201f": '"',
        "\u2026": "...",
        "\u00a0": " ",
        "\u2022": "-",
    }
    for orig, rep in replacements.items():
        text = text.replace(orig, rep)
    return text


def select_best_service(client: Dict[str, Any]) -> str:
    """
    Selects the best matching Skytecher service for the client.
    If 'suggested_service' is provided and valid, uses it.
    Otherwise, applies heuristic rules against website, industry, and description.
    """
    suggested = client.get("suggested_service", "").strip()
    if suggested:
        for official_service in config.COMPANY_SERVICES:
            if official_service.lower() in suggested.lower() or suggested.lower() in official_service.lower():
                return official_service
        return suggested

    website = str(client.get("website", "")).strip().lower()
    industry = str(client.get("industry", "")).strip().lower()
    description = str(client.get("description", "")).strip().lower()
    combined_text = f"{industry} {description}"

    if not website or website in ["none", "n/a", "no website", "nil", "-"]:
        return "Website Design & Development"

    scores: Dict[str, int] = {service: 0 for service in config.COMPANY_SERVICES}
    for service, keywords in config.SERVICE_MATCHING_RULES.items():
        for kw in keywords:
            if kw in combined_text:
                scores[service] += 2
            if kw in website:
                scores[service] += 1

    best_service = max(scores, key=scores.get)
    if scores[best_service] > 0:
        return best_service

    return "Website Design & Development"


def test_nvidia_connection(
    api_key: str,
    model: str = "openai/gpt-oss-20b",
    base_url: str = "https://integrate.api.nvidia.com/v1",
) -> Tuple[bool, str]:
    """
    Tests connectivity to NVIDIA NIM API with the given API key and model.
    Returns (True, message) on success, or (False, error_message) on failure.
    """
    api_key = (api_key or "").strip()
    if not api_key or api_key.startswith("your_"):
        return False, "NVIDIA API Key is empty or placeholder (must start with 'nvapi-')."

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    payload = {
        "model": model.strip(),
        "messages": [
            {"role": "user", "content": "Reply with only: OK"}
        ],
        "temperature": 0.1,
        "max_tokens": 20,
    }
    url = f"{base_url.rstrip('/')}/chat/completions"

    import time
    start_t = time.time()
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=25)
        elapsed = time.time() - start_t
        if resp.status_code == 200:
            data = resp.json()
            choice = data.get("choices", [{}])[0]
            msg_obj = choice.get("message", {}) or {}
            content = msg_obj.get("content") or msg_obj.get("reasoning_content") or msg_obj.get("reasoning") or ""
            reply = str(content).strip()
            return True, f"Connection successful in {elapsed:.2f}s! Model '{model}' is online and verified."
        else:
            err_msg = resp.text
            try:
                err_data = resp.json()
                if "error" in err_data:
                    err_msg = err_data["error"].get("message", err_msg)
            except Exception:
                pass
            return False, f"NVIDIA API Error ({resp.status_code}): {err_msg}"
    except requests.exceptions.Timeout:
        return False, "Request timed out after 25s. The model may be initializing on NVIDIA servers."
    except Exception as e:
        return False, f"Network connection failed: {str(e)}"


class EmailWriter:
    """
    Generates personalized cold emails using NVIDIA NIM API, Anthropic Claude,
    or smart multi-variation templates.
    """

    def __init__(self, provider: Optional[str] = None):
        self.provider = (provider or config.AI_PROVIDER or "nvidia").lower()
        self.use_ai = config.USE_AI

        # NVIDIA Setup
        self.nvidia_key = config.NVIDIA_API_KEY
        self.nvidia_base_url = config.NVIDIA_BASE_URL or "https://integrate.api.nvidia.com/v1"
        self.nvidia_model = config.NVIDIA_MODEL or "openai/gpt-oss-20b"
        self.has_nvidia = bool(self.nvidia_key and not self.nvidia_key.startswith("your_"))

        # Anthropic Setup
        self.anthropic_key = config.ANTHROPIC_API_KEY
        self.anthropic_model = config.ANTHROPIC_MODEL or "claude-3-5-sonnet-20241022"
        self.has_anthropic = bool(self.anthropic_key and not self.anthropic_key.startswith("your_"))

    def generate_email(
        self,
        client: Dict[str, Any],
        provider_override: Optional[str] = None,
        model_override: Optional[str] = None,
    ) -> Dict[str, str]:
        """
        Generates subject, plain-text body, and HTML body for a given client record.
        Returns dict with keys: 'subject', 'plain_body', 'html_body', 'service', 'source'.
        """
        service = select_best_service(client)
        active_provider = (provider_override or self.provider).lower()
        active_nvidia_model = model_override or self.nvidia_model

        # 1. Try NVIDIA NIM API if selected or available
        if self.use_ai and (active_provider == "nvidia" or (active_provider != "anthropic" and self.has_nvidia)):
            if self.has_nvidia:
                try:
                    result = self._generate_with_nvidia(client, service, model_override=active_nvidia_model)
                    if result:
                        result["service"] = service
                        result["source"] = f"NVIDIA NIM ({active_nvidia_model})"
                        return result
                except Exception as e:
                    logger.warning(f"NVIDIA API call failed for {client.get('company_name')}: {e}")

        # 2. Try Anthropic Claude if selected or available
        if self.use_ai and (active_provider == "anthropic" or self.has_anthropic):
            if self.has_anthropic:
                try:
                    result = self._generate_with_claude(client, service)
                    if result:
                        result["service"] = service
                        result["source"] = f"Anthropic Claude ({self.anthropic_model})"
                        return result
                except Exception as e:
                    logger.warning(f"Claude API failed for {client.get('company_name')}: {e}")

        # 3. Smart Template Fallback
        result = self._generate_with_template(client, service)
        result["service"] = service
        result["source"] = "Smart Template Engine"
        return result

    def _build_prompts(self, client: Dict[str, Any], service: str) -> Tuple[str, str]:
        """Constructs standardized system and user prompts for LLMs."""
        client_name = client.get("client_name") or "Business Leader"
        company_name = client.get("company_name") or "your company"
        industry = client.get("industry") or "your industry"
        website = client.get("website") or "Not provided"
        city = client.get("city") or ""
        description = client.get("description") or "Growing business looking to scale operations."
        language = config.EMAIL_LANGUAGE

        service_info = config.SERVICE_DETAILS.get(
            service,
            {"focus": service, "bullets": ["Tailored technical architecture", "Measurable ROI"]}
        )

        sender_phone = getattr(config, "COMPANY_PHONE", "+91-8960061745")
        system_prompt = f"""You are an elite B2B cold email copywriter for Skytecher.
Sender details:
- Name: Skytecher Team
- Company: Skytecher
- Offering: Custom Software, High-Converting Websites & Mobile App Solutions
- Phone / Website: {sender_phone} | skytecher.com
- Contact Email: skytechersolutions@gmail.com

Your objective is to craft a warm, concise, highly personalized cold email pitching Skytecher's digital engineering and growth services.

CRITICAL INSTRUCTIONS:
- Tone: Professional, friendly, confident but polite, peer-to-peer.
- Length: STRICTLY under 120 words total (keep it scannable, punchy, and formatted for easy reading).
- NO SPAM WORDS: Absolutely DO NOT use words like "FREE!!!", "Guaranteed", "Once in a lifetime", "Risk-free", "Act Now".
- Target Language: {language}.
- Follow this exact structure:
  1. Catchy subject line (personalized with client company or result, under 50 characters).
  2. Personalized greeting: "Hi {client_name},"
  3. One-line opening showing you researched their business and acknowledge what they do.
  4. Short paragraph on the specific problem or pain point you can solve for their business.
  5. Exactly 3 bullet points on the tangible benefits or results Skytecher offers.
  6. One proof point (client result, metric, or case study).
  7. Clear, low-pressure call to action (quick 15-minute call).
  8. Polite closing: "Best regards,"
  9. Professional signature:
     Skytecher Team
     Skytecher
     {sender_phone} | skytecher.com
  10. P.S. line adding subtle urgency without sounding pushy.
  11. Opt-out footer:
     "If you'd prefer not to receive emails from us, just reply 'Unsubscribe'."

OUTPUT FORMAT:
Respond ONLY with a valid, parseable JSON object without markdown fences, reasoning, explanation, or preamble. Begin immediately with '{' and end with '}':
{{
  "subject": "string (catchy, under 50 characters, personalized with company name)",
  "opening": "string (1 line showing you researched their business)",
  "problem_opportunity": "string (1 short paragraph on their specific industry pain point)",
  "bullet_points": ["bullet 1 benefit", "bullet 2 benefit", "bullet 3 benefit"],
  "credibility": "string (1 proof point or client metric)",
  "cta": "string (low-pressure ask for a 15-minute call this week)",
  "ps_line": "string (subtle urgency without sounding pushy)"
}}
"""

        has_web = client.get("has_website") or ("Yes" if website and website != "Not provided" else "No")
        lead_type = client.get("lead_type") or "Cold Lead"
        user_prompt = f"""Generate a personalized cold outreach email for:
- Client Name: {client_name}
- Company Name: {company_name}
- Industry: {industry}
- Location: {city}
- Website: {website} (Has Website: {has_web})
- Lead Status: {lead_type}
- Business Details / Context: {description}
- Pitched Skytecher Service: {service}
- Service Focus: {service_info.get('focus')}
- Reference Bullet Ideas: {json.dumps(service_info.get('bullets', []))}
"""
        return system_prompt, user_prompt

    def _generate_with_nvidia(
        self, client: Dict[str, Any], service: str, model_override: Optional[str] = None
    ) -> Optional[Dict[str, str]]:
        """
        Calls NVIDIA NIM API (OpenAI-compatible) to author the cold email.
        """
        active_model = model_override or self.nvidia_model
        system_prompt, user_prompt = self._build_prompts(client, service)
        client_name = client.get("client_name") or "Business Leader"
        company_name = client.get("company_name") or "your company"

        content = None

        # Method 1: Using openai client if installed
        try:
            from openai import OpenAI
            ai_client = OpenAI(base_url=self.nvidia_base_url, api_key=self.nvidia_key, timeout=35.0)
            completion = ai_client.chat.completions.create(
                model=active_model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0.6,
                max_tokens=2048,
            )
            choice = completion.choices[0]
            content = choice.message.content or ""
            if not content and hasattr(choice.message, "reasoning_content"):
                content = getattr(choice.message, "reasoning_content", "") or ""
            if not content and hasattr(choice.message, "reasoning"):
                content = getattr(choice.message, "reasoning", "") or ""
            content = str(content).strip()
        except Exception as e:
            logger.info(f"OpenAI SDK call to NVIDIA failed ({e}), attempting direct HTTP request...")

        # Method 2: Fallback to direct HTTP request via requests
        if not content:
            headers = {
                "Authorization": f"Bearer {self.nvidia_key}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            }
            payload = {
                "model": active_model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.6,
                "max_tokens": 2048,
            }
            url = f"{self.nvidia_base_url.rstrip('/')}/chat/completions"
            resp = requests.post(url, headers=headers, json=payload, timeout=35)
            if resp.status_code == 200:
                data = resp.json()
                choice = data.get("choices", [{}])[0]
                msg_obj = choice.get("message", {}) or {}
                content = msg_obj.get("content") or msg_obj.get("reasoning_content") or msg_obj.get("reasoning") or ""
                content = str(content).strip()
            else:
                raise RuntimeError(f"NVIDIA API HTTP {resp.status_code}: {resp.text}")

        return self._parse_llm_json_response(content, client_name, company_name, service)

    def _generate_with_claude(self, client: Dict[str, Any], service: str) -> Optional[Dict[str, str]]:
        """Calls Anthropic Claude API."""
        import anthropic
        ai_client = anthropic.Anthropic(api_key=self.anthropic_key)
        system_prompt, user_prompt = self._build_prompts(client, service)
        client_name = client.get("client_name") or "Business Leader"
        company_name = client.get("company_name") or "your company"

        response = ai_client.messages.create(
            model=self.anthropic_model,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        content = response.content[0].text.strip()
        return self._parse_llm_json_response(content, client_name, company_name, service)

    def _parse_llm_json_response(
        self, content: str, client_name: str, company_name: str, service: str
    ) -> Dict[str, str]:
        """Parses LLM output into clean subject, plain body, and HTML."""
        # 1. Remove thinking/reasoning tags if generated by reasoning models
        cleaned = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL).strip()
        if not cleaned and content:
            cleaned = content.strip()

        data = {}
        # 1. Try finding json inside markdown code fence
        fence_match = re.search(r"```(?:json)?\s*(\{[\s\S]*?\})\s*```", cleaned)
        if fence_match:
            try:
                data = json.loads(fence_match.group(1))
            except Exception:
                pass

        # 2. Try parsing cleaned directly
        if not data:
            try:
                data = json.loads(cleaned)
            except Exception:
                pass

        # 3. Try finding outermost { and }
        if not data:
            first_brace = cleaned.find("{")
            last_brace = cleaned.rfind("}")
            if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
                candidate = cleaned[first_brace : last_brace + 1]
                try:
                    data = json.loads(candidate)
                except Exception:
                    pass

        # 4. Fallback regex extraction if model output was slightly malformed or truncated
        if not data:
            sub_m = re.search(r'"subject"\s*:\s*"([^"\r\n]+)"', cleaned)
            op_m = re.search(r'"opening"\s*:\s*"([^"\r\n]+)"', cleaned)
            prob_m = re.search(r'"problem_opportunity"\s*:\s*"([^"\r\n]+)"', cleaned)
            cred_m = re.search(r'"credibility"\s*:\s*"([^"\r\n]+)"', cleaned)
            cta_m = re.search(r'"cta"\s*:\s*"([^"\r\n]+)"', cleaned)
            ps_m = re.search(r'"ps_line"\s*:\s*"([^"\r\n]+)"', cleaned)

            bullets_list = []
            bullet_block = re.search(r'"bullet_points"\s*:\s*\[([\s\S]*?)\]', cleaned)
            if bullet_block:
                bullets_list = [b.strip() for b in re.findall(r'"([^"\r\n]{5,150})"', bullet_block.group(1))]

            if sub_m or op_m or prob_m:
                data = {
                    "subject": sub_m.group(1) if sub_m else f"Quick question regarding {company_name}",
                    "opening": op_m.group(1) if op_m else "",
                    "problem_opportunity": prob_m.group(1) if prob_m else "",
                    "bullet_points": bullets_list,
                    "credibility": cred_m.group(1) if cred_m else "",
                    "cta": cta_m.group(1) if cta_m else "",
                    "ps_line": ps_m.group(1) if ps_m else "",
                }

        service_info = config.SERVICE_DETAILS.get(
            service,
            {"focus": service, "bullets": ["Tailored technical architecture", "Measurable ROI"]}
        )

        # If model returned structured JSON or salvaged fields
        if data and ("subject" in data or "opening" in data or "problem_opportunity" in data):
            subject = data.get("subject", f"Quick question regarding {company_name}")
            phone = getattr(config, "COMPANY_PHONE", "+91-8960061745")
            ps_line = data.get("ps_line", "P.S. We have two onboarding slots open this month - happy to reserve one if there's a strong mutual fit.")
            ps_txt = f"\n\n{ps_line}" if ps_line else ""

            raw_opening = data.get("opening", "").strip()
            # Clean duplicate greeting in opening if LLM included it
            clean_opening = re.sub(r"^(hi|hello|dear)\s+[^,\n]+[,:\-]?\s*", "", raw_opening, flags=re.IGNORECASE).strip()
            if not clean_opening:
                clean_opening = f"I came across {company_name} and was impressed by your presence in the market."

            greeting_opening = f"Hi {client_name},\n\n{clean_opening}"
            opening_clean = clean_opening

            problem = data.get("problem_opportunity", f"Many businesses in your industry face challenges scaling their digital presence efficiently.").strip()
            bullets = data.get("bullet_points") or service_info.get("bullets", [])
            bullets_txt = "\n".join([f"- {b}" for b in bullets])
            credibility = data.get("credibility", "For example, we helped a recent client boost client inquiries by 40% in under 60 days.").strip()
            cta = data.get("cta", "Would you be open to a quick 15-minute call this week to see if we can help?").strip()

            # Always guarantee clean, simple, professional plain text structure
            plain_body = (
                f"{greeting_opening}\n\n"
                f"{problem} At Skytecher, we help solve this by:\n"
                f"{bullets_txt}\n\n"
                f"{credibility}\n\n"
                f"{cta}\n\n"
                f"Best regards,\n\n"
                f"Skytecher Team\n"
                f"Skytecher\n"
                f"{phone} | skytecher.com\n"
                f"skytechersolutions@gmail.com"
                f"{ps_txt}\n\n"
                f"---\n"
                f"If you'd prefer not to receive emails from us, just reply 'Unsubscribe'."
            )

            html_body = self._build_html(
                client_name=client_name,
                company_name=company_name,
                opening=opening_clean,
                problem=problem,
                service=service,
                bullets=bullets,
                credibility=credibility,
                cta=cta,
                ps_line=ps_line,
            )
        else:
            # If freeform text or unparseable, check for json fragments
            if "{" in cleaned or '"subject"' in cleaned:
                fallback = self._generate_with_template(
                    {"client_name": client_name, "company_name": company_name}, service
                )
                return fallback

            text_lines = cleaned.splitlines()
            subject = f"Quick question regarding {company_name}"
            body_lines = []
            for line in text_lines:
                if re.match(r"^Subject:\s*", line, re.IGNORECASE):
                    subject = re.sub(r"^Subject:\s*", "", line, flags=re.IGNORECASE).strip()
                else:
                    body_lines.append(line)

            plain_body = "\n".join(body_lines).strip()
            if not plain_body:
                fallback = self._generate_with_template(
                    {"client_name": client_name, "company_name": company_name}, service
                )
                return fallback

            if "unsubscribe" not in plain_body.lower():
                plain_body += "\n\n---\nIf you'd prefer not to receive emails from us, just reply 'Unsubscribe'."

            html_body = f"""<div style="font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 1.6; color: #111827; max-width: 600px;">
{plain_body.replace(chr(10), '<br>')}
</div>"""

        return {
            "subject": normalize_text(subject),
            "plain_body": normalize_text(plain_body),
            "html_body": normalize_text(html_body),
        }

    def _generate_with_template(
        self, client: Dict[str, Any], service: str
    ) -> Dict[str, str]:
        """Generates cold email using rotating smart templates."""
        client_name = client.get("client_name") or "Business Leader"
        company_name = client.get("company_name") or "your company"
        industry = client.get("industry") or "your industry"
        city = client.get("city") or ""
        website = client.get("website") or ""
        desc = client.get("description") or ""

        service_data = config.SERVICE_DETAILS.get(
            service,
            {
                "focus": service,
                "bullets": [
                    "Bespoke technical solutions engineered for your business goals",
                    "Rapid implementation cycle with dedicated project management",
                    "Transparent communication with measurable ROI",
                ],
            },
        )
        bullets = service_data.get("bullets", [])

        subjects = [
            f"Quick question regarding {company_name}'s digital roadmap",
            f"Idea for {company_name} | {service}",
            f"Scaling {company_name}'s digital presence",
            f"{company_name} + Skytecher: quick collaboration thought",
        ]
        subject = random.choice(subjects)

        city_phrase = f" based in {city}" if city else ""
        if desc and len(desc) > 20:
            openings = [
                f"I came across {company_name}{city_phrase} and was impressed by your work in {industry.lower()}, particularly your dedication to client success.",
                f"While researching innovative teams in the {industry.lower()} space{city_phrase}, I noticed the momentum {company_name} has been building.",
                f"I've been following {company_name}'s footprint in {industry.lower()} and wanted to reach out directly regarding your current digital initiatives.",
            ]
        else:
            openings = [
                f"I came across {company_name}{city_phrase} and was very impressed by your presence in the {industry.lower()} sector.",
                f"While reviewing high-potential organizations in {industry.lower()}{city_phrase}, {company_name} immediately caught our attention.",
                f"I've been looking into forward-thinking businesses in {industry.lower()}, and {company_name}'s work stood out.",
            ]
        opening = random.choice(openings)

        if service == "Website Design & Development":
            if not website or website in ["none", "n/a", "-"]:
                problem = f"In today's digital landscape, customers expect an instant, high-speed online hub to evaluate and trust a business like {company_name}."
            else:
                problem = f"With customer expectations higher than ever, having a fast, modern web experience is essential for converting visitors into loyal clients for {company_name}."
        elif service == "Mobile App Development":
            problem = f"As mobile-first transactions continue to outpace desktop, providing a seamless on-demand mobile application can significantly boost retention for {company_name}."
        elif service == "Software / Custom Development":
            problem = f"Many growing firms in {industry.lower()} hit bottlenecks when outgrowing manual spreadsheets and off-the-shelf software tools."
        elif service == "Digital Marketing & SEO":
            problem = f"With increasing competition in {industry.lower()}, capturing high-intent organic search traffic is one of the highest-ROI ways to generate consistent client inquiries."
        elif service == "Branding & UI/UX Design":
            problem = f"A distinctive brand identity and polished user interface immediately signal enterprise-grade quality and justify premium pricing in {industry.lower()}."
        else:
            problem = f"Maintaining rock-solid cloud infrastructure and foolproof data security is critical to ensure {company_name} operates with zero unexpected downtime."

        credibility_statements = [
            "At Skytecher, we've helped over 50+ growing businesses modernize their technology and drive tangible growth.",
            "Our engineering team at Skytecher has spent the last 5 years helping businesses like yours turn technology into a core competitive edge.",
            "Skytecher specializes in end-to-end digital solutions, partnering with growing firms to deliver measurable business outcomes.",
        ]
        credibility = random.choice(credibility_statements)

        ctas = [
            "Would you be open to a brief 15-minute call this Thursday to share notes, or simply reply to this email to see if there's a fit?",
            "Are you free for a quick 15-minute introductory call next week, or feel free to reply directly here if you'd like more details.",
            "Would you be against a 15-minute chat this week to explore ideas for {company_name}? Alternatively, let me know if now isn't the right time.",
        ]
        cta = random.choice(ctas).format(company_name=company_name)

        phone = getattr(config, "COMPANY_PHONE", "+91-8960061745")
        bullets_text = "\n".join([f"- {b}" for b in bullets])
        ps_line = "P.S. We have two onboarding slots open this month - happy to reserve one if there's a strong mutual fit."

        plain_body = f"""Hi {client_name},

{opening}

{problem} At Skytecher, we help solve this by:
{bullets_text}

{credibility}

{cta}

Best regards,

Skytecher Team
Skytecher
{phone} | skytecher.com
skytechersolutions@gmail.com

{ps_line}

---
If you'd prefer not to receive emails from us, just reply 'Unsubscribe'.
"""

        html_body = self._build_html(
            client_name=client_name,
            company_name=company_name,
            opening=opening,
            problem=problem,
            service=service,
            bullets=bullets,
            credibility=credibility,
            cta=cta,
            ps_line=ps_line,
        )

        return {
            "subject": normalize_text(subject),
            "plain_body": normalize_text(plain_body.strip()),
            "html_body": normalize_text(html_body.strip()),
        }

    def _build_html(
        self,
        client_name: str,
        company_name: str,
        opening: str,
        problem: str,
        service: str,
        bullets: List[str],
        credibility: str,
        cta: str,
        ps_line: str = "",
    ) -> str:
        """
        Builds a simple, professional, clean email message format.
        No table cards, no gradients, no grey background wrappers —
        looks 100% like a genuine 1-on-1 personal email to maximize inbox deliverability.
        """
        phone = getattr(config, "COMPANY_PHONE", "+91-8960061745")
        bullet_items = "".join(
            [f'<li style="margin-bottom: 6px;">{b}</li>' for b in bullets]
        )
        ps_html = (
            f'<p style="margin: 16px 0 16px 0; font-size: 14px; color: #4b5563;">{ps_line}</p>'
            if ps_line
            else ""
        )

        html = f"""<div style="font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 1.6; color: #111827; max-width: 600px;">
  <p style="margin: 0 0 16px 0;">Hi {client_name},</p>

  <p style="margin: 0 0 16px 0;">{opening}</p>

  <p style="margin: 0 0 12px 0;">{problem} At Skytecher, we help solve this by:</p>

  <ul style="margin: 0 0 16px 0; padding-left: 20px;">
    {bullet_items}
  </ul>

  <p style="margin: 0 0 16px 0;">{credibility}</p>

  <p style="margin: 0 0 20px 0;">{cta}</p>

  <p style="margin: 0 0 4px 0;">Best regards,</p>
  <p style="margin: 0 0 2px 0; font-weight: bold;">Skytecher Team</p>
  <p style="margin: 0 0 4px 0;">Skytecher</p>
  <p style="margin: 0 0 16px 0; color: #4b5563;">
    <a href="tel:{phone.replace(' ', '')}" style="color: #2563eb; text-decoration: none;">{phone}</a> |
    <a href="https://skytecher.com" target="_blank" style="color: #2563eb; text-decoration: none;">skytecher.com</a> |
    <a href="mailto:skytechersolutions@gmail.com" style="color: #4b5563; text-decoration: none;">skytechersolutions@gmail.com</a>
  </p>

  {ps_html}

  <p style="margin-top: 24px; font-size: 12px; color: #9ca3af; border-top: 1px solid #e5e7eb; padding-top: 12px;">
    If you'd prefer not to receive emails from us, just reply 'Unsubscribe'.
  </p>
</div>"""
        return html

    def generate_follow_up(
        self, client: Dict[str, Any], previous_subject: Optional[str] = None
    ) -> Dict[str, str]:
        """
        Generates a concise 3-day follow-up email if no reply was received.
        """
        client_name = client.get("client_name") or "there"
        company_name = client.get("company_name") or "your team"
        phone = getattr(config, "COMPANY_PHONE", "+91-8960061745")
        subject = f"Re: {previous_subject}" if previous_subject else f"Quick follow up regarding {company_name}"

        plain_body = f"""Hi {client_name},

I know you're busy, so keeping this brief.

Following up on my note below - we recently helped a growing business in your space streamline their digital operations and increase customer inquiries by 40% in 60 days.

Would 10 minutes this Thursday or Friday work for a quick introductory chat?

Best regards,

Skytecher Team
Skytecher
{phone} | skytecher.com
skytechersolutions@gmail.com

P.S. If you're all set on digital and software development for now, no worries at all - just let me know and I won't follow up again.
"""
        html_body = f"""<div style="font-family: Arial, Helvetica, sans-serif; font-size: 15px; line-height: 1.6; color: #111827; max-width: 600px;">
  <p style="margin: 0 0 16px 0;">Hi {client_name},</p>
  <p style="margin: 0 0 16px 0;">I know you're busy, so keeping this brief.</p>
  <p style="margin: 0 0 16px 0;">Following up on my note below - we recently helped a growing business in your space streamline their digital operations and increase customer inquiries by 40% in 60 days.</p>
  <p style="margin: 0 0 20px 0;"><strong>Would 10 minutes this Thursday or Friday work for a quick introductory chat?</strong></p>
  <p style="margin: 0 0 4px 0;">Best regards,</p>
  <p style="margin: 0 0 2px 0; font-weight: bold;">Skytecher Team</p>
  <p style="margin: 0 0 4px 0;">Skytecher</p>
  <p style="margin: 0 0 16px 0; color: #4b5563;">
    <a href="tel:{phone.replace(' ', '')}" style="color: #2563eb; text-decoration: none;">{phone}</a> |
    <a href="https://skytecher.com" target="_blank" style="color: #2563eb; text-decoration: none;">skytecher.com</a> |
    <a href="mailto:skytechersolutions@gmail.com" style="color: #4b5563; text-decoration: none;">skytechersolutions@gmail.com</a>
  </p>
  <p style="margin-top: 16px; font-size: 14px; color: #4b5563;">
    P.S. If you're all set on digital and software development for now, no worries at all - just let me know and I won't follow up again.
  </p>
</div>"""

        return {
            "subject": normalize_text(subject),
            "plain_body": normalize_text(plain_body.strip()),
            "html_body": normalize_text(html_body.strip()),
        }
