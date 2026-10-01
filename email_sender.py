"""
Skytecher Cold Email Automation - Email Sender Module
=====================================================
Handles Gmail SMTP transmission via SSL/TLS, multipart MIME formatting,
exponential backoff retries, anti-spam random delay throttling, and daily quotas.
"""

import json
import logging
import smtplib
import ssl
import time
import random
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
from pathlib import Path
from typing import Tuple, Dict, Any, Optional

import config

# Setup module logger to write to both console and email_log.txt
logger = logging.getLogger("SkytecherSender")
logger.setLevel(logging.INFO)

if not logger.handlers:
    # File handler
    file_handler = logging.FileHandler(config.LOG_FILE, encoding="utf-8")
    file_fmt = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(file_fmt)
    logger.addHandler(file_handler)

    # Console handler
    console_handler = logging.StreamHandler()
    console_fmt = logging.Formatter("[%(levelname)s] %(message)s")
    console_handler.setFormatter(console_fmt)
    logger.addHandler(console_handler)


class DailyTracker:
    """
    Tracks how many emails have been sent today to adhere to DAILY_SEND_LIMIT.
    Persists count to daily_stats.json.
    """

    def __init__(self, stats_file: Path):
        self.stats_file = stats_file

    def _read_data(self) -> Dict[str, Any]:
        if not self.stats_file.exists():
            return {}
        try:
            with open(self.stats_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _write_data(self, data: Dict[str, Any]) -> None:
        try:
            with open(self.stats_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            logger.warning(f"Could not persist daily stats: {e}")

    def get_count_today(self) -> int:
        today_str = datetime.date.today().isoformat()
        data = self._read_data()
        return data.get(today_str, 0)

    def increment_count(self) -> int:
        today_str = datetime.date.today().isoformat()
        data = self._read_data()
        count = data.get(today_str, 0) + 1
        data[today_str] = count
        self._write_data(data)
        return count


class EmailSender:
    """
    Manages Gmail SMTP connections, authentication, and safe email transmission.
    """

    def __init__(self):
        self.smtp_server = config.SMTP_SERVER
        self.smtp_port = config.SMTP_PORT
        self.use_ssl = config.USE_SSL
        self.username = config.SENDER_EMAIL
        self.password = config.GMAIL_APP_PASSWORD.replace(" ", "")  # App passwords might have spaces
        self.daily_tracker = DailyTracker(config.DAILY_STATS_FILE)

    def test_connection(self) -> Tuple[bool, str]:
        """
        Tests connection and authentication to Gmail SMTP server.
        """
        if not self.password or self.password.startswith("your_"):
            return (
                False,
                "Gmail App Password is not set! Please set GMAIL_APP_PASSWORD in your .env file.",
            )

        try:
            server = self._create_connection()
            server.quit()
            return True, f"SMTP Connection & Authentication successful as {self.username}!"
        except smtplib.SMTPAuthenticationError as e:
            msg = (
                f"SMTP Authentication Error (Code {e.smtp_code}): {e.smtp_error.decode('utf-8', errors='ignore')}. "
                "Ensure you generated a 16-character 'App Password' from Google Account settings."
            )
            logger.error(msg)
            return False, msg
        except Exception as e:
            msg = f"Failed to connect to SMTP server ({self.smtp_server}:{self.smtp_port}): {str(e)}"
            logger.error(msg)
            return False, msg

    def _create_connection(self) -> smtplib.SMTP:
        """
        Creates and logs into an SMTP/SSL or SMTP/TLS connection.
        """
        if self.use_ssl or self.smtp_port == 465:
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(
                self.smtp_server, self.smtp_port, context=context, timeout=20
            )
        else:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=20)
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()

        server.login(self.username, self.password)
        return server

    def can_send_today(self) -> Tuple[bool, int, int]:
        """
        Returns (can_send, current_sent_count, daily_limit).
        """
        current_count = self.daily_tracker.get_count_today()
        limit = config.DAILY_SEND_LIMIT
        return (current_count < limit, current_count, limit)

    def send_email(
        self,
        recipient_email: str,
        subject: str,
        plain_body: str,
        html_body: str,
        recipient_name: Optional[str] = None,
    ) -> Tuple[bool, str]:
        """
        Sends a single email with multipart/alternative structure.
        Implements retries with exponential backoff for transient failures.
        """
        # Check daily sending limit
        can_send, count, limit = self.can_send_today()
        if not can_send:
            err_msg = f"Daily send limit reached ({count}/{limit} emails sent today). Sending paused."
            logger.error(err_msg)
            return False, err_msg

        # Construct MIME Message
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = f"{config.SENDER_NAME} <{self.username}>"
        if recipient_name:
            msg["To"] = f"{recipient_name} <{recipient_email}>"
        else:
            msg["To"] = recipient_email
        msg["Date"] = formatdate(localtime=True)
        msg["Message-ID"] = make_msgid(domain="skytecher.com")

        # Attach text and HTML parts (plain text first, HTML second)
        part_text = MIMEText(plain_body, "plain", "utf-8")
        msg.attach(part_text)
        if html_body and html_body.strip():
            part_html = MIMEText(html_body, "html", "utf-8")
            msg.attach(part_html)

        # Retry loop
        last_error = ""
        for attempt in range(1, config.MAX_RETRIES + 1):
            server = None
            try:
                logger.info(
                    f"Attempting to send email to {recipient_email} (Attempt {attempt}/{config.MAX_RETRIES})..."
                )
                server = self._create_connection()
                server.sendmail(self.username, [recipient_email], msg.as_string())
                server.quit()

                # Increment and log success
                new_count = self.daily_tracker.increment_count()
                logger.info(
                    f"SUCCESS: Email sent to {recipient_email} | Subject: '{subject}' (Today's count: {new_count}/{limit})"
                )
                return True, ""

            except smtplib.SMTPAuthenticationError as e:
                last_error = f"Authentication failed: {e.smtp_error.decode('utf-8', errors='ignore')}"
                logger.error(f"[Permanent Error] {last_error}")
                # Don't retry auth errors
                break

            except (smtplib.SMTPRecipientsRefused, smtplib.SMTPSenderRefused) as e:
                last_error = f"Recipient/Sender refused: {str(e)}"
                logger.error(f"[Permanent Error] {last_error}")
                break

            except (
                smtplib.SMTPServerDisconnected,
                smtplib.SMTPConnectError,
                smtplib.SMTPHeloError,
                TimeoutError,
                OSError,
            ) as e:
                last_error = f"Network/SMTP connection error: {str(e)}"
                logger.warning(
                    f"Transient error on attempt {attempt}: {last_error}. Backing off {config.RETRY_BACKOFF_SECONDS * attempt}s..."
                )
                if attempt < config.MAX_RETRIES:
                    time.sleep(config.RETRY_BACKOFF_SECONDS * attempt)

            except Exception as e:
                last_error = f"Unexpected error: {str(e)}"
                logger.warning(f"Error on attempt {attempt}: {last_error}")
                if attempt < config.MAX_RETRIES:
                    time.sleep(config.RETRY_BACKOFF_SECONDS * attempt)

            finally:
                if server is not None:
                    try:
                        server.close()
                    except Exception:
                        pass

        logger.error(f"FAILURE: Could not send email to {recipient_email} after {config.MAX_RETRIES} attempts. Reason: {last_error}")
        return False, last_error

    def apply_anti_spam_delay(self, min_sec: Optional[int] = None, max_sec: Optional[int] = None) -> int:
        """
        Sleeps for a random duration between min_sec and max_sec to simulate human behavior.
        Displays a clean console countdown timer.
        """
        min_delay = min_sec if min_sec is not None else config.MIN_DELAY_SECONDS
        max_delay = max_sec if max_sec is not None else config.MAX_DELAY_SECONDS

        if max_delay < min_delay:
            max_delay = min_delay

        sleep_time = random.randint(min_delay, max_delay)
        if sleep_time <= 0:
            return 0

        print(f"\n[Antispam Delay] Pausing for {sleep_time} seconds before the next email...")
        for remaining in range(sleep_time, 0, -1):
            print(f"  Next send in {remaining:02d}s...", end="\r", flush=True)
            time.sleep(1)
        print(" " * 40, end="\r", flush=True)  # Clear countdown line
        return sleep_time
