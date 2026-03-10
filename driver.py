"""
Selenium-based WhatsApp Web driver.
"""

from urllib.parse import quote

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    WebDriverException, StaleElementReferenceException,
)
from webdriver_manager.chrome import ChromeDriverManager

from config import CHROME_PROFILE_DIR, STATUS_SENT, STATUS_FAILED, STATUS_NO_WA
from utils import log, fatal, human_delay


class WhatsAppDriver:
    """Manages Chrome browser session for WhatsApp Web."""

    def __init__(self):
        self.driver = None

    def start(self):
        """Launch Chrome with WhatsApp Web profile (persists login)."""
        log("Starting Chrome browser...")
        options = Options()
        options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
        except WebDriverException as e:
            fatal(f"Cannot start Chrome: {e}")

        log("Opening WhatsApp Web...")
        self.driver.get("https://web.whatsapp.com")

        log("Waiting for WhatsApp Web to load (scan QR code if prompted)...")
        try:
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[contenteditable="true"][data-tab="3"]'))
            )
            log("WhatsApp Web loaded successfully!")
        except TimeoutException:
            fatal("WhatsApp Web did not load in 60 seconds. Please scan QR code and try again.")

    def is_alive(self):
        """Check if the browser session is still active."""
        try:
            _ = self.driver.title
            return True
        except Exception:
            return False

    def restart_if_needed(self):
        """Restart browser if it crashed or was closed."""
        if not self.is_alive():
            log("Browser session lost. Restarting...")
            try:
                self.driver.quit()
            except Exception:
                pass
            self.start()
            return True
        return False

    def send_message(self, phone, message):
        """Send a message to a phone number. Returns (status, error_details)."""
        url = f"https://web.whatsapp.com/send?phone={phone}&text={quote(message)}"
        self.driver.get(url)

        try:
            WebDriverWait(self.driver, 25).until(
                lambda d: self._find_message_box(d) or self._detect_invalid_number(d)
            )
        except TimeoutException:
            return STATUS_FAILED, "Page did not load in time"

        if self._detect_invalid_number(self.driver):
            self._dismiss_popup()
            return STATUS_NO_WA, "Phone number is not on WhatsApp"

        msg_box = self._find_message_box(self.driver)
        if not msg_box:
            return STATUS_FAILED, "Could not find message input box"

        human_delay(1.0, 2.0)

        try:
            msg_box.send_keys(Keys.ENTER)
        except Exception as e:
            return STATUS_FAILED, f"Could not press Enter: {e}"

        human_delay(3.0, 5.0)
        return STATUS_SENT, None

    # ── Private helpers ──────────────────────────────────────

    def _find_message_box(self, driver):
        """Find the WhatsApp message input box."""
        try:
            for selector in [
                'div[contenteditable="true"][data-tab="10"]',
                'div[contenteditable="true"][title="Type a message"]',
                "footer div[contenteditable='true']",
            ]:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                if elements:
                    return elements[0]
        except (NoSuchElementException, StaleElementReferenceException):
            pass
        return None

    def _detect_invalid_number(self, driver):
        """Detect if WhatsApp shows 'invalid number' popup."""
        try:
            for selector in [
                "div[data-animate-modal-popup='true']",
                "div._3J6wB",
            ]:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                for el in elements:
                    text = el.text.lower()
                    if any(k in text for k in ["invalid", "not on whatsapp", "phone number shared via url"]):
                        return True

            body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
            if "phone number shared via url is invalid" in body_text:
                return True
        except Exception:
            pass
        return False

    def _dismiss_popup(self):
        """Try to close any popup/dialog."""
        try:
            for btn in self.driver.find_elements(By.CSS_SELECTOR, "div[role='button']"):
                if btn.text.strip().lower() in ("ok", "close"):
                    btn.click()
                    human_delay(0.5, 1.0)
                    return
        except Exception:
            pass

    def quit(self):
        """Close the browser."""
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
