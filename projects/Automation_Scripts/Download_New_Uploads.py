"""
Download_New_Uploads.py
 
Logs into the Digiverifier admin portal, asks for exact Applicant IDs,
opens only those cases' QC pages, downloads the
"Download All Data And Document" zip, unzips it, deletes the zip,
and moves the unzipped folder into the local New_Cases folder (ready
for Tracker_Filling_Automation.py to pick up).
 
-------------------------------------------------------------------
IMPORTANT - READ THIS FIRST
-------------------------------------------------------------------
This script was written WITHOUT access to the live portal (this dev
environment has no network access and no login of yours), so the
element selectors below are best-effort guesses built only from your
screenshots and description. Run it once with --headless off (the
default) so you can SEE what's happening, and if it stops at a step
with a "COULD NOT FIND ..." error, send me:
    1. the exact error printed, and
    2. a screenshot of that part of the page (or right-click ->
       Inspect on the element and paste the HTML)
...and I'll fix that one selector precisely. The automation flow is
implemented, but portal-specific labels and table selectors are still
best-effort guesses and must be confirmed against your live portal.
 
-------------------------------------------------------------------
SETUP (one-time)
-------------------------------------------------------------------
    pip install selenium webdriver-manager --break-system-packages
 
Your login is never hardcoded here. Before running, set it as
environment variables (PowerShell example):
 
    $env:PORTAL_USERNAME = "your.email@company.com"
    $env:PORTAL_PASSWORD = "your-password"
    python Download_New_Uploads.py                 (prompts for Applicant IDs)
 
If you skip this, the script prompts for username/password at the
terminal (password won't be shown on screen). It then prompts for the
Applicant IDs to download.
 
-------------------------------------------------------------------
USAGE
-------------------------------------------------------------------
    python Download_New_Uploads.py
    python Download_New_Uploads.py --new-cases-folder "D:\\Work\\New_Cases"
    python Download_New_Uploads.py --headless      (no visible browser window)
"""
 
import argparse
import getpass
import logging
import os
import shutil
import sys
import tempfile
import time
import zipfile
from datetime import datetime
 
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.common.exceptions import (
        TimeoutException, NoSuchElementException, ElementClickInterceptedException,
        StaleElementReferenceException,
    )
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print("Missing packages. Run this first:")
    print("    pip install selenium webdriver-manager --break-system-packages")
    sys.exit(1)
 
 
# ===========================
# CONFIG
# ===========================
 
# Some corporate networks do SSL inspection (a self-signed cert gets
# inserted into the HTTPS chain), which makes webdriver-manager's own
# download check fail with CERTIFICATE_VERIFY_FAILED even though Chrome
# itself browses fine. This tells webdriver-manager to skip verifying
# that connection - it only affects the one-time driver-version lookup,
# not your portal login or anything else.
os.environ.setdefault("WDM_SSL_VERIFY", "0")
 
PORTAL_URL = "https://vpnnewportal.digiverifier.com/#/admin/orgadminDashboard"
RUN_TS = datetime.now().strftime("%Y%m%d_%H%M%S")

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--new-cases-folder", default=os.path.join(os.path.expanduser("~"), "Downloads", "New_Cases"), help="Final destination folder (default: Downloads\\New_Cases)")
parser.add_argument("--download-folder", default=None, help="Temp zip download folder (default: Portal_Downloads_<run timestamp>, removed when empty)")
parser.add_argument("--headless", action="store_true", help="Run without a visible browser window")
parser.add_argument("--wait-timeout", type=int, default=30, help="Seconds to wait for each page element")
parser.add_argument("--log-file", default=None, help="Default: Tracker Filling-Summary.log in the current folder")
parser.add_argument(
    "--chromedriver-path", default=None,
    help="Path to an already-downloaded chromedriver.exe. Use this if "
         "auto-download fails (e.g. corporate network SSL inspection blocks it) - "
         "download the driver matching your Chrome version from "
         "https://googlechromelabs.github.io/chrome-for-testing/ yourself.",
)
args = parser.parse_args()
 
NEW_CASES_DIR = os.path.abspath(args.new_cases_folder)
AUTO_DOWNLOAD_DIR = args.download_folder is None
DOWNLOAD_DIR_NAME = args.download_folder or f"Portal_Downloads_{RUN_TS}"
DOWNLOAD_DIR = os.path.abspath(DOWNLOAD_DIR_NAME)
WAIT_TIMEOUT = args.wait_timeout
LOG_FILE = args.log_file or "Tracker Filling-Summary.log"
 
os.makedirs(NEW_CASES_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
 
# ===========================
# LOGGING
# ===========================
 
logger = logging.getLogger("new_uploads_downloader")
logger.setLevel(logging.INFO)
_fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%Y-%m-%d %H:%M:%S")
 
_console = logging.StreamHandler(sys.stdout)
_console.setFormatter(_fmt)
logger.addHandler(_console)
 
_file_handler = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8")
_file_handler.setFormatter(_fmt)
logger.addHandler(_file_handler)
 
logger.info("Run started. NewCasesFolder=%s", NEW_CASES_DIR)
 
 
# ===========================
# LOGIN CREDENTIALS - never hardcoded
# ===========================
 
username = os.environ.get("PORTAL_USERNAME") or input("Portal username/email: ").strip()
 
if os.environ.get("PORTAL_PASSWORD"):
    password = os.environ["PORTAL_PASSWORD"]
else:
    print(
        "Portal password (hidden while typing - nothing will show on "
        "screen, not even *, that's normal. Just type it and press Enter):"
    )
    try:
        password = getpass.getpass("> ")
    except Exception:
        # Some terminals (older cmd.exe, certain IDE consoles) can't do
        # hidden input at all - fall back to plain, visible input rather
        # than getting stuck.
        print(
            "(Hidden input isn't supported in this terminal - it will "
            "be VISIBLE as you type. Close this window after if that's "
            "a concern.)"
        )
        password = input("Portal password: ")
 
if not password:
    print("No password entered - aborting.")
    sys.exit(1)
 
 
# ===========================
# BROWSER SETUP
# ===========================
 
def build_driver():
    options = Options()
    if args.headless:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    # Preserve the authenticated Chrome window if automation fails, so the
    # user can inspect the portal state instead of losing the session.
    options.add_experimental_option("detach", True)
 
    # Auto-download zip files straight to DOWNLOAD_DIR, no "Save As" prompt.
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
 
    if args.chromedriver_path:
        driver_path = args.chromedriver_path
        if not os.path.isfile(driver_path):
            print(f"--chromedriver-path given but file not found: {driver_path}")
            sys.exit(1)
    else:
        try:
            driver_path = ChromeDriverManager().install()
        except Exception as e:
            print(
                "Could not auto-download chromedriver (network/SSL issue "
                f"most likely): {e}\n\n"
                "Fix: download the chromedriver matching your Chrome version "
                "from https://googlechromelabs.github.io/chrome-for-testing/ "
                "(check your Chrome version via chrome://version), unzip it "
                "anywhere, then re-run with:\n"
                '    python Download_New_Uploads.py --chromedriver-path "C:\\path\\to\\chromedriver.exe"'
            )
            sys.exit(1)
 
    service = Service(driver_path)
    return webdriver.Chrome(service=service, options=options)
 
 
def wait_for(driver, by, value, timeout=None, clickable=False):
    timeout = timeout or WAIT_TIMEOUT
    condition = EC.element_to_be_clickable((by, value)) if clickable else EC.presence_of_element_located((by, value))
    return WebDriverWait(driver, timeout).until(condition)
 
 
def try_click(driver, element):
    """Plain click, falling back to a JS click if something is overlapping it."""
    try:
        element.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", element)
 
 
# ===========================
# STEP 1 - LOGIN
# ===========================
 
def login(driver):
    logger.info("Opening portal: %s", PORTAL_URL)
    driver.get(PORTAL_URL)
 
    # ADJUST THIS if your login page uses different field names/ids.
    # Common patterns tried in order: id="email"/"username", type="email",
    # then any text input in a form as a last resort.
    login_field_candidates = [
        (By.ID, "email"),
        (By.ID, "username"),
        (By.NAME, "email"),
        (By.NAME, "username"),
        (By.CSS_SELECTOR, "input[type='email']"),
    ]
    password_field_candidates = [
        (By.CSS_SELECTOR, "input[type='password']"),
        (By.ID, "password"),
        (By.NAME, "password"),
    ]
    # Two-step login: email -> click Continue/Next -> password field
    # then appears. Tried only if the password field isn't already
    # present right after typing the email.
    continue_button_candidates = [
        (By.XPATH, "//button[contains(text(),'Continue')]"),
        (By.XPATH, "//button[contains(text(),'Next')]"),
        (By.XPATH, "//*[@type='submit' and contains(@value,'Continue')]"),
        (By.XPATH, "//button[@type='submit']"),
    ]
    login_button_candidates = [
        (By.XPATH, "//button[contains(text(),'Login')]"),
        (By.XPATH, "//button[contains(text(),'Sign in')]"),
        (By.XPATH, "//button[contains(text(),'Sign In')]"),
        (By.XPATH, "//button[@type='submit']"),
    ]
 
    def find_first(candidates, timeout=5, root=driver):
        for by, val in candidates:
            try:
                return wait_for(root, by, val, timeout=timeout)
            except TimeoutException:
                continue
        return None

    def find_visible_first(candidates, timeout=WAIT_TIMEOUT):
        """Wait once for any visible candidate, rather than timing out per selector."""
        def visible_candidate(d):
            for by, val in candidates:
                for element in d.find_elements(by, val):
                    try:
                        if element.is_displayed() and element.is_enabled():
                            return element
                    except StaleElementReferenceException:
                        continue
            return False
        try:
            return WebDriverWait(driver, timeout).until(visible_candidate)
        except TimeoutException:
            return None
 
    email_el = find_first(login_field_candidates)
 
    if email_el is None:
        raise RuntimeError(
            "COULD NOT FIND the username/email field on the login page. "
            "Send me the login page's HTML (right-click the email box -> Inspect)."
        )
 
    email_el.clear()
    email_el.send_keys(username)
 
    # Is the password field already on screen (single-step login)?
    pwd_el = None
    for by, val in password_field_candidates:
        for candidate in driver.find_elements(by, val):
            try:
                if candidate.is_displayed() and candidate.is_enabled():
                    pwd_el = candidate
                    break
            except StaleElementReferenceException:
                continue
        if pwd_el is not None:
            break
 
    if pwd_el is None:
        # Two-step login - click Continue/Next, then wait for the
        # password field to appear.
        continue_btn = find_first(continue_button_candidates, timeout=5)
        if continue_btn is None:
            raise RuntimeError(
                "Typed the email but COULD NOT FIND a password field OR a "
                "Continue/Next button. Send me the login page's HTML at "
                "this point (right-click -> Inspect near the email box)."
            )
        try_click(driver, continue_btn)
        logger.info("Clicked Continue/Next after entering email.")
 
        pwd_el = find_visible_first(password_field_candidates, timeout=WAIT_TIMEOUT)
        if pwd_el is None:
            raise RuntimeError(
                "Clicked Continue but COULD NOT FIND the password field "
                "afterwards. Send me the HTML of that next screen."
            )
 
    pwd_el.clear()
    pwd_el.send_keys(password)
 
    # The portal may authenticate and replace the login view while Selenium is
    # still resolving the submit button. Check for the dashboard first and
    # recover from stale references by locating a fresh visible button.
    def dashboard_visible():
        try:
            body = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
            return "upload details" in body and "view uploads" in body
        except Exception:
            return False

    time.sleep(.5)
    if dashboard_visible():
        logger.info("Login completed automatically; dashboard is already visible.")
        return

    login_btn = find_visible_first(login_button_candidates, timeout=5)
    if login_btn is not None:
        try:
            try_click(driver, login_btn)
            logger.info("Clicked Login/Sign-in button.")
        except StaleElementReferenceException:
            if dashboard_visible():
                logger.info("Login completed while the page refreshed; dashboard is visible.")
                return
            refreshed_btn = find_visible_first(login_button_candidates, timeout=5)
            if refreshed_btn is None:
                if dashboard_visible():
                    logger.info("Login completed while the page refreshed; dashboard is visible.")
                    return
                raise RuntimeError("Login page refreshed and its submit button disappeared before authentication was confirmed.")
            try_click(driver, refreshed_btn)
            logger.info("Re-found and clicked the refreshed Login/Sign-in button.")
    else:
        pwd_el.send_keys(Keys.RETURN)

    # The dashboard's visible label varies across portal versions and may
    # render after the Angular route changes. Wait on several independent
    # signs of successful login instead of one brittle exact dashboard label.
    # Some portal sessions need longer than the normal element timeout to
    # finish authentication after the sign-in click.
    login_wait = max(WAIT_TIMEOUT, 90)
    deadline = time.time() + login_wait
    last_url = ""
    otp_prompted = False
    while time.time() < deadline:
        last_url = driver.current_url
        body = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        if any(term in body for term in ("invalid password", "incorrect password", "invalid credentials", "login failed")):
            raise RuntimeError("Portal rejected the login credentials. Check the username/password and try again.")
        if any(term in body for term in ("otp", "one time password", "verification code", "two-factor", "2fa")):
            if not otp_prompted:
                logger.info("Portal is asking for a verification code; complete it in Chrome.")
                input("Chrome-la OTP/verification code enter panni submit pannitu, inga Enter press pannunga: ")
                # Continue waiting for the dashboard after the manual verification.
                otp_prompted = True
                deadline = time.time() + login_wait
                continue
        if "upload details" in body and "view uploads" in body:
            logger.info("Login successful, uploads dashboard is visible.")
            return
        time.sleep(0.5)
    screenshot = os.path.join(os.getcwd(), f"login_failure_{RUN_TS}.png")
    try:
        driver.save_screenshot(screenshot)
    except Exception:
        screenshot = "(screenshot could not be saved)"
    raise RuntimeError(
        "Login was submitted, but the portal did not reach the dashboard within "
        f"{login_wait}s (current URL: {last_url}). The browser may be showing an "
        f"OTP, consent, CAPTCHA, or portal error. Screenshot saved to: {screenshot}"
    )
 
 
# ===========================
# STEP 2 - OPEN INLINE NEW UPLOAD TABLE
# ===========================

def table_search_box(driver):
    """Find the visible search input used by the inline New Upload table.

    The portal's search field has changed markup between screens; do not
    require a literal placeholder. Prefer a search-labelled field, then a
    visible text input in the Conventional New Upload card.
    """
    candidates = []
    for element in driver.find_elements(By.CSS_SELECTOR, "input:not([type='hidden'])"):
        try:
            if not element.is_displayed() or not element.is_enabled():
                continue
            attrs = " ".join((element.get_attribute(a) or "") for a in
                             ("type", "placeholder", "aria-label", "name", "id", "title", "class")).lower()
            if "date" in attrs or (element.get_attribute("type") or "").lower() in ("date", "password", "email"):
                continue
            score = 0
            if "search" in attrs:
                score += 10
            if (element.get_attribute("type") or "").lower() == "search":
                score += 5
            # Search box is in the lower-right of the inline table card.
            rect = driver.execute_script(
                "const r=arguments[0].getBoundingClientRect(); return {x:r.x,y:r.y,w:r.width,h:r.height};", element
            )
            if rect["y"] > driver.execute_script("return window.innerHeight * .35"):
                score += 2
            if rect["w"] >= 100:
                score += 1
            candidates.append((score, element))
        except StaleElementReferenceException:
            continue
    if not candidates:
        return None
    candidates.sort(key=lambda pair: pair[0], reverse=True)
    return candidates[0][1]


def save_page_diagnostics(driver, prefix):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.join(os.getcwd(), f"{prefix}_{stamp}")
    try:
        driver.save_screenshot(base + ".png")
    except Exception:
        pass
    try:
        with open(base + ".html", "w", encoding="utf-8") as handle:
            # Keep useful selector structure while avoiding a dump of case
            # names, email addresses, phone numbers, and IDs from the portal.
            handle.write(driver.execute_script("""
                const esc=document.createElement('div');
                const chunks=['<!-- Sanitized control/table structure; text and input values omitted -->'];
                for(const el of document.querySelectorAll('input:not([type="hidden"])')) {
                  const r=el.getBoundingClientRect(); if(!r.width || !r.height) continue;
                  const clone=el.cloneNode(false); clone.removeAttribute('value');
                  chunks.push('<!-- visible input -->'+clone.outerHTML);
                }
                for(const table of document.querySelectorAll('table')) {
                  const headers=[...table.querySelectorAll('thead th')].map(x=>x.outerHTML).join('');
                  const row=table.querySelector('tbody tr');
                  if(!headers && !row) continue;
                  let rowHtml='';
                  if(row){ const copy=row.cloneNode(true); const walker=document.createTreeWalker(copy,NodeFilter.SHOW_TEXT);
                    while(walker.nextNode()) walker.currentNode.nodeValue='[text omitted]';
                    for(const el of copy.querySelectorAll('input')) el.removeAttribute('value');
                    rowHtml=copy.outerHTML;
                  }
                  chunks.push('<table><thead>'+headers+'</thead><tbody>'+rowHtml+'</tbody></table>');
                }
                return chunks.join('\\n');
            """))
    except Exception:
        pass
    details = []
    for element in driver.find_elements(By.CSS_SELECTOR, "input:not([type='hidden'])"):
        try:
            if element.is_displayed():
                details.append({key: element.get_attribute(key) for key in
                                ("type", "placeholder", "aria-label", "name", "id", "class", "value")})
        except StaleElementReferenceException:
            continue
    logger.info("Visible input diagnostics: %s", details)
    return base
 
def open_new_uploads_list(driver):
    """Click the user-confirmed In Progress chart point and wait for its search table."""
    if table_search_box(driver) is not None:
        return

    click_target = driver.execute_script("""
      const scale=window.devicePixelRatio||1;
      const x=592/scale, y=595/scale;
      if(x<0 || y<0 || x>=window.innerWidth || y>=window.innerHeight)
        return {error:`Scaled point (${x.toFixed(1)}, ${y.toFixed(1)}) from screenshot point (592, 595) is outside browser viewport ${window.innerWidth}x${window.innerHeight}`};
      const element=document.elementFromPoint(x,y);
      if(!element) return {error:`No chart element at (${x}, ${y})`};
      const r=element.getBoundingClientRect();
      return {element,x:x-(r.left+r.width/2),y:y-(r.top+r.height/2),
              tag:element.tagName,cls:String(element.className||''),
              viewport:`${window.innerWidth}x${window.innerHeight}`,scale};
    """)
    if not click_target or click_target.get("error"):
        prefix = save_page_diagnostics(driver, "in_progress_chart_diagnostics")
        raise RuntimeError(f"Could not click the chart at the provided point: {click_target}. Diagnostics: {prefix}")
    ActionChains(driver).move_to_element_with_offset(
        click_target["element"], click_target["x"], click_target["y"]
    ).click().perform()
    logger.info(
        "Clicked In Progress chart at user-provided screenshot point (592, 595), scaled by %.2f; target=%s class=%s viewport=%s.",
        click_target.get("scale", 1), click_target.get("tag"), click_target.get("cls"), click_target.get("viewport")
    )

    deadline = time.time() + min(WAIT_TIMEOUT, 20)
    while time.time() < deadline:
        if table_search_box(driver) is not None:
            logger.info("In Progress case table and search box are visible.")
            return
        time.sleep(.4)
    prefix = save_page_diagnostics(driver, "in_progress_chart_diagnostics")
    raise RuntimeError(
        "Clicked the In Progress chart point (screenshot 592, 595), but the case search field did not appear. "
        f"Diagnostics: {prefix}.png and {prefix}.html"
    )

def find_candidate_row(driver, applicant_id):
    """Find the exact ID's nearest row/container that contains its action controls."""
    id_node = driver.find_elements(By.XPATH,
        f"//tr[td[normalize-space(.)='{applicant_id}']] | //*[@role='row'][.//*[@role='cell' and normalize-space(.)='{applicant_id}']]")
    for node in id_node:
        try:
            if node.is_displayed():
                row = node
                return row
        except StaleElementReferenceException:
            continue
    return None


def click_case_action(driver, applicant_id):
    """Click the blue Action control in the exact matching row atomically.

    Selenium WebElement references can go stale while Angular refreshes the
    filtered row. Locate the ID and click its row control in one browser-side
    operation, retrying briefly if the row is being re-rendered.
    """
    script = r"""
      const id=arguments[0];
      const rows=[...document.querySelectorAll('tr,[role="row"]')];
      const row=rows.find(r=>[...r.querySelectorAll('td,[role="cell"]')]
        .some(c=>(c.innerText||c.textContent||'').trim()===id));
      if(!row) return 'row-missing';
      const controls=[...row.querySelectorAll('button,a,[role="button"]')]
        .filter(b=>{const r=b.getBoundingClientRect(); return r.width>0 && r.height>0;});
      if(!controls.length) return 'button-missing';
      const blue=controls.find(b=>{
        const c=getComputedStyle(b).backgroundColor;
        return /primary|blue/i.test(String(b.className)) || /rgb\(\s*\d+\s*,\s*(?:\d+\s*,\s*\d+)/.test(c) &&
          (c.includes('0, 123') || c.includes('13, 110') || c.includes('0, 112'));
      });
      (blue||controls[0]).click();
      return blue ? 'clicked-blue' : 'clicked-first';
    """
    deadline = time.time() + WAIT_TIMEOUT
    last_state = "row-missing"
    while time.time() < deadline:
        last_state = driver.execute_script(script, applicant_id)
        if last_state.startswith("clicked"):
            logger.info("Clicked case %s Action control (%s).", applicant_id, last_state)
            return
        time.sleep(.25)
    raise RuntimeError(f"Could not click Action for Applicant ID {applicant_id}: {last_state}")


def get_candidate_name(driver, applicant_id):
    """Read Candidate Name from the exact filtered row before opening QC."""
    value = driver.execute_script("""
      const id=arguments[0];
      const row=[...document.querySelectorAll('tr,[role="row"]')].find(r=>
        [...r.querySelectorAll('td,[role="cell"]')].some(c=>(c.innerText||c.textContent||'').trim()===id));
      if(!row) return '';
      const cells=[...row.querySelectorAll('td,[role="cell"]')];
      const table=row.closest('table');
      const headers=table ? [...table.querySelectorAll('thead th')].map(h=>(h.innerText||'').trim().toLowerCase()) : [];
      let index=headers.findIndex(h=>h.replace(/\\s+/g,' ').includes('candidate name'));
      if(index<0) index=2; // table screenshot: #, Agent Name, Candidate Name
      return cells[index] ? (cells[index].innerText||cells[index].textContent||'').trim() : '';
    """, applicant_id)
    return value or ""


def safe_folder_part(value):
    """Make a Windows-safe folder name component without losing the name."""
    value = " ".join((value or "").split())
    for char in '<>:"/\\|?*':
        value = value.replace(char, "_")
    return value.strip(" .") or "Unknown Candidate"


def search_case_id(driver, applicant_id):
    """Use the table's search box to filter to one exact Applicant ID."""
    search_box = table_search_box(driver)
    if search_box is None:
        raise RuntimeError("The New Upload table is visible, but its search input could not be identified.")
    search_box.clear()
    search_box.send_keys(applicant_id)
    # Prefer the visible magnifier control; use Enter only if the field has
    # no adjacent search button. Firing both can cause two Angular refreshes
    # and detach the result row before its Action click.
    clicked_search = False
    try:
        search_button = driver.execute_script("""
            const input=arguments[0];
            let p=input;
            for(let i=0;i<4 && p;i++,p=p.parentElement){
              const buttons=[...p.querySelectorAll('button,a,[role="button"]')]
                .filter(b=>{const r=b.getBoundingClientRect(); return r.width>0 && r.height>0;});
              const b=buttons.find(b=>/search|fa-search|magnif/i.test((b.innerText||'')+' '+(b.getAttribute('aria-label')||'')+' '+b.className));
              if(b) return b;
            }
            return null;
        """, search_box)
        if search_button is not None:
            try_click(driver, search_button)
            clicked_search = True
    except StaleElementReferenceException:
        pass
    if not clicked_search:
        # Reacquire the field because the DOM may have refreshed between
        # typing and submitting.
        search_box = table_search_box(driver)
        if search_box is None:
            raise RuntimeError("The New Upload search field disappeared before search submission.")
        search_box.send_keys(Keys.RETURN)
    try:
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            lambda d: find_candidate_row(d, applicant_id) is not None
            or any(term in (d.find_element(By.TAG_NAME, "body").text or "").lower()
                   for term in ("no matching", "no records", "no data"))
        )
    except TimeoutException:
        pass
    if find_candidate_row(driver, applicant_id) is None:
        raise RuntimeError(f"Applicant ID {applicant_id} was not found by the candidate list search.")
    logger.info("Search matched Applicant ID %s.", applicant_id)
 
 
def apply_date_filter(driver):
    """Select the portal's Month-to-Date preset before opening New Uploads."""
    mtd_xpath = (
        "//*[self::button or self::a or @role='button']"
        "[normalize-space(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='mtd' "
        "or contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'mtd')]"
    )
    try:
        button = wait_for(driver, By.XPATH, mtd_xpath, timeout=15, clickable=True)
    except TimeoutException as exc:
        screenshot = os.path.join(os.getcwd(), f"date_filter_diagnostics_{RUN_TS}.png")
        try:
            driver.save_screenshot(screenshot)
        except Exception:
            screenshot = "(screenshot could not be saved)"
        raise RuntimeError(
            f"Could not find the portal's MTD date preset. Screenshot: {screenshot}"
        ) from exc
    try_click(driver, button)
    logger.info("Selected the portal's MTD (Month to Date) date preset.")
    time.sleep(2)
# STEP 3 - COLLECT CURRENT PAGE'S CASE ROWS
# ===========================
 
def get_new_upload_case_ids(driver):
    """The portal opens the dedicated 'Conventional New Upload' table."""
    rows = visible_case_rows(driver)
    logger.info("Found %d case row(s) on this New Upload page.", len(rows))
    return len(rows)
 
 
def get_blue_action_button(row):
    # The QC action is the blue action control in the far-right Action column.
    buttons = row.find_elements(By.XPATH, ".//button | .//a | .//*[@role='button']")
    if not buttons:
        raise RuntimeError("COULD NOT FIND any action button in this row.")
    for button in buttons:
        try:
            class_name = (button.get_attribute("class") or "").lower()
            color = button.value_of_css_property("background-color")
            if any(term in class_name for term in ("primary", "blue")) or color in (
                "rgb(0, 123, 255)", "rgb(13, 110, 253)", "rgb(0, 112, 192)"
            ):
                return button
        except StaleElementReferenceException:
            continue
    # Screenshot shows blue first and orange second; retain that documented
    # ordering as the fallback when the site uses custom theme colors.
    return buttons[0]


def visible_case_rows(driver):
    rows = driver.find_elements(By.XPATH, "//table//tbody/tr[td] | //*[@role='row'][.//*[@role='cell']]")
    if not rows:
        rows = driver.find_elements(By.XPATH, "//table//tr[td] | //tr[td]")
    result = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td | .//*[@role='cell']")
            row_text = (row.text or "").strip().lower()
            if row.is_displayed() and len(cells) > 1 and row_text and "no data" not in row_text:
                result.append(row)
        except StaleElementReferenceException:
            continue
    return result


def go_to_next_page(driver):
    """Advance one paginator page; return False when Next is disabled/absent."""
    old_first = visible_case_rows(driver)
    old_text = old_first[0].text if old_first else ""
    next_controls = driver.find_elements(By.XPATH,
        "//*[self::button or self::a or @role='button'][translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='next']")
    next_control = next((el for el in next_controls if el.is_displayed()), None)
    if next_control is None:
        return False
    disabled = (next_control.get_attribute("disabled") is not None
                or (next_control.get_attribute("aria-disabled") or "").lower() == "true"
                or "disabled" in (next_control.get_attribute("class") or "").lower())
    if disabled:
        return False
    driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", next_control)
    try_click(driver, next_control)
    try:
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            lambda d: bool(visible_case_rows(d)) and visible_case_rows(d)[0].text != old_text
        )
    except TimeoutException:
        if visible_case_rows(driver) and visible_case_rows(driver)[0].text == old_text:
            return False
    logger.info("Advanced to next New Upload page.")
    return True
 
 
def get_case_id_from_row(row):
    # The screenshot labels the stable identifier as Applicant ID.
    cells = row.find_elements(By.XPATH, "./td | .//*[@role='cell']")
    headers = row.find_elements(By.XPATH, "ancestor::table[1]//th")
    for index, header in enumerate(headers):
        if "applicant id" in (header.text or "").strip().lower() and index < len(cells):
            value = cells[index].text.strip()
            if value:
                return value
    for cell in cells:
        text = cell.text.strip()
        if text.isdigit() and len(text) >= 6:
            return text
    return f"row_{abs(hash(row.text))}"  # fallback, not ideal but keeps the run going
 
 
# ===========================
# STEP 4 - ON THE QC PAGE, CLICK DOWNLOAD
# ===========================
 
def download_case_zip(driver):
    """
    Clicks "Download All Data And Document" on the QC page, waits for
    the zip to finish downloading into DOWNLOAD_DIR, and returns the
    downloaded file's path.
    """
    before_files = set(os.listdir(DOWNLOAD_DIR))
 
    btn = wait_for(driver, By.XPATH,
        "//*[self::button or self::a or @role='button'][contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'download all data') and contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'document')]",
        clickable=True)
    try_click(driver, btn)
    logger.info("Clicked 'Download All Data And Document'.")
 
    # Poll until a new, fully-downloaded (.zip, not .crdownload) file appears.
    deadline = time.time() + 120
    new_file = None
    while time.time() < deadline:
        current_files = set(os.listdir(DOWNLOAD_DIR))
        new_names = current_files - before_files
        finished = [f for f in new_names if f.lower().endswith(".zip")]
        still_downloading = [f for f in new_names if f.endswith(".crdownload")]
        if finished and not still_downloading:
            new_file = finished[0]
            break
        time.sleep(.25)
 
    if new_file is None:
        raise RuntimeError("Zip download did not finish within 120 seconds.")
 
    return os.path.join(DOWNLOAD_DIR, new_file)
 
 
# ===========================
# STEP 5 - UNZIP, DELETE ZIP, MOVE INTO New_Cases
# ===========================
 
def unzip_delete_move(zip_path, case_id, candidate_name=""):
    # Use a unique staging folder; the final folder uses the candidate's
    # actual name and Applicant ID, matching manual extraction in Explorer.
    extract_to = tempfile.mkdtemp(prefix=f".extract_{case_id}_", dir=DOWNLOAD_DIR)
 
    with zipfile.ZipFile(zip_path, "r") as zf:
        root = os.path.abspath(extract_to)
        for member in zf.infolist():
            target = os.path.abspath(os.path.join(root, member.filename))
            if os.path.commonpath([root, target]) != root:
                raise RuntimeError(f"Unsafe path in downloaded zip: {member.filename}")
        zf.extractall(extract_to)
 
    os.remove(zip_path)  # delete the zip as requested
 
    # If the zip contains one top-level folder, move that folder as the final
    # folder; otherwise move the extracted files into the named case folder.
    entries = os.listdir(extract_to)
    if len(entries) == 1 and os.path.isdir(os.path.join(extract_to, entries[0])):
        source_dir = os.path.join(extract_to, entries[0])
    else:
        source_dir = extract_to

    dest_name = f"{safe_folder_part(candidate_name)}_{case_id}" if candidate_name else str(case_id)
    dest_dir = os.path.join(NEW_CASES_DIR, dest_name)
    if os.path.exists(dest_dir):
        suffix = 2
        while os.path.exists(os.path.join(NEW_CASES_DIR, f"{dest_name}_{suffix}")):
            suffix += 1
        dest_dir = os.path.join(NEW_CASES_DIR, f"{dest_name}_{suffix}")
 
    shutil.move(source_dir, dest_dir)
 
    # clean up the now-empty extract_to wrapper, if anything's left
    if os.path.isdir(extract_to):
        shutil.rmtree(extract_to, ignore_errors=True)
 
    logger.info("Moved case %s files into: %s", case_id, dest_dir)
    return dest_dir


def process_case(driver, case_id, main_window):
    """Open the exact case row's blue QC action, download and file its ZIP."""
    list_url = driver.current_url
    candidate_name = get_candidate_name(driver, case_id)
    if not candidate_name:
        logger.warning("Candidate Name was not readable for %s; folder will use Applicant ID only.", case_id)
    windows_before = set(driver.window_handles)
    click_case_action(driver, case_id)

    new_windows = set(driver.window_handles) - windows_before
    if new_windows:
        driver.switch_to.window(new_windows.pop())
    try:
        zip_path = download_case_zip(driver)
        unzip_delete_move(zip_path, case_id, candidate_name)
    finally:
        if driver.current_window_handle != main_window:
            driver.close()
            driver.switch_to.window(main_window)
        elif driver.current_url != list_url:
            driver.back()
            WebDriverWait(driver, WAIT_TIMEOUT).until(
                lambda d: d.current_url == list_url or table_search_box(d) is not None
            )


def cleanup_temporary_download_folder():
    """Remove this run's generated download directory when it is empty."""
    if not AUTO_DOWNLOAD_DIR:
        return
    expected = os.path.abspath(os.path.join(os.getcwd(), f"Portal_Downloads_{RUN_TS}"))
    if os.path.normcase(DOWNLOAD_DIR) != os.path.normcase(expected):
        logger.warning("Temporary download folder was custom; leaving it in place: %s", DOWNLOAD_DIR)
        return
    if not os.path.isdir(DOWNLOAD_DIR):
        return
    try:
        # ZIPs are removed after extraction. rmdir intentionally refuses to
        # remove a folder containing a partial/failed download.
        os.rmdir(DOWNLOAD_DIR)
        logger.info("Removed empty temporary download folder: %s", DOWNLOAD_DIR)
    except OSError:
        logger.warning(
            "Temporary download folder still contains files, so it was kept for recovery: %s",
            DOWNLOAD_DIR,
        )
 
 
# ===========================
# MAIN
# ===========================
 
def main():
    print(
        "Copy the Applicant ID cells from Excel and paste them here. "
        "Comma-separated IDs also work. Press Enter on a blank line when done."
    )
    pasted_lines = []
    while True:
        line = input("Applicant IDs: " if not pasted_lines else "             ")
        if not line.strip():
            break
        pasted_lines.append(line.strip())
    raw_ids = "\n".join(pasted_lines)
    requested_ids = {
        item.strip()
        for item in raw_ids.replace(";", ",").replace("\t", ",").replace("\n", ",").split(",")
        if item.strip()
    }
    if not requested_ids:
        raise RuntimeError("No Applicant IDs entered; nothing to download.")
    logger.info("Requested Applicant IDs: %s", sorted(requested_ids))

    driver = build_driver()
    processed, failed = [], []
 
    completed = False
    try:
        login(driver)
        apply_date_filter(driver)
        open_new_uploads_list(driver)

        main_window = driver.current_window_handle
        for applicant_id in sorted(requested_ids):
            logger.info("--- Searching Applicant ID %s ---", applicant_id)
            try:
                open_new_uploads_list(driver)
                search_case_id(driver, applicant_id)
                process_case(driver, applicant_id, main_window)
                processed.append(applicant_id)
            except Exception as e:
                logger.exception("FAILED Applicant ID %s", applicant_id)
                failed.append((applicant_id, str(e)))
        completed = True
    except Exception:
        screenshot = os.path.join(os.getcwd(), f"run_failure_{RUN_TS}.png")
        try:
            driver.save_screenshot(screenshot)
            logger.exception("Downloader stopped with an error. Chrome will stay open. Screenshot: %s", screenshot)
        except Exception:
            logger.exception("Downloader stopped with an error. Chrome will stay open; failure screenshot could not be saved.")
        raise
    finally:
        if completed:
            # With detach enabled, explicitly close Chrome after a clean run.
            try:
                for handle in list(driver.window_handles):
                    driver.switch_to.window(handle)
                    driver.close()
            finally:
                driver.quit()
        else:
            # Ending the WebDriver session leaves the detached Chrome window open.
            try:
                driver.quit()
            except Exception:
                logger.exception("Could not release the failed WebDriver session.")
 
    logger.info("=" * 60)
    logger.info("Processed %d case(s) successfully: %s", len(processed), processed)
    if failed:
        logger.error("%d case(s) failed:", len(failed))
        for cid, err in failed:
            logger.error("  - %s: %s", cid, err)
    logger.info("Final folder: %s", NEW_CASES_DIR)
    cleanup_temporary_download_folder()
    logger.info("Full log saved to: %s", LOG_FILE)
 
 
if __name__ == "__main__":
    main()





















