import json
import re
import time
from pathlib import Path
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(r"D:\Gen_AI_Class\Playwright")

EXCEL_FILE = BASE_DIR / "contacts.xlsx"

OUTPUT_DIR = BASE_DIR / "output"

SCREENSHOT_DIR = OUTPUT_DIR / "screenshots"

PROFILE_DIR = BASE_DIR / "whatsapp_profile"

REPORT_JSON = OUTPUT_DIR / "whatsapp_report.json"

REPORT_EXCEL = OUTPUT_DIR / "whatsapp_report.xlsx"

# ------------------------------------------------------------
# IMPORTANT
# True  = test mode, messages WILL NOT be sent
# False = messages WILL be sent
# ------------------------------------------------------------

DRY_RUN = False

# Default message if Message column is blank
DEFAULT_MESSAGE = "Hello {name}, this is a test message."

# Wait time between contacts
WAIT_BETWEEN_CONTACTS = 3

# WhatsApp page load timeout
PAGE_TIMEOUT = 60000


# ============================================================
# CREATE REQUIRED DIRECTORIES
# ============================================================

BASE_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

PROFILE_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def clean_phone(value):
    """
    Convert Excel phone number into a clean format.

    Example:
        +91 98765 43210
        becomes
        +919876543210
    """

    if pd.isna(value):
        return ""

    # Handle Excel numeric values such as 919876543210.0
    if isinstance(value, float) and value.is_integer():
        value = int(value)

    phone = str(value).strip()

    # Remove spaces, brackets, hyphens
    phone = re.sub(r"[\s()\-.]", "", phone)

    return phone


def prepare_message(name, template):
    """
    Replace {name} with actual contact name.
    """

    if pd.isna(template):
        template = ""

    template = str(template).strip()

    if not template:
        template = DEFAULT_MESSAGE

    return template.replace("{name}", name)


def safe_filename(name):
    """
    Make a Windows-safe filename.
    """

    return re.sub(
        r'[^A-Za-z0-9_-]',
        '_',
        name
    )


def save_json_report(results):
    """
    Save current results to JSON.
    """

    with open(
        REPORT_JSON,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            results,
            file,
            indent=4,
            ensure_ascii=False
        )


def save_excel_report(results):
    """
    Save current results to Excel.
    """

    rows = []

    for item in results:

        rows.append({
            "Name": item.get("name", ""),
            "Phone": item.get("phone", ""),
            "Message": item.get("message", ""),
            "Status": item.get("status", ""),
            "Screenshot": item.get("screenshot", ""),
            "Last 3 Messages": json.dumps(
                item.get("last_3_messages", []),
                ensure_ascii=False
            ),
            "Error": item.get("error", ""),
            "Timestamp": item.get("timestamp", "")
        })

    report_df = pd.DataFrame(rows)

    report_df.to_excel(
        REPORT_EXCEL,
        index=False
    )


# ============================================================
# FIND MESSAGE COMPOSER
# ============================================================

def find_message_composer(page):

    """
    Try several possible WhatsApp Web composer selectors.

    WhatsApp Web DOM can change, so we intentionally use
    multiple fallback selectors.
    """

    selectors = [

        # Current/known WhatsApp selectors
        '[data-testid="conversation-compose-box-input"]',

        # Accessibility selectors
        '[contenteditable="true"][aria-label="Type a message"]',

        '[contenteditable="true"][title="Type a message"]',

        # Role-based selector
        '[contenteditable="true"][role="textbox"]',

        # data-tab selector
        '[contenteditable="true"][data-tab="10"]',

        # General contenteditable
        'div[contenteditable="true"][spellcheck="true"]',

        # Last fallback
        '[contenteditable="true"]'
    ]

    print("\nSearching for WhatsApp message composer...")

    for selector in selectors:

        try:

            locator = page.locator(selector)

            count = locator.count()

            print(
                f"Checking: {selector} "
                f"--> {count} found"
            )

            for i in range(count):

                try:

                    candidate = locator.nth(i)

                    if candidate.is_visible():

                        print(
                            "\nSUCCESS - Message composer found"
                        )

                        print(
                            "Selector:",
                            selector
                        )

                        return candidate

                except Exception:
                    continue

        except Exception as error:

            print(
                f"Selector error: {error}"
            )

    return None


# ============================================================
# DEBUG WHATSAPP PAGE
# ============================================================

def debug_whatsapp_page(page):

    """
    Print useful DOM information if composer isn't found.
    """

    print("\n")
    print("=" * 60)
    print("WHATSAPP DEBUG INFORMATION")
    print("=" * 60)

    print("\nCurrent URL:")
    print(page.url)

    print("\nPage title:")
    print(page.title())

    print("\nContenteditable elements:")

    try:

        elements = page.locator(
            '[contenteditable="true"]'
        )

        count = elements.count()

        print(
            "Total:",
            count
        )

        for i in range(count):

            try:

                element = elements.nth(i)

                print("\nElement:", i)

                print(
                    "Visible:",
                    element.is_visible()
                )

                print(
                    "Role:",
                    element.get_attribute("role")
                )

                print(
                    "Title:",
                    element.get_attribute("title")
                )

                print(
                    "Aria-label:",
                    element.get_attribute("aria-label")
                )

                print(
                    "Data-tab:",
                    element.get_attribute("data-tab")
                )

                print(
                    "Data-testid:",
                    element.get_attribute("data-testid")
                )

                print(
                    "Class:",
                    element.get_attribute("class")
                )

            except Exception as error:

                print(
                    "Unable to inspect element:",
                    error
                )

    except Exception as error:

        print(
            "Unable to inspect contenteditable:",
            error
        )

    # Save screenshot
    debug_screenshot = (
        OUTPUT_DIR / "composer_not_found.png"
    )

    try:

        page.screenshot(
            path=str(debug_screenshot),
            full_page=True
        )

        print(
            "\nDiagnostic screenshot:"
        )

        print(
            debug_screenshot
        )

    except Exception as error:

        print(
            "Could not save screenshot:",
            error
        )

    print("=" * 60)


# ============================================================
# EXTRACT LAST 3 INCOMING MESSAGES
# ============================================================

def extract_last_3_messages(page):

    """
    Try to extract the last 3 incoming messages.

    WhatsApp commonly uses message-in for incoming messages.

    We also use data-pre-plain-text to obtain message metadata.
    """

    messages = []

    try:

        # Preferred: incoming messages only
        incoming = page.locator(
            "div.message-in"
        )

        count = incoming.count()

        print(
            f"Incoming message containers found: {count}"
        )

        start_index = max(
            0,
            count - 3
        )

        for i in range(
            start_index,
            count
        ):

            try:

                element = incoming.nth(i)

                text = element.inner_text(
                    timeout=3000
                ).strip()

                if not text:
                    continue

                metadata = element.get_attribute(
                    "data-pre-plain-text"
                )

                messages.append({
                    "direction": "incoming",
                    "metadata": metadata or "",
                    "text": text
                })

            except Exception as error:

                print(
                    "Message extraction error:",
                    error
                )

    except Exception as error:

        print(
            "Incoming message extraction failed:",
            error
        )

    return messages[-3:]


# ============================================================
# WAIT FOR CHAT
# ============================================================

def wait_for_chat(page):

    """
    Wait until the WhatsApp conversation page is loaded.
    """

    print(
        "Waiting for WhatsApp conversation..."
    )

    page.wait_for_timeout(5000)

    # Look for contenteditable elements
    try:

        page.locator(
            '[contenteditable="true"]'
        ).first.wait_for(
            state="visible",
            timeout=30000
        )

        print(
            "WhatsApp conversation appears ready."
        )

        return True

    except PlaywrightTimeoutError:

        print(
            "Composer was not detected during initial wait."
        )

        return False


# ============================================================
# OPEN CHAT
# ============================================================

def open_chat(page, phone):

    """
    Open WhatsApp conversation directly using phone number.

    This avoids relying on WhatsApp search UI.
    """

    digits = phone.replace(
        "+",
        ""
    )

    chat_url = (
        f"https://web.whatsapp.com/send?phone={digits}"
    )

    print(
        "\nOpening:",
        chat_url
    )

    page.goto(
        chat_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT
    )

    page.wait_for_timeout(7000)

    return page.url


# ============================================================
# SEND MESSAGE
# ============================================================

def send_message(page, composer, message):

    """
    Type and send message.
    """

    print(
        "\nMessage:"
    )

    print(
        message
    )

    composer.click()

    page.wait_for_timeout(500)

    composer.fill(
        message
    )

    page.wait_for_timeout(1000)

    if DRY_RUN:

        print(
            "\nDRY RUN ENABLED"
        )

        print(
            "Message was prepared but NOT sent."
        )

        # Clear draft
        composer.fill("")

        return "DRY_RUN_NOT_SENT"

    # Actual send
    print(
        "\nSending message..."
    )

    composer.press(
        "Enter"
    )

    # Allow WhatsApp to process/send
    page.wait_for_timeout(3000)

    print(
        "Message send action completed."
    )

    return "SEND_ATTEMPTED"


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():

    print("\n")
    print("=" * 60)
    print("WHATSAPP PLAYWRIGHT AUTOMATION")
    print("=" * 60)

    print(
        "\nProject folder:",
        BASE_DIR
    )

    print(
        "Excel file:",
        EXCEL_FILE
    )

    print(
        "Dry run:",
        DRY_RUN
    )

    # --------------------------------------------------------
    # CHECK EXCEL FILE
    # --------------------------------------------------------

    if not EXCEL_FILE.exists():

        print(
            "\nERROR:"
        )

        print(
            f"Excel file not found:\n{EXCEL_FILE}"
        )

        return

    # --------------------------------------------------------
    # READ EXCEL
    # --------------------------------------------------------

    print(
        "\nReading contacts.xlsx..."
    )

    try:

        df = pd.read_excel(
            EXCEL_FILE,
            dtype={
                "Phone": str
            }
        )

    except Exception as error:

        print(
            "Unable to read Excel:",
            error
        )

        return

    print(
        "Rows found:",
        len(df)
    )

    # --------------------------------------------------------
    # CHECK REQUIRED COLUMNS
    # --------------------------------------------------------

    required_columns = {
        "Name",
        "Phone"
    }

    missing = (
        required_columns
        - set(df.columns)
    )

    if missing:

        print(
            "\nERROR - Missing Excel columns:"
        )

        print(
            missing
        )

        print(
            "\nRequired columns:"
        )

        print(
            "Name | Phone | Message"
        )

        return

    # Message column optional
    if "Message" not in df.columns:

        df["Message"] = ""

    # Remove rows with missing name/phone
    df = df.dropna(
        subset=[
            "Name",
            "Phone"
        ]
    )

    # --------------------------------------------------------
    # RESULTS
    # --------------------------------------------------------

    results = []

    # --------------------------------------------------------
    # START PLAYWRIGHT
    # --------------------------------------------------------

    with sync_playwright() as p:

        print(
            "\nStarting Chromium..."
        )

        context = p.chromium.launch_persistent_context(

            user_data_dir=str(
                PROFILE_DIR
            ),

            headless=False,

            viewport={
                "width": 1366,
                "height": 900
            },

            args=[
                "--start-maximized"
            ]
        )

        # Get existing page or create one
        if context.pages:

            page = context.pages[0]

        else:

            page = context.new_page()

        # ----------------------------------------------------
        # OPEN WHATSAPP
        # ----------------------------------------------------

        print(
            "\nOpening WhatsApp Web..."
        )

        page.goto(
            "https://web.whatsapp.com",
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT
        )

        print(
            "\n=================================================="
        )

        print(
            "WHATSAPP LOGIN"
        )

        print(
            "=================================================="
        )

        print(
            "\nIf QR code is displayed:"
        )

        print(
            "1. Open WhatsApp on your phone"
        )

        print(
            "2. Go to Settings / Linked Devices"
        )

        print(
            "3. Select Link a Device"
        )

        print(
            "4. Scan the QR code"
        )

        print(
            "\nWait until your WhatsApp chats are visible."
        )

        input(
            "\nAfter WhatsApp is fully logged in, "
            "press ENTER here..."
        )

        # ----------------------------------------------------
        # PROCESS CONTACTS
        # ----------------------------------------------------

        for index, row in df.iterrows():

            name = str(
                row["Name"]
            ).strip()

            phone = clean_phone(
                row["Phone"]
            )

            message = prepare_message(
                name,
                row["Message"]
            )

            print("\n")
            print("=" * 60)

            print(
                f"CONTACT {index + 1}/{len(df)}"
            )

            print(
                "Name:",
                name
            )

            print(
                "Phone:",
                phone
            )

            print("=" * 60)

            record = {

                "name": name,

                "phone": phone,

                "message": message,

                "status": "NOT_STARTED",

                "screenshot": "",

                "last_3_messages": [],

                "error": "",

                "timestamp": datetime.now().isoformat(
                    timespec="seconds"
                )
            }

            try:

                # ------------------------------------------------
                # VALIDATE PHONE
                # ------------------------------------------------

                if not phone:

                    raise ValueError(
                        "Phone number is empty."
                    )

                if not phone.startswith("+"):

                    raise ValueError(
                        "Phone must contain country code "
                        "and start with +."
                    )

                if not phone[1:].isdigit():

                    raise ValueError(
                        "Phone contains invalid characters."
                    )

                # ------------------------------------------------
                # OPEN CHAT
                # ------------------------------------------------

                open_chat(
                    page,
                    phone
                )

                # ------------------------------------------------
                # CHECK PAGE
                # ------------------------------------------------

                chat_ready = wait_for_chat(
                    page
                )

                # ------------------------------------------------
                # FIND COMPOSER
                # ------------------------------------------------

                composer = find_message_composer(
                    page
                )

                if composer is None:

                    debug_whatsapp_page(
                        page
                    )

                    raise RuntimeError(
                        "Message composer not found."
                    )

                # ------------------------------------------------
                # EXTRACT LAST 3 INCOMING MESSAGES
                # ------------------------------------------------

                print(
                    "\nExtracting last 3 incoming messages..."
                )

                last_messages = (
                    extract_last_3_messages(
                        page
                    )
                )

                record[
                    "last_3_messages"
                ] = last_messages

                print(
                    "Messages extracted:",
                    len(last_messages)
                )

                # ------------------------------------------------
                # SEND / DRY RUN
                # ------------------------------------------------

                status = send_message(
                    page,
                    composer,
                    message
                )

                record[
                    "status"
                ] = status

                # ------------------------------------------------
                # SCREENSHOT
                # ------------------------------------------------

                if not DRY_RUN:

                    page.wait_for_timeout(
                        2000
                    )

                    screenshot_name = (
                        f"{index + 1}_"
                        f"{safe_filename(name)}.png"
                    )

                    screenshot_path = (
                        SCREENSHOT_DIR
                        / screenshot_name
                    )

                    page.screenshot(
                        path=str(
                            screenshot_path
                        ),
                        full_page=False
                    )

                    record[
                        "screenshot"
                    ] = str(
                        screenshot_path
                    )

                    print(
                        "\nScreenshot saved:"
                    )

                    print(
                        screenshot_path
                    )

                    # Re-extract messages after sending
                    record[
                        "last_3_messages"
                    ] = extract_last_3_messages(
                        page
                    )

                # ------------------------------------------------
                # SUCCESS
                # ------------------------------------------------

                print(
                    "\nCompleted:",
                    name
                )

            except PlaywrightTimeoutError as error:

                record[
                    "status"
                ] = "TIMEOUT"

                record[
                    "error"
                ] = str(error)

                print(
                    "\nTIMEOUT:",
                    error
                )

            except Exception as error:

                record[
                    "status"
                ] = "ERROR"

                record[
                    "error"
                ] = str(error)

                print(
                    "\nERROR for",
                    name,
                    ":",
                    error
                )

            # Add result
            results.append(
                record
            )

            # ----------------------------------------------------
            # SAVE PROGRESS
            # ----------------------------------------------------

            save_json_report(
                results
            )

            save_excel_report(
                results
            )

            print(
                "\nReports updated."
            )

            # Wait before next contact
            time.sleep(
                WAIT_BETWEEN_CONTACTS
            )

        # --------------------------------------------------------
        # FINAL REPORT
        # --------------------------------------------------------

        save_json_report(
            results
        )

        save_excel_report(
            results
        )

        print("\n")
        print("=" * 60)
        print("AUTOMATION COMPLETED")
        print("=" * 60)

        print(
            "\nJSON report:"
        )

        print(
            REPORT_JSON
        )

        print(
            "\nExcel report:"
        )

        print(
            REPORT_EXCEL
        )

        print(
            "\nScreenshots:"
        )

        print(
            SCREENSHOT_DIR
        )

        input(
            "\nPress ENTER to close the browser..."
        )

        context.close()


# ============================================================
# PROGRAM ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()
