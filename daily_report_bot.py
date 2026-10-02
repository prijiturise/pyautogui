import pyautogui
import pyperclip
import win32com.client
import win32gui
import win32con

import subprocess
import time
import re
from datetime import datetime
from pathlib import Path


# ============================================================
# CONFIGURATION
# ============================================================

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.35

NSE_URL = "https://www.nseindia.com/market-data/live-equity-market"

# IMPORTANT:
# This is automatically the folder where this Python file lives.
BASE_FOLDER = Path(__file__).resolve().parent

# Your requested output folder
OUTPUT_FOLDER = BASE_FOLDER


# ============================================================
# UTILITY
# ============================================================

def wait(seconds):
    time.sleep(seconds)


# ============================================================
# OPEN CHROME + NSE
# ============================================================

def open_nse():

    print()
    print("=" * 65)
    print("STEP 1 - OPENING NSE")
    print("=" * 65)

    print("Opening Chrome...")

    subprocess.Popen(
        [
            "cmd",
            "/c",
            "start",
            "",
            "chrome",
            NSE_URL
        ]
    )

    wait(10)

    print("Loading NSE page...")

    # Make sure Chrome has focus
    pyautogui.hotkey("ctrl", "l")

    wait(1)

    pyautogui.write(
        NSE_URL,
        interval=0.01
    )

    pyautogui.press("enter")

    print("Waiting for NSE...")

    wait(12)

    print("NSE should now be loaded.")


# ============================================================
# READ NIFTY 50 FROM SCREEN
# ============================================================

def get_nifty_data():

    print()
    print("=" * 65)
    print("STEP 2 - READING NIFTY 50 FROM NSE")
    print("=" * 65)

    print("Selecting NSE page...")

    pyautogui.hotkey(
        "ctrl",
        "a"
    )

    wait(1)

    print("Copying page text...")

    pyautogui.hotkey(
        "ctrl",
        "c"
    )

    wait(2)

    page_text = pyperclip.paste()

    if not page_text:
        raise Exception(
            "Could not copy NSE page text."
        )

    # Diagnostic file
    debug_file = OUTPUT_FOLDER / "nse_debug.txt"

    debug_file.write_text(
        page_text,
        encoding="utf-8"
    )

    print(
        "NSE data copied."
    )

    print(
        "Debug file:",
        debug_file
    )

    # --------------------------------------------------------
    # Locate NIFTY 50
    # --------------------------------------------------------

    position = page_text.find(
        "NIFTY 50"
    )

    if position == -1:

        raise Exception(
            "NIFTY 50 was not found on the NSE page.\n"
            "Check nse_debug.txt."
        )

    section = page_text[
        position:
        position + 1500
    ]

    # --------------------------------------------------------
    # Extract decimal numbers
    # --------------------------------------------------------

    pattern = re.compile(
        r"\d{1,3}(?:,\d{3})*\.\d{1,2}"
    )

    matches = pattern.findall(
        section
    )

    # Convert plausible NIFTY prices
    prices = []

    for value in matches:

        try:

            number = float(
                value.replace(",", "")
            )

            if 10000 <= number <= 50000:

                prices.append(
                    value
                )

        except ValueError:

            continue

    if len(prices) < 5:

        raise Exception(
            "Could not extract enough NIFTY values.\n"
            "Check nse_debug.txt."
        )

    # NSE row format:
    #
    # NIFTY 50
    # OPEN
    # HIGH
    # LOW
    # PREV CLOSE
    # LTP

    open_price = prices[0]
    high_price = prices[1]
    low_price = prices[2]
    previous_close = prices[3]
    ltp = prices[4]

    # --------------------------------------------------------
    # Calculate change
    # --------------------------------------------------------

    previous_number = float(
        previous_close.replace(",", "")
    )

    ltp_number = float(
        ltp.replace(",", "")
    )

    change = (
        ltp_number -
        previous_number
    )

    percent_change = (
        change /
        previous_number
    ) * 100

    point_change = f"{change:+.2f}"
    percent_text = f"{percent_change:+.2f}%"

    print()
    print("NIFTY 50 DATA")
    print("-" * 40)

    print(
        "Open           :",
        open_price
    )

    print(
        "High           :",
        high_price
    )

    print(
        "Low            :",
        low_price
    )

    print(
        "Previous Close :",
        previous_close
    )

    print(
        "LTP            :",
        ltp
    )

    print(
        "Point Change   :",
        point_change
    )

    print(
        "% Change       :",
        percent_text
    )

    return {
        "open": open_price,
        "high": high_price,
        "low": low_price,
        "previous_close": previous_close,
        "ltp": ltp,
        "change": point_change,
        "percent": percent_text
    }


# ============================================================
# START A COMPLETELY NEW EXCEL INSTANCE
# ============================================================

def start_excel():

    print()
    print("=" * 65)
    print("STEP 3 - STARTING A CLEAN EXCEL INSTANCE")
    print("=" * 65)

    print(
        "Closing old Excel instances..."
    )

    # We intentionally close Excel instances started
    # by previous failed runs.
    #
    # IMPORTANT:
    # Save your own Excel documents before running the bot.
    subprocess.run(
        [
            "taskkill",
            "/F",
            "/IM",
            "EXCEL.EXE"
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    wait(3)

    print(
        "Starting a fresh Excel instance..."
    )

    # DispatchEx creates a separate Excel instance.
    excel = win32com.client.DispatchEx(
        "Excel.Application"
    )

    excel.Visible = True
    excel.DisplayAlerts = False

    wait(4)

    print(
        "Creating blank workbook..."
    )

    workbook = excel.Workbooks.Add()

    wait(4)

    worksheet = workbook.Worksheets(1)

    # --------------------------------------------------------
    # Bring Excel to the front
    # --------------------------------------------------------

    try:

        hwnd = excel.Hwnd

        win32gui.ShowWindow(
            hwnd,
            win32con.SW_RESTORE
        )

        win32gui.SetForegroundWindow(
            hwnd
        )

    except Exception:

        pass

    wait(2)

    print(
        "Blank Excel workbook is ready."
    )

    return excel, workbook, worksheet


# ============================================================
# ENTER DATA USING PYAUTOGUI
# ============================================================

def enter_data(data):

    print()
    print("=" * 65)
    print("STEP 4 - ENTERING REPORT USING PYAUTOGUI")
    print("=" * 65)

    # Generate timestamp at runtime.
    now = datetime.now()

    timestamp = now.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    print(
        "Runtime date/time:",
        timestamp
    )

    # --------------------------------------------------------
    # Build spreadsheet content
    # --------------------------------------------------------

    report = (
        "Date & Time\t"
        "Index\t"
        "NIFTY 50 Value\t"
        "Point Change\t"
        "% Change\t"
        "Comment\n"

        f"{timestamp}\t"
        "NIFTY 50\t"
        f"{data['ltp']}\t"
        f"{data['change']}\t"
        f"{data['percent']}\t"
        "Monitor market movement"
    )

    print()
    print("Pasting report into Excel...")

    # Clipboard allows the entire row to be pasted
    # reliably without hundreds of individual keystrokes.
    pyperclip.copy(
        report
    )

    # Go to A1
    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    wait(1)

    # Paste
    pyautogui.hotkey(
        "ctrl",
        "v"
    )

    wait(4)

    print(
        "Data entered into worksheet."
    )


# ============================================================
# FORMAT USING PYAUTOGUI
# ============================================================

def format_excel():

    print()
    print("=" * 65)
    print("STEP 5 - FORMATTING WORKSHEET")
    print("=" * 65)

    # Go to A1
    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    wait(1)

    # Select row 1
    pyautogui.hotkey(
        "shift",
        "space"
    )

    wait(1)

    # Bold
    pyautogui.hotkey(
        "ctrl",
        "b"
    )

    wait(1)

    # Return to A1
    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    # Select columns A:F
    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    # Auto-fit used columns using Excel shortcut:
    # select used region and auto-fit columns
    pyautogui.hotkey(
        "ctrl",
        "a"
    )

    wait(1)

    # Alt+H,O,I is Excel's AutoFit Column Width command.
    # Unlike Alt+D, this is specifically an Excel
    # Home command and does not navigate the Save dialog.
    pyautogui.hotkey(
        "alt",
        "h"
    )

    wait(0.5)

    pyautogui.press(
        "o"
    )

    wait(0.5)

    pyautogui.press(
        "i"
    )

    wait(2)

    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    print(
        "Formatting completed."
    )


# ============================================================
# SAVE THROUGH EXCEL COM
# ============================================================

def save_excel(excel, workbook):

    print()
    print("=" * 65)
    print("STEP 6 - SAVING EXCEL")
    print("=" * 65)

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    target_file = (
        OUTPUT_FOLDER /
        f"daily_report_{today}.xlsx"
    )

    print()
    print("Target folder:")
    print(OUTPUT_FOLDER)

    print()
    print("Target filename:")
    print(target_file.name)

    print()
    print("Full target path:")
    print(target_file)

    # --------------------------------------------------------
    # If today's file already exists, remove it.
    # --------------------------------------------------------

    if target_file.exists():

        try:

            target_file.unlink()

            print(
                "Existing report removed."
            )

        except Exception as error:

            raise Exception(
                "Could not replace existing report:\n"
                + str(error)
            )

    wait(2)

    print()
    print("Saving directly to target path...")

    # xlOpenXMLWorkbook = 51
    workbook.SaveAs(
        str(target_file),
        FileFormat=51
    )

    wait(5)

    # --------------------------------------------------------
    # Verify
    # --------------------------------------------------------

    if target_file.exists():

        print()
        print("=" * 65)
        print("EXCEL FILE SAVED SUCCESSFULLY")
        print("=" * 65)

        print(
            target_file
        )

        return target_file

    raise Exception(
        "Excel SaveAs completed but file "
        "was not found:\n"
        + str(target_file)
    )


# ============================================================
# TAKE FINAL SCREENSHOT
# ============================================================

def take_screenshot():

    print()
    print("=" * 65)
    print("STEP 7 - TAKING FINAL SCREENSHOT")
    print("=" * 65)

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    screenshot_file = (
        OUTPUT_FOLDER /
        f"daily_report_{today}.png"
    )

    # Make sure Excel is in front.
    wait(2)

    print(
        "Capturing Excel worksheet..."
    )

    pyautogui.screenshot(
        str(screenshot_file)
    )

    wait(2)

    if screenshot_file.exists():

        print()
        print(
            "Screenshot saved:"
        )

        print(
            screenshot_file
        )

        return screenshot_file

    raise Exception(
        "Screenshot could not be verified."
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 65)
    print("NSE DAILY MARKET REPORT BOT")
    print("=" * 65)

    print()
    print("Python script:")
    print(Path(__file__).resolve())

    print()
    print("Output folder:")
    print(OUTPUT_FOLDER)

    print()
    print(
        "IMPORTANT:"
    )

    print(
        "Do not move the mouse/keyboard while the bot runs."
    )

    print()

    # --------------------------------------------------------
    # STEP 1 + 2
    # --------------------------------------------------------

    open_nse()

    data = get_nifty_data()

    # --------------------------------------------------------
    # STEP 3
    # --------------------------------------------------------

    excel = None
    workbook = None

    try:

        excel, workbook, worksheet = start_excel()

        # ----------------------------------------------------
        # STEP 4
        # ----------------------------------------------------

        enter_data(data)

        # ----------------------------------------------------
        # STEP 5
        # ----------------------------------------------------

        format_excel()

        # ----------------------------------------------------
        # STEP 6
        # ----------------------------------------------------

        target_file = save_excel(
            excel,
            workbook
        )

        # ----------------------------------------------------
        # STEP 7
        # ----------------------------------------------------

        screenshot_file = (
            take_screenshot()
        )

        # ----------------------------------------------------
        # COMPLETE
        # ----------------------------------------------------

        print()
        print()
        print("=" * 65)
        print("                 REPORT COMPLETE")
        print("=" * 65)

        print()
        print("Excel file:")
        print(target_file)

        print()
        print("Screenshot:")
        print(screenshot_file)

        print()
        print("Both files have been verified.")

        print()
        print("=" * 65)

        # Keep Excel open so you can inspect it.
        print()
        print(
            "Excel will remain open for inspection."
        )

    except Exception as error:

        print()
        print("=" * 65)
        print("                  ERROR")
        print("=" * 65)

        print()
        print(
            str(error)
        )

        print()
        print(
            "The Excel window will remain open."
        )

        # Diagnostic screenshot
        try:

            diagnostic = (
                OUTPUT_FOLDER /
                "bot_error.png"
            )

            pyautogui.screenshot(
                str(diagnostic)
            )

            print()
            print(
                "Diagnostic screenshot:"
            )

            print(
                diagnostic
            )

        except Exception:

            pass


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        print()
        print()
        print(
            "BOT STOPPED BY USER."
        )

    except Exception as error:

        print()
        print(
            "FATAL ERROR:"
        )

        print(
            error
        )
