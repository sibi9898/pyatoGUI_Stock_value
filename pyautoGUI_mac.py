"""Fetch the TCS quote in Chrome and log it to an Excel workbook."""

from __future__ import annotations

import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk

import pyautogui
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


QUOTE_URL = "https://www.google.com/finance/quote/TCS:NSE"
COMPANY_NAME = "Tata Consultancy Services (TCS)"
PERSONAL_COMMENT = "good to buy"
OUTPUT_DIR = Path(__file__).resolve().parent


def copy_quote_page_text(clipboard: tk.Tk) -> str:
	"""Open the quote page in Chrome and copy its visible text."""
	# macOS's `open` command launches Chrome (or opens a tab if it is running).
	subprocess.run(["open", "-a", "Google Chrome", QUOTE_URL], check=True)
	time.sleep(10)  # allow Chrome and the quote page to load

	clipboard.clipboard_clear()
	clipboard.update()
	pyautogui.hotkey("command", "a")
	pyautogui.hotkey("command", "c")

	# Give the clipboard a moment to receive the copied page text.
	for _ in range(20):
		time.sleep(0.25)
		clipboard.update()
		try:
			copied_text = clipboard.clipboard_get()
		except tk.TclError:
			continue
		if copied_text.strip():
			return copied_text

	raise RuntimeError("Could not copy page text from Chrome.")


def extract_price(page_text: str) -> float:
	"""Extract a rupee-denominated quote from the copied Google Finance page."""
	patterns = (
		r"₹\s*([\d,]+(?:\.\d+)?)",
		r"\bINR\s*([\d,]+(?:\.\d+)?)",
		r"\bRs\.?\s*([\d,]+(?:\.\d+)?)",
	)
	for pattern in patterns:
		match = re.search(pattern, page_text, flags=re.IGNORECASE)
		if match:
			return float(match.group(1).replace(",", ""))

	excerpt = " ".join(page_text.split())[:500]
	raise ValueError(
		"Could not find a rupee quote in the copied page text. "
		f"Check the website or page language. Copied text begins: {excerpt!r}"
	)


def save_workbook(quote_time: datetime, stock_price: float) -> Path:
	"""Create the requested four-entry worksheet and save it for today."""
	workbook = Workbook()
	sheet = workbook.active
	sheet.title = "TCS Stock"
	sheet.append(("Details", "Value"))
	sheet.append(("Current date and time", quote_time))
	sheet.append(("Company name", COMPANY_NAME))
	sheet.append(("Fetched stock price (INR)", stock_price))
	sheet.append(("Personal comment", PERSONAL_COMMENT))

	header_fill = PatternFill("solid", fgColor="17365D")
	for cell in sheet[1]:
		cell.font = Font(color="FFFFFF", bold=True)
		cell.fill = header_fill
		cell.alignment = Alignment(vertical="center")

	sheet[2][1].number_format = "yyyy-mm-dd hh:mm:ss"
	sheet[4][1].number_format = '"₹"#,##0.00'
	sheet.column_dimensions[get_column_letter(1)].width = 32
	sheet.column_dimensions[get_column_letter(2)].width = 42
	for row in sheet.iter_rows(min_row=2):
		for cell in row:
			cell.alignment = Alignment(vertical="center")
	sheet.freeze_panes = "A2"
	sheet.sheet_view.showGridLines = False

	file_date = quote_time.strftime("%Y-%m-%d")
	workbook_path = OUTPUT_DIR / f"TCS_Stock_{file_date}.xlsx"
	workbook.save(workbook_path)
	return workbook_path


def open_excel_and_capture(workbook_path: Path, screenshot_path: Path) -> None:
	"""Open the saved workbook in Microsoft Excel and capture the screen."""
	subprocess.run(
		["open", "-a", "Microsoft Excel", str(workbook_path)], check=True
	)
	time.sleep(10)
	# Use macOS full-screen mode so the worksheet is prominent in the capture.
	pyautogui.hotkey("ctrl", "command", "f")
	time.sleep(1)
	pyautogui.screenshot().save(screenshot_path)


def main() -> None:
	if sys.platform != "darwin":
		raise RuntimeError("This script is for macOS. Run it on a MacBook.")

	pyautogui.FAILSAFE = True
	pyautogui.PAUSE = 0.3
	clipboard = tk.Tk()
	clipboard.withdraw()
	try:
		page_text = copy_quote_page_text(clipboard)
		stock_price = extract_price(page_text)
		quote_time = datetime.now()
		workbook_path = save_workbook(quote_time, stock_price)
		screenshot_path = OUTPUT_DIR / f"TCS_Stock_{quote_time:%Y-%m-%d}.png"
		open_excel_and_capture(workbook_path, screenshot_path)
		print(f"Stock price: ₹{stock_price:,.2f}")
		print(f"Excel workbook: {workbook_path}")
		print(f"Screenshot: {screenshot_path}")
	finally:
		clipboard.destroy()


if __name__ == "__main__":
	main()
