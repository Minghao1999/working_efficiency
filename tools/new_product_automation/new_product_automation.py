"""
Demo: upload an Excel file, then fill WMS product logistics fields with
Playwright.

This version does not control your real mouse or keyboard. Playwright opens its
own browser window and types directly into the web page.

First run:
1. Run this script.
2. Choose an Excel file in the small upload window.
3. Click Start.
4. If WMS asks you to log in, log in manually in the browser that opens.
5. Navigate to the product-edit page if needed.
6. Click OK in the confirmation popup to start filling.

Later runs:
- The browser profile in ./wms_browser_profile keeps your login state, so you
  usually do not need to log in every time.

Expected columns: product code, length, width, height, weight.
The default Excel headers are Chinese; they are stored below as Unicode escapes
so this file stays readable in Windows terminals.

Important:
- This demo only fills fields.
- It does not click submit/save.
- Keep MAX_ROWS = 1 while testing.
"""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


PROFILE_DIR = Path("wms_browser_profile")

# TODO: Put the real WMS product-edit page URL here.
WMS_URL = "https://iwms.us.jdlglobal.com/#/newGoodsAdd"

# Use 1 for a safe demo. Change to None after you are confident.
MAX_ROWS: int | None = 1

# Pause after each filled row so you can inspect the page.
WAIT_AFTER_EACH_ROW_MS = 3000

SKU_COLUMN = "\u5546\u54c1\u7f16\u7801"
LENGTH_COLUMN = "\u957f"
WIDTH_COLUMN = "\u5bbd"
HEIGHT_COLUMN = "\u9ad8"
WEIGHT_COLUMN = "\u91cd\u91cf"

FIELD_COLUMNS = [
    ("sku", SKU_COLUMN),
    ("length", LENGTH_COLUMN),
    ("width", WIDTH_COLUMN),
    ("height", HEIGHT_COLUMN),
    ("weight", WEIGHT_COLUMN),
]

FIELD_LABELS = {
    "sku": "\u5546\u54c1\u7f16\u7801/\u5305\u88c5\u4ee3\u7801/\u8017\u6750\u7f16\u7801",
    "length": "\u957f",
    "width": "\u5bbd",
    "height": "\u9ad8",
    "weight": "\u91cd\u91cf",
}


def clean_value(value: object) -> str:
    """Convert Excel values to text suitable for filling into the web form."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def choose_excel_file() -> Path | None:
    selected_file: dict[str, Path | None] = {"path": None}

    root = tk.Tk()
    root.title("WMS Excel Upload")
    root.geometry("520x180")
    root.resizable(False, False)

    file_label = tk.StringVar(value="No Excel file selected")

    def browse_file() -> None:
        path = filedialog.askopenfilename(
            title="Choose Excel file",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
        if path:
            selected_file["path"] = Path(path)
            file_label.set(path)

    def start() -> None:
        if selected_file["path"] is None:
            messagebox.showwarning("Missing Excel", "Please choose an Excel file first.")
            return
        root.destroy()

    def cancel() -> None:
        selected_file["path"] = None
        root.destroy()

    tk.Label(root, text="Choose Excel file, then click Start").pack(
        anchor="w", padx=18, pady=(18, 8)
    )
    tk.Label(root, textvariable=file_label, anchor="w", relief="sunken").pack(
        fill="x", padx=18
    )

    button_frame = tk.Frame(root)
    button_frame.pack(anchor="e", padx=18, pady=18)
    tk.Button(button_frame, text="Choose Excel", width=14, command=browse_file).pack(
        side="left", padx=(0, 8)
    )
    tk.Button(button_frame, text="Start", width=10, command=start).pack(
        side="left", padx=(0, 8)
    )
    tk.Button(button_frame, text="Cancel", width=10, command=cancel).pack(side="left")

    root.mainloop()
    return selected_file["path"]


def load_rows(excel_path: Path) -> pd.DataFrame:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path.resolve()}")

    df = pd.read_excel(excel_path)
    df.columns = [str(column).strip() for column in df.columns]

    missing_columns = [
        column_name for _, column_name in FIELD_COLUMNS if column_name not in df.columns
    ]
    if missing_columns:
        available = ", ".join(df.columns)
        missing = ", ".join(missing_columns)
        raise ValueError(f"Missing Excel columns: {missing}. Available columns: {available}")

    if MAX_ROWS is not None:
        df = df.head(MAX_ROWS)

    return df


def wait_until_user_is_ready(page: Page) -> None:
    print(f"Opened: {page.url}")
    print("If WMS asks you to log in, finish login in the opened browser.")
    print("Navigate to the product-edit page if the script is not already there.")

    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Ready to start",
        "Log in to WMS and open the product-edit page.\n\n"
        "Click OK when the form is ready, and the script will start filling.",
    )
    root.destroy()


def show_done_message() -> None:
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Finished",
        "Demo finished.\n\nNo submit/save action was performed.",
    )
    root.destroy()


def fill_input_after_label(page: Page, field_key: str, value: str) -> None:
    label = FIELD_LABELS[field_key]

    # Best case: the page wires the visible label to the input with for/id.
    try:
        page.get_by_label(label, exact=False).fill(value, timeout=3000)
        print(f"  {field_key}: filled by label")
        return
    except PlaywrightTimeoutError:
        pass

    # Common WMS/Ant Design layout: text label is near the next input. The
    # colon may be a real full-width character or only CSS, so match loosely.
    input_after_label_xpath = (
        "xpath=//*[self::label or self::span or self::div or self::p]"
        f"[contains(normalize-space(.), '{label}')]"
        "/following::input[not(@type='hidden')][1]"
    )
    try:
        page.locator(input_after_label_xpath).first.fill(value, timeout=3000)
        print(f"  {field_key}: filled after nearby text")
        return
    except PlaywrightTimeoutError:
        pass

    raise RuntimeError(
        f"Could not find input for field {field_key!r} with label {label!r}. "
        "The page structure is different from the screenshot."
    )


def fill_one_row(page: Page, row: pd.Series) -> None:
    values = {
        field_key: clean_value(row[column_name])
        for field_key, column_name in FIELD_COLUMNS
    }

    # Fill in the form. No submit/save click here.
    fill_input_after_label(page, "sku", values["sku"])
    fill_input_after_label(page, "length", values["length"])
    fill_input_after_label(page, "width", values["width"])
    fill_input_after_label(page, "height", values["height"])
    fill_input_after_label(page, "weight", values["weight"])


def main() -> None:
    excel_path = choose_excel_file()
    if excel_path is None:
        print("Cancelled. No Excel file selected.")
        return

    df = load_rows(excel_path)
    print(f"Loaded {len(df)} row(s) from {excel_path}.")

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            viewport={"width": 1600, "height": 900},
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(WMS_URL)

        wait_until_user_is_ready(page)

        for row_number, row in df.iterrows():
            sku = clean_value(row[SKU_COLUMN])
            print(f"Filling row {row_number + 1}: {sku}")

            fill_one_row(page, row)

            print("Filled only. No submit/save click. Please check the page.")
            page.wait_for_timeout(WAIT_AFTER_EACH_ROW_MS)

        print("Demo finished. No submit/save action was performed.")
        show_done_message()
        context.close()


if __name__ == "__main__":
    main()
