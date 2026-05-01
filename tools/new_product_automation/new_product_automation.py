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
- The tool fills one Excel row, clicks submit/save in iWMS, then continues to
  the next row until the Excel file is complete.
"""

from __future__ import annotations

from pathlib import Path
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


APP_DIR = Path(__file__).resolve().parent
PROFILE_DIR = APP_DIR / "wms_browser_profile"

# TODO: Put the real WMS product-edit page URL here.
WMS_URL = "https://iwms.us.jdlglobal.com/#/newGoodsAdd"

# Fill every row in the uploaded Excel.
MAX_ROWS: int | None = None

SUBMIT_BUTTON_PATTERN = re.compile(r"^(\u63d0\u4ea4|\u4fdd\u5b58|Submit|Save)$", re.IGNORECASE)
CONFIRM_BUTTON_PATTERN = re.compile(r"^(\u786e\u5b9a|\u786e\u8ba4|OK|Confirm)$", re.IGNORECASE)

SKU_COLUMN = "\u5546\u54c1\u7f16\u7801"
LENGTH_COLUMN = "\u957f"
WIDTH_COLUMN = "\u5bbd"
HEIGHT_COLUMN = "\u9ad8"
WEIGHT_COLUMN = "\u91cd\u91cf"

REQUIRED_COLUMNS = [
    ("sku", SKU_COLUMN),
    ("length", LENGTH_COLUMN),
    ("width", WIDTH_COLUMN),
    ("height", HEIGHT_COLUMN),
]

FIELD_LABELS = {
    "sku": "\u5546\u54c1\u7f16\u7801/\u5305\u88c5\u4ee3\u7801/\u8017\u6750\u7f16\u7801",
    "length": "\u957f",
    "width": "\u5bbd",
    "height": "\u9ad8",
    "weight": "\u91cd\u91cf",
}

COLORS = {
    "bg": "#f7f8fb",
    "card": "#ffffff",
    "border": "#e6eaf0",
    "text": "#172033",
    "muted": "#667085",
    "primary": "#2563eb",
    "primary_dark": "#1d4ed8",
    "navy": "#111827",
    "success_bg": "#ecfdf3",
    "success_text": "#027a48",
    "warning_bg": "#fff7ed",
    "warning_text": "#b45309",
}


def clean_value(value: object) -> str:
    """Convert Excel values to text suitable for filling into the web form."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


class WorkSightAutomationApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("WorkSight Pro - New Product Automation")
        self.root.geometry("860x640")
        self.root.minsize(760, 560)
        self.root.configure(bg=COLORS["bg"])

        self.excel_path: Path | None = None
        self.ready_event = threading.Event()
        self.worker: threading.Thread | None = None

        self.file_text = tk.StringVar(value="No Excel file selected")
        self.status_text = tk.StringVar(value="Choose an Excel file to begin.")
        self.rows_text = tk.StringVar(
            value="Required columns: product code, length, width, height"
        )

        self._build_ui()

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["bg"])
        outer.pack(fill="both", expand=True, padx=28, pady=24)

        header = tk.Frame(outer, bg=COLORS["bg"])
        header.pack(fill="x", pady=(0, 18))
        logo = tk.Label(
            header,
            text="▥",
            bg=COLORS["bg"],
            fg=COLORS["primary"],
            font=("Segoe UI", 28, "bold"),
        )
        logo.pack(side="left", padx=(0, 12))
        title_box = tk.Frame(header, bg=COLORS["bg"])
        title_box.pack(side="left")
        tk.Label(
            title_box,
            text="WorkSight Pro",
            bg=COLORS["bg"],
            fg=COLORS["navy"],
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_box,
            text="New Product Maintenance Automation",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
        ).pack(anchor="w")

        card = self._card(outer)
        card.pack(fill="x", pady=(0, 14))
        tk.Label(
            card,
            text="新品维护自动化",
            bg=COLORS["card"],
            fg=COLORS["navy"],
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            card,
            text="Upload Excel, open iWMS locally, then fill and submit each product row automatically.",
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=("Segoe UI", 11),
            wraplength=760,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

        grid = tk.Frame(outer, bg=COLORS["bg"])
        grid.pack(fill="x", pady=(0, 14))

        upload = self._card(grid)
        upload.pack(side="left", fill="both", expand=True, padx=(0, 8))
        tk.Label(upload, text="1. Upload Excel", bg=COLORS["card"], fg=COLORS["navy"], font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(upload, textvariable=self.rows_text, bg=COLORS["card"], fg=COLORS["muted"], font=("Segoe UI", 10), wraplength=350, justify="left").pack(anchor="w", pady=(8, 12))
        file_bar = tk.Frame(upload, bg="#f8fafc", highlightbackground=COLORS["border"], highlightthickness=1)
        file_bar.pack(fill="x", pady=(0, 12))
        tk.Label(file_bar, textvariable=self.file_text, bg="#f8fafc", fg=COLORS["text"], font=("Segoe UI", 10), anchor="w").pack(fill="x", padx=10, pady=9)
        tk.Button(upload, text="Choose File", command=self.choose_file, **self._ghost_button()).pack(side="left")

        start = self._card(grid)
        start.pack(side="left", fill="both", expand=True, padx=(8, 0))
        tk.Label(start, text="2. Start Automation", bg=COLORS["card"], fg=COLORS["navy"], font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(start, text="First-time login requires your iWMS account and password. After you confirm the form is ready, the tool will fill and submit every Excel row.", bg=COLORS["card"], fg=COLORS["muted"], font=("Segoe UI", 10), wraplength=350, justify="left").pack(anchor="w", pady=(8, 12))
        button_row = tk.Frame(start, bg=COLORS["card"])
        button_row.pack(fill="x")
        self.start_button = tk.Button(button_row, text="Start", command=self.start, **self._primary_button())
        self.start_button.pack(side="left", padx=(0, 8))
        self.ready_button = tk.Button(button_row, text="Initial page is ready", command=self.mark_form_ready, state="disabled", **self._ghost_button())
        self.ready_button.pack(side="left")

        status = self._card(outer)
        status.pack(fill="both", expand=True)
        tk.Label(status, text="Status", bg=COLORS["card"], fg=COLORS["navy"], font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(status, textvariable=self.status_text, bg=COLORS["warning_bg"], fg=COLORS["warning_text"], font=("Segoe UI", 10, "bold"), anchor="w").pack(fill="x", pady=(10, 10), ipady=8, padx=0)
        self.log_box = tk.Text(
            status,
            height=12,
            bg="#f8fafc",
            fg=COLORS["text"],
            relief="flat",
            wrap="word",
            font=("Consolas", 10),
        )
        self.log_box.pack(fill="both", expand=True)
        self.log("Ready. Choose an Excel file, then click Start.")

    def _card(self, parent: tk.Misc) -> tk.Frame:
        frame = tk.Frame(parent, bg=COLORS["card"], padx=18, pady=16, highlightbackground=COLORS["border"], highlightthickness=1)
        return frame

    def _primary_button(self) -> dict[str, object]:
        return {
            "bg": COLORS["primary"],
            "fg": "white",
            "activebackground": COLORS["primary_dark"],
            "activeforeground": "white",
            "relief": "flat",
            "bd": 0,
            "padx": 18,
            "pady": 9,
            "font": ("Segoe UI", 10, "bold"),
            "cursor": "hand2",
        }

    def _ghost_button(self) -> dict[str, object]:
        return {
            "bg": "white",
            "fg": COLORS["text"],
            "activebackground": "#f8fafc",
            "activeforeground": COLORS["text"],
            "relief": "solid",
            "bd": 1,
            "padx": 16,
            "pady": 8,
            "font": ("Segoe UI", 10, "bold"),
            "cursor": "hand2",
        }

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        self.excel_path = Path(path)
        self.file_text.set(str(self.excel_path))
        self.status_text.set("Excel selected. Click Start when ready.")
        self.log(f"Selected Excel: {self.excel_path}")

    def start(self) -> None:
        if self.excel_path is None:
            messagebox.showwarning("Missing Excel", "Please choose an Excel file first.")
            return
        if self.worker and self.worker.is_alive():
            return
        self.ready_event.clear()
        self.start_button.configure(state="disabled")
        self.status_text.set("Starting local browser...")
        self.worker = threading.Thread(target=run_automation, args=(self, self.excel_path), daemon=True)
        self.worker.start()

    def wait_until_form_ready(self, url: str) -> None:
        self.post_status(
            "iWMS is open. Log in if needed, then confirm the initial new-product page is ready."
        )
        self.log(f"Opened iWMS: {url}")
        self.root.after(0, lambda: self.ready_button.configure(state="normal"))
        self.ready_event.wait()
        self.root.after(0, lambda: self.ready_button.configure(state="disabled"))

    def mark_form_ready(self) -> None:
        self.log("User confirmed product form is ready.")
        self.ready_event.set()

    def post_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_text.set(text))

    def log(self, text: str) -> None:
        def append() -> None:
            self.log_box.insert("end", f"{text}\n")
            self.log_box.see("end")
        self.root.after(0, append)

    def finish(self) -> None:
        self.post_status("Finished all rows.")
        self.log("Finished all rows.")
        self.root.after(0, lambda: self.start_button.configure(state="normal"))
        self.root.after(0, lambda: messagebox.showinfo("Finished", "Finished all rows."))

    def fail(self, error: Exception) -> None:
        self.post_status("Automation failed. Check the message below.")
        self.log(f"ERROR: {error}")
        self.root.after(0, lambda: self.start_button.configure(state="normal"))
        self.root.after(0, lambda: messagebox.showerror("Automation failed", str(error)))

    def run(self) -> None:
        self.root.mainloop()


def load_rows(excel_path: Path) -> pd.DataFrame:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path.resolve()}")

    df = pd.read_excel(excel_path)
    df.columns = [str(column).strip() for column in df.columns]

    missing_columns = [
        column_name for _, column_name in REQUIRED_COLUMNS if column_name not in df.columns
    ]
    if missing_columns:
        available = ", ".join(df.columns)
        missing = ", ".join(missing_columns)
        raise ValueError(f"Missing Excel columns: {missing}. Available columns: {available}")

    if MAX_ROWS is not None:
        df = df.head(MAX_ROWS)

    return df


def find_input_after_label(page: Page, field_key: str, timeout: int = 3000):
    label = FIELD_LABELS[field_key]

    # Best case: the page wires the visible label to the input with for/id.
    try:
        locator = page.get_by_label(label, exact=False)
        locator.wait_for(state="visible", timeout=timeout)
        return locator
    except PlaywrightTimeoutError:
        pass

    # Common WMS/Ant Design layout: text label is near the next input. The
    # colon may be a real full-width character or only CSS, so match loosely.
    input_after_label_xpath = (
        "xpath=//*[self::label or self::span or self::div or self::p]"
        f"[contains(normalize-space(.), '{label}')]"
        "/following::input[not(@type='hidden')][1]"
    )
    locator = page.locator(input_after_label_xpath).first
    try:
        locator.wait_for(state="visible", timeout=timeout)
        return locator
    except PlaywrightTimeoutError:
        pass

    raise RuntimeError(
        f"Could not find input for field {field_key!r} with label {label!r}. "
        "The page structure is different from the screenshot."
    )


def fill_input_after_label(page: Page, field_key: str, value: str, timeout: int = 3000) -> None:
    find_input_after_label(page, field_key, timeout=timeout).fill(value, timeout=timeout)
    print(f"  {field_key}: filled")


def search_sku_then_wait_for_details(page: Page, sku: str) -> None:
    sku_input = find_input_after_label(page, "sku", timeout=15000)

    # Use real keyboard-style input here because iWMS listens for Enter on the
    # focused field before it renders the logistics section.
    sku_input.click(timeout=5000)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.type(sku, delay=25)
    page.keyboard.press("Enter")

    # After Enter, iWMS loads the lower logistics section. Wait for its fields.
    try:
        find_input_after_label(page, "length", timeout=12000)
    except RuntimeError:
        # Some builds only react to a second Enter after the input event settles.
        sku_input.click(timeout=5000)
        page.keyboard.press("Enter")
        find_input_after_label(page, "length", timeout=12000)
    page.wait_for_timeout(800)


def fill_dimensions(page: Page, row: pd.Series) -> None:
    values = {
        field_key: clean_value(row[column_name])
        for field_key, column_name in REQUIRED_COLUMNS
    }

    fill_input_after_label(page, "length", values["length"], timeout=5000)
    fill_input_after_label(page, "width", values["width"], timeout=5000)
    fill_input_after_label(page, "height", values["height"], timeout=5000)

    if WEIGHT_COLUMN in row.index:
        weight = clean_value(row[WEIGHT_COLUMN])
        if weight:
            try:
                fill_input_after_label(page, "weight", weight, timeout=2000)
            except RuntimeError:
                pass


def click_button_by_pattern(page: Page, pattern: re.Pattern[str], timeout: int = 5000) -> bool:
    buttons = page.locator("button:not([disabled])")
    count = min(buttons.count(), 20)
    for index in range(count):
        button = buttons.nth(index)
        is_debug_button = button.evaluate(
            """element => {
                const id = element.id || "";
                const className = String(element.className || "");
                return id === "debugSubmit"
                    || className.toLowerCase().includes("debug")
                    || Boolean(element.closest("[class*='debug'], #debugMenu, .debugTools"));
            }"""
        )
        if is_debug_button:
            continue

        text = button.inner_text(timeout=1000).strip()
        if pattern.search(text):
            button.click(timeout=timeout)
            return True

    return False


def submit_current_row(page: Page) -> None:
    if not click_button_by_pattern(page, SUBMIT_BUTTON_PATTERN):
        raise RuntimeError(
            "Could not find the submit/save button. Expected button text: Submit, Save, or Chinese submit/save."
        )

    page.wait_for_timeout(800)

    # Some WMS pages show a confirmation dialog after clicking submit/save.
    if click_button_by_pattern(page, CONFIRM_BUTTON_PATTERN, timeout=1500):
        page.wait_for_timeout(800)

    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except PlaywrightTimeoutError:
        pass


def run_automation(app: WorkSightAutomationApp, excel_path: Path) -> None:
    context = None
    try:
        df = load_rows(excel_path)
        app.post_status(f"Loaded {len(df)} row(s). Opening iWMS...")
        app.log(f"Loaded {len(df)} row(s) from {excel_path}.")

        with sync_playwright() as playwright:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=PROFILE_DIR,
                headless=False,
                viewport={"width": 1600, "height": 900},
            )
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(WMS_URL)

            app.wait_until_form_ready(page.url)

            for row_number, row in df.iterrows():
                sku = clean_value(row[SKU_COLUMN])
                app.post_status(f"Searching SKU {row_number + 1} of {len(df)}: {sku}")
                app.log(f"Opening new-product page for row {row_number + 1}: {sku}")

                page.goto(WMS_URL)
                find_input_after_label(page, "sku", timeout=15000)

                app.log("Entering product code and pressing Enter.")
                search_sku_then_wait_for_details(page, sku)

                app.post_status(f"Filling dimensions {row_number + 1} of {len(df)}: {sku}")
                app.log("Logistics fields loaded. Filling length, width, and height.")
                fill_dimensions(page, row)

                app.post_status(f"Submitting row {row_number + 1} of {len(df)}: {sku}")
                app.log("Submitting this row in iWMS.")
                submit_current_row(page)
                app.log(f"Submitted row {row_number + 1}: {sku}")

            context.close()
            app.finish()
    except Exception as error:
        if context:
            try:
                context.close()
            except Exception:
                pass
        app.fail(error)


def main() -> None:
    WorkSightAutomationApp().run()


if __name__ == "__main__":
    main()
