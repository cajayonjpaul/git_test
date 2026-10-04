import os
import re
from datetime import datetime, date
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

try:
    import xlsxwriter
except ImportError:
    raise SystemExit(
        "Missing package: XlsxWriter\n\n"
        "Install it using:\n"
        "python -m pip install XlsxWriter"
    )

APP_TITLE = "Directory Health Checker"

# ============================================================
# CONFIGURATION
# ============================================================

KNOWN_BRANDS = {"KR", "SBC", "PB", "LA", "TORI", "RJ", "BS"}

# Warning messages found in the extractor log.
WARNING_PATTERNS = [
    "could not load file",
    "exception",
    "being used by another process",
    "failed to load",
    "timeout",
    "connection failed",
    "access denied",
    "cannot access",
    "error",
]

# Explicit unhealthy messages.
NOT_HEALTHY_PATTERNS = [
    "not healthy",
    "is unhealthy",
    "database validity failed",
    "invalid database",
]

TIMESTAMP_RE = re.compile(
    r"^(?P<timestamp>\d{2}/\d{2}/\d{4}\s+"
    r"\d{1,2}:\d{2}:\d{2}\s+[AP]M)\s*:\s*(?P<message>.*)$",
    re.IGNORECASE,
)

# ============================================================
# LOG PARSING
# ============================================================

def parse_log_line(line):
    """
    Parses:
    10/03/2026 8:53:10 PM : Extracting... \\server\...\KR\kr_chinatown\20261003.rar
    """
    line = line.strip()
    match = TIMESTAMP_RE.match(line)

    if not match:
        return None, line

    timestamp_text = match.group("timestamp")
    message = match.group("message").strip()

    try:
        timestamp = datetime.strptime(
            timestamp_text,
            "%m/%d/%Y %I:%M:%S %p"
        )
    except ValueError:
        timestamp = None

    return timestamp, message


def extract_store_from_message(message):
    """
    Gets the brand and store name from extractor paths.

    Example:
    \\192.168.100.5\Aloha Dated Folder\KR\kr_chinatown\20261003.rar

    Returns:
    ("KR", "kr_chinatown")
    """
    normalized = message.replace("/", "\\")
    parts = [p.strip() for p in normalized.split("\\") if p.strip()]

    for index, part in enumerate(parts):
        brand = part.upper()

        if brand in KNOWN_BRANDS and index + 1 < len(parts):
            store = parts[index + 1].strip()

            # Ignore a date folder if it accidentally appears immediately after brand.
            if re.fullmatch(r"\d{8}(?:\.rar)?", store, re.IGNORECASE):
                return None

            store = re.sub(
                r"\s+IS\s+HEALTHY.*$",
                "",
                store,
                flags=re.IGNORECASE
            )
            store = re.sub(
                r"\s+(?:IS\s+)?(?:NOT\s+HEALTHY|UNHEALTHY).*$",
                "",
                store,
                flags=re.IGNORECASE
            )

            if store.lower().endswith(".rar"):
                store = store[:-4]

            store = store.strip()

            if store:
                return brand, store

    return None


def status_priority(status):
    """
    Used if two events for the same store have exactly the same timestamp.
    """
    return {
        "HEALTHY": 1,
        "WARNING": 2,
        "NOT HEALTHY": 3,
    }.get(status, 0)


def save_latest_result(results, brand, store, timestamp, status):
    """
    Keeps only the latest status per store inside the chosen date range.
    """
    key = (brand.upper(), store.lower())

    new_result = {
        "brand": brand.upper(),
        "store": store,
        "status": status,
        "timestamp": timestamp,
    }

    old = results.get(key)

    if old is None:
        results[key] = new_result
        return

    old_time = old["timestamp"] or datetime.min
    new_time = timestamp or datetime.min

    if new_time > old_time:
        results[key] = new_result
    elif (
        new_time == old_time
        and status_priority(status) >= status_priority(old["status"])
    ):
        results[key] = new_result


def finalize_store(current_store, results):
    """
    Determines the status of one store extraction block.

    STATUS RULES

    HEALTHY
      - Log contains IS HEALTHY
      - Or store extraction reached Done with no warning / unhealthy signal

    WARNING
      - Store completed but an error/warning message occurred
        such as:
        Could not load file
        Exception
        being used by another process
        timeout
        access denied
        etc.

    NOT HEALTHY
      - Explicit NOT HEALTHY / UNHEALTHY
      - Database validity failed
      - Extraction started but never completed
    """
    if current_store is None:
        return

    timestamp = (
        current_store["last_timestamp"]
        or current_store["start_timestamp"]
    )

    if current_store["not_healthy"]:
        status = "NOT HEALTHY"

    elif current_store["warnings"]:
        status = "WARNING"

    elif current_store["is_healthy"]:
        status = "HEALTHY"

    elif current_store["done"]:
        status = "HEALTHY"

    else:
        status = "NOT HEALTHY"

    save_latest_result(
        results,
        current_store["brand"],
        current_store["store"],
        timestamp,
        status,
    )


def analyze_log(file_path, start_date, end_date):
    """
    Reads the whole TXT file but only processes log lines whose DATE
    is inside the selected inclusive date range.

    Returns one row per store:
        Store Name | Brand | Status

    If the same store appears several times in the selected range,
    the latest status is kept.
    """
    with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
        lines = file.readlines()

    results = {}
    current_store = None

    for raw_line in lines:
        timestamp, message = parse_log_line(raw_line)

        # We need a valid timestamp because the user selected a date range.
        if timestamp is None:
            continue

        log_date = timestamp.date()

        # Inclusive date filtering.
        if log_date < start_date or log_date > end_date:
            continue

        upper_message = message.upper()
        lower_message = message.lower()

        # ----------------------------------------------------
        # New store extraction starts here
        # ----------------------------------------------------
        if "EXTRACTING..." in upper_message:
            # Finish the previous store first.
            finalize_store(current_store, results)

            store_info = extract_store_from_message(message)

            if store_info:
                brand, store = store_info

                current_store = {
                    "brand": brand,
                    "store": store,
                    "start_timestamp": timestamp,
                    "last_timestamp": timestamp,
                    "is_healthy": False,
                    "done": False,
                    "warnings": [],
                    "not_healthy": [],
                }
            else:
                current_store = None

            continue

        # Ignore non-store messages until an Extracting... store is known.
        if current_store is None:
            continue

        current_store["last_timestamp"] = timestamp

        # ----------------------------------------------------
        # HEALTHY
        # ----------------------------------------------------
        if "IS HEALTHY" in upper_message:
            health_store = extract_store_from_message(message)

            # Normally this is the same current store.
            if health_store:
                health_brand, health_store_name = health_store

                if (
                    health_brand == current_store["brand"]
                    and health_store_name.lower() == current_store["store"].lower()
                ):
                    current_store["is_healthy"] = True
            else:
                current_store["is_healthy"] = True

        # ----------------------------------------------------
        # NOT HEALTHY
        # ----------------------------------------------------
        if any(pattern in lower_message for pattern in NOT_HEALTHY_PATTERNS):
            current_store["not_healthy"].append(message)

        # A normal "Checking for Lost Database file" line is informational.
        # Only treat it as unhealthy when it is not the normal checking message.
        if (
            "lost database file" in lower_message
            and "checking for lost database file" not in lower_message
        ):
            current_store["not_healthy"].append(message)

        # ----------------------------------------------------
        # WARNING
        # ----------------------------------------------------
        for warning_pattern in WARNING_PATTERNS:
            if warning_pattern in lower_message:
                current_store["warnings"].append(message)
                break

        # ----------------------------------------------------
        # STORE COMPLETION
        # ----------------------------------------------------
        # Individual-store line, for example:
        # Done \\server\...\KR\kr_chinatown\20261003.rar
        if upper_message.startswith("DONE ") or " DONE \\\\" in upper_message:
            done_store = extract_store_from_message(message)

            if done_store:
                done_brand, done_store_name = done_store

                if (
                    done_brand == current_store["brand"]
                    and done_store_name.lower() == current_store["store"].lower()
                ):
                    current_store["done"] = True
            else:
                current_store["done"] = True

        # Brand-level completion closes the current store block.
        # Example:
        # ====DONE KR====
        if re.search(r"=+\s*DONE\s+[A-Z]+\s*=+", upper_message):
            finalize_store(current_store, results)
            current_store = None

    # Finish last store at end of file.
    finalize_store(current_store, results)

    rows = list(results.values())

    # Final Excel order.
    rows.sort(
        key=lambda item: (
            item["brand"],
            item["store"].lower()
        )
    )

    return rows


# ============================================================
# EXCEL EXPORT
# ============================================================

def export_to_excel(rows, output_path):
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet("Health Report")

    title_format = workbook.add_format({
        "bold": True,
        "font_color": "#FFFFFF",
        "bg_color": "#1F4E78",
        "font_size": 16,
        "align": "center",
        "valign": "vcenter",
    })

    header_format = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    })

    normal_format = workbook.add_format({
        "border": 1,
        "valign": "vcenter",
    })

    healthy_format = workbook.add_format({
        "bold": True,
        "font_color": "#006100",
        "bg_color": "#C6EFCE",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    })

    warning_format = workbook.add_format({
        "bold": True,
        "font_color": "#9C6500",
        "bg_color": "#FFEB9C",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    })

    not_healthy_format = workbook.add_format({
        "bold": True,
        "font_color": "#9C0006",
        "bg_color": "#FFC7CE",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    })

    worksheet.merge_range("A1:C1", "DIRECTORY HEALTH CHECKER", title_format)
    worksheet.set_row(0, 28)

    headers = ["Store Name", "Brand", "Status"]

    for column, header in enumerate(headers):
        worksheet.write(2, column, header, header_format)

    for row_number, row in enumerate(rows, start=3):
        worksheet.write(row_number, 0, row["store"], normal_format)
        worksheet.write(row_number, 1, row["brand"], normal_format)

        if row["status"] == "HEALTHY":
            status_format = healthy_format
        elif row["status"] == "WARNING":
            status_format = warning_format
        else:
            status_format = not_healthy_format

        worksheet.write(
            row_number,
            2,
            row["status"],
            status_format
        )

    worksheet.set_column("A:A", 35)
    worksheet.set_column("B:B", 14)
    worksheet.set_column("C:C", 18)

    worksheet.freeze_panes(3, 0)

    if rows:
        worksheet.autofilter(
            2,
            0,
            2 + len(rows),
            2
        )

    workbook.close()


# ============================================================
# USER INTERFACE
# ============================================================

class DirectoryHealthChecker:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x650")
        self.root.minsize(800, 560)

        self.file_path = tk.StringVar()
        self.start_date = tk.StringVar()
        self.end_date = tk.StringVar()
        self.results = []

        self.build_ui()

    def build_ui(self):
        # ---------------- Header ----------------
        tk.Label(
            self.root,
            text="DIRECTORY HEALTH CHECKER",
            font=("Segoe UI", 19, "bold"),
            fg="#1F4E78",
        ).pack(pady=(20, 4))

        tk.Label(
            self.root,
            text="Select a date range, upload the extractor log, analyze store health, then export to Excel.",
            font=("Segoe UI", 10),
        ).pack(pady=(0, 18))

        # ---------------- Main controls ----------------
        controls = ttk.LabelFrame(
            self.root,
            text="Analysis Settings",
            padding=15,
        )
        controls.pack(fill="x", padx=25, pady=(0, 15))

        # Date range row
        date_frame = tk.Frame(controls)
        date_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            date_frame,
            text="Start Date:",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w")

        tk.Entry(
            date_frame,
            textvariable=self.start_date,
            width=15,
            font=("Segoe UI", 10),
        ).grid(row=1, column=0, sticky="w", padx=(0, 25), ipady=4)

        tk.Label(
            date_frame,
            text="End Date:",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=1, sticky="w")

        tk.Entry(
            date_frame,
            textvariable=self.end_date,
            width=15,
            font=("Segoe UI", 10),
        ).grid(row=1, column=1, sticky="w", padx=(0, 20), ipady=4)

        tk.Label(
            date_frame,
            text="Format: YYYY-MM-DD   Example: 2026-10-03",
            font=("Segoe UI", 9),
            fg="#666666",
        ).grid(row=1, column=2, sticky="w")

        # File uploader row
        tk.Label(
            controls,
            text="Extractor Log File:",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        file_frame = tk.Frame(controls)
        file_frame.pack(fill="x", pady=(5, 0))

        tk.Entry(
            file_frame,
            textvariable=self.file_path,
            font=("Segoe UI", 10),
        ).pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8),
            ipady=5,
        )

        tk.Button(
            file_frame,
            text="Upload TXT File",
            command=self.select_file,
            width=16,
        ).pack(side="right")

        # ---------------- Buttons ----------------
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill="x", padx=25, pady=(0, 12))

        self.analyze_button = tk.Button(
            button_frame,
            text="ANALYZE",
            command=self.run_analysis,
            bg="#1F4E78",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=25,
            pady=9,
        )
        self.analyze_button.pack(side="left")

        self.export_button = tk.Button(
            button_frame,
            text="EXPORT EXCEL",
            command=self.export_results,
            state="disabled",
            font=("Segoe UI", 10, "bold"),
            padx=25,
            pady=9,
        )
        self.export_button.pack(side="left", padx=(10, 0))

        tk.Button(
            button_frame,
            text="Clear",
            command=self.clear,
            padx=18,
            pady=9,
        ).pack(side="left", padx=(10, 0))

        # ---------------- Summary ----------------
        self.summary_label = tk.Label(
            self.root,
            text="No analysis yet.",
            font=("Segoe UI", 10, "bold"),
            anchor="w",
        )
        self.summary_label.pack(
            fill="x",
            padx=25,
            pady=(4, 8),
        )

        # ---------------- Results preview ----------------
        tk.Label(
            self.root,
            text="Analysis Result:",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=25)

        self.preview = scrolledtext.ScrolledText(
            self.root,
            font=("Consolas", 10),
            height=20,
            wrap=tk.NONE,
        )
        self.preview.pack(
            fill="both",
            expand=True,
            padx=25,
            pady=(5, 20),
        )

        self.preview.insert(
            tk.END,
            "Select a date range and extractor TXT file, then click ANALYZE.\n"
        )
        self.preview.config(state="disabled")

    # --------------------------------------------------------
    # UI helpers
    # --------------------------------------------------------

    def select_file(self):
        selected = filedialog.askopenfilename(
            title="Select Extractor Log",
            filetypes=[
                ("Text files", "*.txt"),
                ("Log files", "*.log"),
                ("All files", "*.*"),
            ],
        )

        if selected:
            self.file_path.set(selected)

    def parse_date_entry(self, value, field_name):
        try:
            return datetime.strptime(
                value.strip(),
                "%Y-%m-%d"
            ).date()
        except ValueError:
            raise ValueError(
                f"{field_name} must use YYYY-MM-DD format.\n"
                f"Example: 2026-10-03"
            )

    def set_preview(self, text):
        self.preview.config(state="normal")
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, text)
        self.preview.config(state="disabled")

    def clear(self):
        self.file_path.set("")
        self.start_date.set("")
        self.end_date.set("")
        self.results = []

        self.summary_label.config(text="No analysis yet.")
        self.export_button.config(state="disabled")

        self.set_preview(
            "Select a date range and extractor TXT file, then click ANALYZE.\n"
        )

    # --------------------------------------------------------
    # Analyze
    # --------------------------------------------------------

    def run_analysis(self):
        selected_file = self.file_path.get().strip()

        if not selected_file:
            messagebox.showwarning(
                APP_TITLE,
                "Please upload an extractor TXT file."
            )
            return

        if not os.path.isfile(selected_file):
            messagebox.showerror(
                APP_TITLE,
                "The selected file does not exist."
            )
            return

        try:
            start = self.parse_date_entry(
                self.start_date.get(),
                "Start Date"
            )

            end = self.parse_date_entry(
                self.end_date.get(),
                "End Date"
            )
        except ValueError as error:
            messagebox.showerror(
                APP_TITLE,
                str(error)
            )
            return

        if start > end:
            messagebox.showerror(
                APP_TITLE,
                "Start Date cannot be later than End Date."
            )
            return

        try:
            self.results = analyze_log(
                selected_file,
                start,
                end
            )
        except Exception as error:
            messagebox.showerror(
                APP_TITLE,
                f"Could not analyze the extractor log:\n\n{error}"
            )
            return

        if not self.results:
            self.export_button.config(state="disabled")

            self.summary_label.config(
                text="No store records found for the selected date range."
            )

            self.set_preview(
                "No store-level extractor records were found between "
                f"{start} and {end}.\n"
            )

            return

        healthy_count = sum(
            1 for row in self.results
            if row["status"] == "HEALTHY"
        )

        warning_count = sum(
            1 for row in self.results
            if row["status"] == "WARNING"
        )

        not_healthy_count = sum(
            1 for row in self.results
            if row["status"] == "NOT HEALTHY"
        )

        self.summary_label.config(
            text=(
                f"Stores: {len(self.results)}    |    "
                f"HEALTHY: {healthy_count}    |    "
                f"WARNING: {warning_count}    |    "
                f"NOT HEALTHY: {not_healthy_count}"
            )
        )

        preview_lines = [
            "DIRECTORY HEALTH CHECKER",
            "=" * 72,
            f"Date Range: {start} to {end}",
            "",
            f"{'STORE NAME':38} {'BRAND':10} {'STATUS':15}",
            "-" * 72,
        ]

        for row in self.results:
            preview_lines.append(
                f"{row['store'][:37]:38} "
                f"{row['brand']:10} "
                f"{row['status']:15}"
            )

        self.set_preview(
            "\n".join(preview_lines)
        )

        self.export_button.config(state="normal")

    # --------------------------------------------------------
    # Export
    # --------------------------------------------------------

    def export_results(self):
        if not self.results:
            messagebox.showwarning(
                APP_TITLE,
                "Please analyze a log file first."
            )
            return

        start_text = self.start_date.get().strip().replace("-", "")
        end_text = self.end_date.get().strip().replace("-", "")

        default_filename = (
            f"Directory_Health_Report_"
            f"{start_text}_to_{end_text}.xlsx"
        )

        output_path = filedialog.asksaveasfilename(
            title="Export Excel Report",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[
                ("Excel Workbook", "*.xlsx")
            ],
        )

        if not output_path:
            return

        try:
            export_to_excel(
                self.results,
                output_path
            )
        except Exception as error:
            messagebox.showerror(
                APP_TITLE,
                f"Could not export the Excel file:\n\n{error}"
            )
            return

        messagebox.showinfo(
            APP_TITLE,
            "Excel export completed successfully.\n\n"
            f"{output_path}"
        )


# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = DirectoryHealthChecker(root)
    root.mainloop()
