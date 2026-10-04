import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

try:
    import xlsxwriter
except ImportError:
    raise SystemExit(
        "Missing package: XlsxWriter\n\n"
        "Install it with:\n"
        "python -m pip install XlsxWriter"
    )

APP_TITLE = "Directory Health Checker"

TIMESTAMP_RE = re.compile(
    r"^(?P<timestamp>\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M)\s*:\s*(?P<message>.*)$",
    re.IGNORECASE,
)

# Warning patterns that should make the current store WARNING.
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

# These indicate a clearly unhealthy store.
NOT_HEALTHY_PATTERNS = [
    "not healthy",
    "is unhealthy",
    "database validity failed",
    "invalid database",
    "lost database file",
]

KNOWN_BRANDS = {"KR", "SBC", "PB", "LA", "TORI", "RJ", "BS"}


def parse_line(line):
    line = line.strip()
    match = TIMESTAMP_RE.match(line)

    if not match:
        return None, line

    raw_ts = match.group("timestamp")
    message = match.group("message").strip()

    try:
        ts = datetime.strptime(raw_ts, "%m/%d/%Y %I:%M:%S %p")
    except ValueError:
        ts = None

    return ts, message


def extract_store_from_path(message):
    """
    Examples:
      \\server\Aloha Dated Folder\KR\kr_chinatown\20261003.rar
      \\server\Aloha Dated Folder\SBC\sbc_nepomall\20261003 IS HEALTHY

    Returns:
      ('KR', 'kr_chinatown')
    """
    normalized = message.replace("/", "\\")
    parts = [p.strip() for p in normalized.split("\\") if p.strip()]

    for i, part in enumerate(parts):
        brand = part.upper()

        if brand in KNOWN_BRANDS and i + 1 < len(parts):
            store = parts[i + 1].strip()

            # Ignore dated folder / archive name.
            if re.fullmatch(r"\d{8}(?:\.rar)?", store, re.IGNORECASE):
                return None

            # Strip accidental trailing status words.
            store = re.sub(r"\s+IS\s+HEALTHY.*$", "", store, flags=re.IGNORECASE)
            store = re.sub(r"\s+IS\s+UNHEALTHY.*$", "", store, flags=re.IGNORECASE)
            store = store.replace(".rar", "").strip()

            if store:
                return brand, store

    return None


def status_priority(status):
    # Higher number wins if same timestamp.
    return {
        "HEALTHY": 1,
        "WARNING": 2,
        "NOT HEALTHY": 3,
    }.get(status, 0)


def update_store(store_results, brand, store, timestamp, status, reason):
    key = (brand.upper(), store.lower())

    new_result = {
        "timestamp": timestamp,
        "brand": brand.upper(),
        "store": store,
        "brand_store": f"{brand.upper()} - {store}",
        "status": status,
        "reason": reason,
        "source": "Extractor Log",
    }

    existing = store_results.get(key)

    if existing is None:
        store_results[key] = new_result
        return

    old_ts = existing["timestamp"] or datetime.min
    new_ts = timestamp or datetime.min

    if new_ts > old_ts:
        store_results[key] = new_result
    elif new_ts == old_ts and status_priority(status) >= status_priority(existing["status"]):
        store_results[key] = new_result


def finalize_current_store(current, store_results):
    """
    Finalize one Extracting... block.

    Rules:
    - Explicit NOT HEALTHY -> NOT HEALTHY
    - Warning found -> WARNING with exact warning message(s)
    - IS HEALTHY -> HEALTHY
    - Done found without health message -> HEALTHY
    - Extracting started but never completed -> NOT HEALTHY
    """
    if not current:
        return

    brand = current["brand"]
    store = current["store"]
    ts = current["last_timestamp"] or current["start_timestamp"]

    if current["not_healthy_reasons"]:
        status = "NOT HEALTHY"
        reason = " | ".join(dict.fromkeys(current["not_healthy_reasons"]))

    elif current["warning_reasons"]:
        status = "WARNING"
        reason = " | ".join(dict.fromkeys(current["warning_reasons"]))

    elif current["is_healthy"]:
        status = "HEALTHY"
        reason = ""

    elif current["done"]:
        status = "HEALTHY"
        reason = ""

    else:
        status = "NOT HEALTHY"
        reason = "Extraction started but did not complete."

    update_store(
        store_results,
        brand,
        store,
        ts,
        status,
        reason,
    )


def analyze_log(file_path):
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    store_results = {}
    current = None
    active_brand = None

    for raw_line in lines:
        ts, message = parse_line(raw_line)
        msg_lower = message.lower()
        msg_upper = message.upper()

        # Keep track of active brand sections.
        compare_match = re.search(r"COMPARING DATABASES\s+([A-Z]+)", msg_upper)
        if compare_match:
            active_brand = compare_match.group(1)

        # Start of a store extraction.
        if "EXTRACTING..." in msg_upper:
            # Finalize previous store first.
            finalize_current_store(current, store_results)

            extracted = extract_store_from_path(message)

            if extracted:
                brand, store = extracted
            else:
                brand = active_brand or "UNKNOWN"
                store = "Unknown Store"

            current = {
                "brand": brand,
                "store": store,
                "start_timestamp": ts,
                "last_timestamp": ts,
                "is_healthy": False,
                "done": False,
                "warning_reasons": [],
                "not_healthy_reasons": [],
            }
            continue

        # If no current store, ignore store-specific status lines.
        if current is None:
            continue

        if ts:
            current["last_timestamp"] = ts

        # Explicit health line.
        if "IS HEALTHY" in msg_upper:
            extracted = extract_store_from_path(message)

            if extracted:
                brand, store = extracted

                # If health line refers to another store, finalize current first.
                if (
                    brand.upper() != current["brand"].upper()
                    or store.lower() != current["store"].lower()
                ):
                    finalize_current_store(current, store_results)
                    current = {
                        "brand": brand,
                        "store": store,
                        "start_timestamp": ts,
                        "last_timestamp": ts,
                        "is_healthy": True,
                        "done": False,
                        "warning_reasons": [],
                        "not_healthy_reasons": [],
                    }
                else:
                    current["is_healthy"] = True
            else:
                current["is_healthy"] = True

        # Explicit unhealthy signals.
        if (
            "NOT HEALTHY" in msg_upper
            or "IS UNHEALTHY" in msg_upper
            or "DATABASE VALIDITY FAILED" in msg_upper
            or "INVALID DATABASE" in msg_upper
        ):
            current["not_healthy_reasons"].append(message)

        # "Lost database file" should only be unhealthy if it is not the normal
        # "Checking for Lost Database file" informational line.
        if "lost database file" in msg_lower and "checking for" not in msg_lower:
            current["not_healthy_reasons"].append(message)

        # Warning signals.
        for pattern in WARNING_PATTERNS:
            if pattern in msg_lower:
                current["warning_reasons"].append(message)
                break

        # Completed individual store extraction.
        if msg_upper.startswith("DONE ") or " DONE \\\\" in msg_upper:
            extracted = extract_store_from_path(message)

            if extracted:
                brand, store = extracted
                if (
                    brand.upper() == current["brand"].upper()
                    and store.lower() == current["store"].lower()
                ):
                    current["done"] = True
            else:
                current["done"] = True

        # Brand completion means current store section is over.
        if re.search(r"=+\s*DONE\s+[A-Z]+\s*=+", msg_upper):
            finalize_current_store(current, store_results)
            current = None

    # End of file.
    finalize_current_store(current, store_results)

    rows = list(store_results.values())

    # Sort by brand then store.
    rows.sort(key=lambda r: (r["brand"], r["store"].lower()))

    return rows


def export_excel(rows, output_path):
    workbook = xlsxwriter.Workbook(output_path)
    ws = workbook.add_worksheet("Monitor Report")

    title_fmt = workbook.add_format({
        "bold": True,
        "font_color": "white",
        "bg_color": "#1F4E78",
        "font_size": 16,
        "align": "center",
        "valign": "vcenter",
    })

    header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    })

    normal_fmt = workbook.add_format({
        "border": 1,
        "valign": "vcenter",
    })

    wrap_fmt = workbook.add_format({
        "border": 1,
        "valign": "vcenter",
        "text_wrap": True,
    })

    timestamp_fmt = workbook.add_format({
        "border": 1,
        "valign": "vcenter",
        "num_format": "mm/dd/yyyy h:mm AM/PM",
    })

    healthy_fmt = workbook.add_format({
        "bold": True,
        "font_color": "#006100",
        "bg_color": "#C6EFCE",
        "border": 1,
        "valign": "vcenter",
    })

    warning_fmt = workbook.add_format({
        "bold": True,
        "font_color": "#9C6500",
        "bg_color": "#FFEB9C",
        "border": 1,
        "valign": "vcenter",
    })

    not_healthy_fmt = workbook.add_format({
        "bold": True,
        "font_color": "#9C0006",
        "bg_color": "#FFC7CE",
        "border": 1,
        "valign": "vcenter",
    })

    ws.merge_range("A1:E1", "DIRECTORY HEALTH CHECKER", title_fmt)
    ws.set_row(0, 28)

    headers = [
        "Timestamp",
        "Brand / Store",
        "Status",
        "Reason",
        "Source / Process",
    ]

    for col, header in enumerate(headers):
        ws.write(2, col, header, header_fmt)

    for row_num, row in enumerate(rows, start=3):
        if row["timestamp"]:
            ws.write_datetime(row_num, 0, row["timestamp"], timestamp_fmt)
        else:
            ws.write(row_num, 0, "", normal_fmt)

        ws.write(row_num, 1, row["brand_store"], normal_fmt)

        if row["status"] == "HEALTHY":
            status_fmt = healthy_fmt
        elif row["status"] == "WARNING":
            status_fmt = warning_fmt
        else:
            status_fmt = not_healthy_fmt

        ws.write(row_num, 2, row["status"], status_fmt)
        ws.write(row_num, 3, row["reason"], wrap_fmt)
        ws.write(row_num, 4, row["source"], normal_fmt)

    ws.set_column("A:A", 23)
    ws.set_column("B:B", 32)
    ws.set_column("C:C", 18)
    ws.set_column("D:D", 75)
    ws.set_column("E:E", 22)

    ws.freeze_panes(3, 0)

    if rows:
        ws.autofilter(2, 0, 2 + len(rows), 4)

    workbook.close()


class DirectoryHealthCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("850x560")
        self.root.minsize(760, 500)

        self.selected_file = tk.StringVar()

        tk.Label(
            root,
            text="DIRECTORY HEALTH CHECKER",
            font=("Segoe UI", 18, "bold"),
            fg="#1F4E78",
        ).pack(pady=(20, 5))

        tk.Label(
            root,
            text="Select an extractor TXT log. The tool will list the latest status of every store found in the file.",
            font=("Segoe UI", 10),
        ).pack(pady=(0, 18))

        file_frame = tk.Frame(root)
        file_frame.pack(fill="x", padx=25)

        tk.Label(
            file_frame,
            text="Extractor TXT Log:",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        chooser = tk.Frame(file_frame)
        chooser.pack(fill="x", pady=(6, 12))

        tk.Entry(
            chooser,
            textvariable=self.selected_file,
            font=("Segoe UI", 10),
        ).pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=5)

        tk.Button(
            chooser,
            text="Browse...",
            command=self.browse,
            width=12,
        ).pack(side="right")

        buttons = tk.Frame(root)
        buttons.pack(fill="x", padx=25)

        tk.Button(
            buttons,
            text="Analyze & Export Excel",
            command=self.analyze,
            bg="#1F4E78",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            padx=15,
            pady=8,
        ).pack(side="left")

        tk.Button(
            buttons,
            text="Clear",
            command=self.clear,
            padx=15,
            pady=8,
        ).pack(side="left", padx=(10, 0))

        tk.Label(
            root,
            text="Result Preview:",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=25, pady=(15, 5))

        self.preview = scrolledtext.ScrolledText(
            root,
            font=("Consolas", 9),
            height=18,
            wrap=tk.WORD,
        )
        self.preview.pack(fill="both", expand=True, padx=25, pady=(0, 20))
        self.preview.insert(tk.END, "Select a TXT log file to begin.\n")
        self.preview.config(state="disabled")

    def browse(self):
        path = filedialog.askopenfilename(
            title="Select Extractor Log",
            filetypes=[
                ("Text files", "*.txt"),
                ("Log files", "*.log"),
                ("All files", "*.*"),
            ],
        )

        if path:
            self.selected_file.set(path)

    def set_preview(self, text):
        self.preview.config(state="normal")
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, text)
        self.preview.config(state="disabled")

    def clear(self):
        self.selected_file.set("")
        self.set_preview("Select a TXT log file to begin.\n")

    def analyze(self):
        path = self.selected_file.get().strip()

        if not path:
            messagebox.showwarning(APP_TITLE, "Please select a TXT log file.")
            return

        if not os.path.isfile(path):
            messagebox.showerror(APP_TITLE, "The selected file does not exist.")
            return

        try:
            rows = analyze_log(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not analyze the file:\n\n{exc}")
            return

        if not rows:
            messagebox.showwarning(
                APP_TITLE,
                "No store-level extraction paths were found in this log file."
            )
            self.set_preview("No store-level extraction paths were found.\n")
            return

        preview_lines = [
            "DIRECTORY HEALTH CHECKER",
            "=" * 95,
            f"{'Timestamp':22} {'Brand / Store':32} {'Status':12} Reason",
            "-" * 95,
        ]

        for row in rows:
            ts = (
                row["timestamp"].strftime("%m/%d/%Y %I:%M %p")
                if row["timestamp"]
                else ""
            )

            preview_lines.append(
                f"{ts:22} "
                f"{row['brand_store'][:31]:32} "
                f"{row['status']:12} "
                f"{row['reason']}"
            )

        self.set_preview("\n".join(preview_lines))

        default_name = (
            "Directory_Health_Checker_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".xlsx"
        )

        output_path = filedialog.asksaveasfilename(
            title="Save Excel Report",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )

        if not output_path:
            return

        try:
            export_excel(rows, output_path)
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE,
                f"Could not create Excel report:\n\n{exc}"
            )
            return

        healthy = sum(1 for r in rows if r["status"] == "HEALTHY")
        warning = sum(1 for r in rows if r["status"] == "WARNING")
        not_healthy = sum(1 for r in rows if r["status"] == "NOT HEALTHY")

        messagebox.showinfo(
            APP_TITLE,
            "Excel report created successfully.\n\n"
            f"Stores found: {len(rows)}\n"
            f"HEALTHY: {healthy}\n"
            f"WARNING: {warning}\n"
            f"NOT HEALTHY: {not_healthy}\n\n"
            f"{output_path}"
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = DirectoryHealthCheckerApp(root)
    root.mainloop()
