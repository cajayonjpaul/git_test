import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

try:
    import xlsxwriter
except ImportError:
    raise SystemExit(
        "Missing package: xlsxwriter\n\n"
        "Open Command Prompt and run:\n"
        "pip install XlsxWriter"
    )


APP_TITLE = "Extractor Health Monitor"

WARNING_PATTERNS = [
    "could not load file",
    "exception",
    "being used by another process",
    "failed to load",
    "timeout",
    "connection failed",
    "access denied",
]

BRAND_ALIASES = {
    "KR": "KR",
    "SBC": "SBC",
    "PB": "PB",
    "LA": "LA",
    "TORI": "TORI",
    "RJ": "RJ",
    "BS": "BS",
}

TIMESTAMP_RE = re.compile(
    r"^(?P<timestamp>\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M)\s*:\s*(?P<message>.*)$",
    re.IGNORECASE,
)


def parse_timestamp(line):
    match = TIMESTAMP_RE.match(line.strip())
    if not match:
        return None, line.strip()

    raw_ts = match.group("timestamp")
    message = match.group("message").strip()

    try:
        ts = datetime.strptime(raw_ts, "%m/%d/%Y %I:%M:%S %p")
    except ValueError:
        ts = None

    return ts, message


def get_latest_cycle(lines):
    """Return only the latest extractor cycle, beginning at the final 'Start..' line."""
    start_indexes = []

    for i, line in enumerate(lines):
        _, message = parse_timestamp(line)
        if message.lower().startswith("start"):
            start_indexes.append(i)

    if not start_indexes:
        return lines

    return lines[start_indexes[-1]:]


def extract_store_name(path_text):
    """
    Tries to extract brand/store from a path such as:
    \\server\Aloha Dated Folder\KR\kr_quezonave2\20261003
    Returns: KR - kr_quezonave2
    """
    normalized = path_text.replace("/", "\\")
    parts = [p for p in normalized.split("\\") if p]

    for i, part in enumerate(parts):
        upper = part.upper()
        if upper in BRAND_ALIASES and i + 1 < len(parts):
            store = parts[i + 1]
            # Avoid using dated folder / rar filename as store.
            if not re.fullmatch(r"\d{8}(?:\.rar)?", store, flags=re.IGNORECASE):
                return f"{BRAND_ALIASES[upper]} - {store}"

    return None


def detect_brand_from_message(message):
    upper = message.upper()

    compare = re.search(r"COMPARING DATABASES\s+([A-Z]+)", upper)
    if compare:
        code = compare.group(1)
        return BRAND_ALIASES.get(code, code)

    collect = re.search(r"COLLECTING\s+(.+?)\.*$", message, re.IGNORECASE)
    if collect:
        raw = collect.group(1).strip().upper()
        mapping = {
            "BAKESHOP": "BS",
            "RUBY JACKS": "RJ",
            "TORI": "TORI",
            "PB": "PB",
            "SBC": "SBC",
            "KR": "KR",
            "LA": "LA",
        }
        return mapping.get(raw)

    return None


def analyze_log(file_path):
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        all_lines = f.readlines()

    cycle_lines = get_latest_cycle(all_lines)

    rows = []
    active_brand = None
    brand_started = {}
    brand_done = {}
    brand_has_health_row = {}
    warning_rows = []
    explicit_not_healthy = []
    cycle_timestamp = None

    # First timestamp in the selected cycle.
    for line in cycle_lines:
        ts, _ = parse_timestamp(line)
        if ts:
            cycle_timestamp = ts
            break

    for line in cycle_lines:
        ts, message = parse_timestamp(line)
        if not message:
            continue

        detected = detect_brand_from_message(message)
        if detected:
            active_brand = detected
            if "COMPARING DATABASES" in message.upper():
                brand_started[active_brand] = ts or cycle_timestamp
                brand_done.setdefault(active_brand, False)
                brand_has_health_row.setdefault(active_brand, False)

        # Explicit health result.
        if "IS HEALTHY" in message.upper():
            target = extract_store_name(message) or active_brand or "Unknown"
            rows.append({
                "timestamp": ts or cycle_timestamp,
                "brand_store": target,
                "status": "HEALTHY",
                "reason": "",
                "source": "Extractor Log",
            })
            if active_brand:
                brand_has_health_row[active_brand] = True

        if "NOT HEALTHY" in message.upper() or "IS UNHEALTHY" in message.upper():
            target = extract_store_name(message) or active_brand or "Unknown"
            explicit_not_healthy.append({
                "timestamp": ts or cycle_timestamp,
                "brand_store": target,
                "status": "NOT HEALTHY",
                "reason": message,
                "source": "Extractor Log",
            })
            if active_brand:
                brand_has_health_row[active_brand] = True

        # Warnings: preserve the actual log message as the reason.
        lower = message.lower()
        if any(pattern in lower for pattern in WARNING_PATTERNS):
            warning_rows.append({
                "timestamp": ts or cycle_timestamp,
                "brand_store": active_brand or "System",
                "status": "WARNING",
                "reason": message,
                "source": "Extractor Log",
            })

        # Completion marker.
        done_match = re.search(r"=+\s*DONE\s+([A-Z]+)\s*=+", message.upper())
        if done_match:
            done_code = done_match.group(1)
            done_brand = BRAND_ALIASES.get(done_code, done_code)

            # Existing extractor logs sometimes show DONE PB while processing LA.
            # If LA is currently active and DONE PB appears, treat the current section
            # as completed rather than incorrectly marking LA as incomplete.
            if active_brand == "LA" and done_brand == "PB":
                brand_done["LA"] = True
            else:
                brand_done[done_brand] = True

    # Produce one brand-level HEALTHY row when a brand completed successfully
    # and there was no store-level health result for that brand.
    brand_order = ["KR", "SBC", "PB", "LA", "TORI", "RJ", "BS"]

    for brand in brand_order:
        if brand not in brand_started:
            continue

        if brand_done.get(brand):
            if not brand_has_health_row.get(brand):
                rows.append({
                    "timestamp": brand_started.get(brand) or cycle_timestamp,
                    "brand_store": brand,
                    "status": "HEALTHY",
                    "reason": "",
                    "source": "Extractor Log",
                })
        else:
            rows.append({
                "timestamp": brand_started.get(brand) or cycle_timestamp,
                "brand_store": brand,
                "status": "NOT HEALTHY",
                "reason": "Extraction cycle did not complete.",
                "source": "Extractor Log",
            })

    rows.extend(explicit_not_healthy)
    rows.extend(warning_rows)

    if not rows:
        rows.append({
            "timestamp": cycle_timestamp,
            "brand_store": "System",
            "status": "NOT HEALTHY",
            "reason": "No recognizable extractor health information was found in the latest cycle.",
            "source": "Extractor Log",
        })

    # Sort by timestamp, then status priority.
    priority = {"NOT HEALTHY": 0, "WARNING": 1, "HEALTHY": 2}

    rows.sort(
        key=lambda r: (
            r["timestamp"] or datetime.min,
            priority.get(r["status"], 9),
            r["brand_store"],
        )
    )

    return rows


def export_excel(rows, output_path):
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet("Monitor Report")

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

    wrapped_fmt = workbook.add_format({
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

    alt_fmt = workbook.add_format({
        "bg_color": "#DDEBF7",
        "border": 1,
        "valign": "vcenter",
    })

    alt_wrapped_fmt = workbook.add_format({
        "bg_color": "#DDEBF7",
        "border": 1,
        "valign": "vcenter",
        "text_wrap": True,
    })

    worksheet.merge_range("A1:E1", "EXTRACTOR HEALTH MONITOR", title_fmt)
    worksheet.set_row(0, 26)

    headers = ["Timestamp", "Brand / Store", "Status", "Reason", "Source / Process"]
    for col, header in enumerate(headers):
        worksheet.write(2, col, header, header_fmt)

    for i, row in enumerate(rows, start=3):
        alt = (i - 3) % 2 == 0

        ts_fmt = timestamp_fmt
        base_fmt = alt_fmt if alt else normal_fmt
        reason_fmt = alt_wrapped_fmt if alt else wrapped_fmt

        if row["timestamp"]:
            worksheet.write_datetime(i, 0, row["timestamp"], ts_fmt)
        else:
            worksheet.write(i, 0, "", base_fmt)

        worksheet.write(i, 1, row["brand_store"], base_fmt)

        status = row["status"]
        if status == "HEALTHY":
            status_fmt = healthy_fmt
        elif status == "WARNING":
            status_fmt = warning_fmt
        else:
            status_fmt = not_healthy_fmt

        worksheet.write(i, 2, status, status_fmt)
        worksheet.write(i, 3, row["reason"], reason_fmt)
        worksheet.write(i, 4, row["source"], base_fmt)

    worksheet.set_column("A:A", 23)
    worksheet.set_column("B:B", 28)
    worksheet.set_column("C:C", 18)
    worksheet.set_column("D:D", 65)
    worksheet.set_column("E:E", 22)

    worksheet.freeze_panes(3, 0)
    worksheet.autofilter(2, 0, 2 + len(rows), 4)

    workbook.close()


class ExtractorHealthApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("760x520")
        self.root.minsize(700, 460)

        self.selected_file = tk.StringVar()

        title = tk.Label(
            root,
            text="EXTRACTOR HEALTH MONITOR",
            font=("Segoe UI", 18, "bold"),
            fg="#1F4E78",
        )
        title.pack(pady=(20, 8))

        subtitle = tk.Label(
            root,
            text="Test Tool — Select an extractor TXT log and generate the Excel health report",
            font=("Segoe UI", 10),
        )
        subtitle.pack(pady=(0, 18))

        file_frame = tk.Frame(root)
        file_frame.pack(fill="x", padx=25)

        tk.Label(file_frame, text="TXT Log File:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        entry_frame = tk.Frame(file_frame)
        entry_frame.pack(fill="x", pady=(6, 12))

        self.file_entry = tk.Entry(
            entry_frame,
            textvariable=self.selected_file,
            font=("Segoe UI", 10),
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=5)

        browse_btn = tk.Button(
            entry_frame,
            text="Browse...",
            command=self.select_file,
            width=12,
        )
        browse_btn.pack(side="right")

        button_frame = tk.Frame(root)
        button_frame.pack(fill="x", padx=25, pady=(0, 12))

        analyze_btn = tk.Button(
            button_frame,
            text="Analyze & Export Excel",
            command=self.run_analysis,
            font=("Segoe UI", 10, "bold"),
            bg="#1F4E78",
            fg="white",
            padx=16,
            pady=8,
        )
        analyze_btn.pack(side="left")

        clear_btn = tk.Button(
            button_frame,
            text="Clear",
            command=self.clear,
            padx=16,
            pady=8,
        )
        clear_btn.pack(side="left", padx=(10, 0))

        tk.Label(root, text="Result Preview:", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", padx=25, pady=(8, 5)
        )

        self.preview = scrolledtext.ScrolledText(
            root,
            height=15,
            font=("Consolas", 9),
            wrap=tk.WORD,
        )
        self.preview.pack(fill="both", expand=True, padx=25, pady=(0, 20))
        self.preview.insert(tk.END, "Select an extractor TXT log to begin.\n")
        self.preview.config(state="disabled")

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Extractor Log",
            filetypes=[
                ("Text files", "*.txt"),
                ("Log files", "*.log"),
                ("All files", "*.*"),
            ],
        )
        if file_path:
            self.selected_file.set(file_path)

    def set_preview(self, text):
        self.preview.config(state="normal")
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, text)
        self.preview.config(state="disabled")

    def clear(self):
        self.selected_file.set("")
        self.set_preview("Select an extractor TXT log to begin.\n")

    def run_analysis(self):
        file_path = self.selected_file.get().strip()

        if not file_path:
            messagebox.showwarning(APP_TITLE, "Please select a TXT log file first.")
            return

        if not os.path.isfile(file_path):
            messagebox.showerror(APP_TITLE, "The selected file does not exist.")
            return

        try:
            rows = analyze_log(file_path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not read the log file:\n\n{exc}")
            return

        preview_lines = [
            "EXTRACTOR HEALTH MONITOR",
            "=" * 72,
            f"{'Timestamp':22} {'Brand / Store':25} {'Status':12} Reason",
            "-" * 72,
        ]

        for row in rows:
            if row["timestamp"]:
                ts = row["timestamp"].strftime("%m/%d/%Y %I:%M %p")
            else:
                ts = ""

            preview_lines.append(
                f"{ts:22} {row['brand_store'][:24]:25} "
                f"{row['status']:12} {row['reason']}"
            )

        self.set_preview("\n".join(preview_lines))

        default_name = "Extractor_Health_Monitor_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"

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
            messagebox.showerror(APP_TITLE, f"Could not create Excel file:\n\n{exc}")
            return

        messagebox.showinfo(
            APP_TITLE,
            f"Excel report created successfully.\n\n{output_path}"
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = ExtractorHealthApp(root)
    root.mainloop()
