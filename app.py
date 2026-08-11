import json
import os
import queue
import threading
import time
import webbrowser
from datetime import date
from tkinter import filedialog, messagebox

import requests
import customtkinter as ctk

APPDATA = os.environ.get("APPDATA", os.path.expanduser("~"))
CONFIG_DIR = os.path.join(APPDATA, "PaperExtractor")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
MAX_CHARS = 45000
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

MODELS = [
    "nvidia/nemotron-3-ultra-550b-a55b:free",
    "google/gemma-4-31b-it:free",
    "google/gemma-4-26b-a4b-it:free",
    "openai/gpt-oss-20b:free",
]

KEYS = ["file_name", "title", "doi", "year", "abstract", "method",
        "research_gap", "key_findings", "limitations", "status"]
HEADERS = ["File name", "Title", "DOI", "Year", "Abstract", "Method",
           "Research gap", "Key findings", "Limitations", "Status"]

SYSTEM_PROMPT = (
    "You are an academic research assistant. Given the text of a research paper, "
    "extract the following fields and return ONLY valid JSON (no markdown, no code fences) "
    "with exactly these keys: "
    "title, doi, year, abstract, method, research_gap, key_findings, limitations.\n\n"
    "Rules:\n"
    "- If a field is not present in the text, return an empty string for it.\n"
    "- abstract: 1-3 sentence summary of the paper's abstract.\n"
    "- method: how the authors did the study or experiment, 2-5 sentences.\n"
    "- research_gap: the problem or gap the paper addresses.\n"
    "- key_findings: the main results, 3-6 bullet points, each starting with '- '.\n"
    "- limitations: 2-4 bullet points, each starting with '- '.\n"
    "- year: 4-digit publication year if shown (e.g. 2024), otherwise empty string.\n"
    "- Do not invent or guess values. Use only what is in the text."
)


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"api_key": "", "model": MODELS[0]}


def save_config(data):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def extract_text(pdf_path):
    import pdfplumber
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            pages.append(text)
    return "\n".join(pages)


def ask_llm(api_key, model, paper_text, log=None):
    if log is None:
        log = lambda m: None
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": paper_text},
    ]
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": 0.2, "max_tokens": 3500}
    for attempt in range(3):
        try:
            r = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=180)
        except requests.exceptions.RequestException as e:
            if attempt < 2:
                log(f"    (network issue - retrying: {e})")
                time.sleep(3)
                continue
            raise RuntimeError(f"Network error: {e}")
        if r.status_code == 429:
            retry = r.headers.get("Retry-After", "30")
            log(f"    (free tier is busy - waiting {retry}s, try {attempt + 1})")
            time.sleep(float(retry) if retry.isdigit() else 30)
            continue
        if r.status_code == 402:
            raise RuntimeError("Not enough credits for this request.")
        if r.status_code != 200:
            detail = r.text
            try:
                detail = r.json().get("error", {}).get("message", r.text)
            except Exception:
                pass
            raise RuntimeError(f"HTTP {r.status_code}: {str(detail)[:400]}")
        data = r.json()
        try:
            return data["choices"][0]["message"]["content"]
        except Exception:
            raise RuntimeError("Unexpected AI response format.")
    raise RuntimeError("Free tier is busy - tried 3 times. Wait a minute and try again.")


def parse_json(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end < start:
        raise ValueError("No JSON object found")
    return json.loads(text[start:end + 1])


def process_pdf(path, api_key, model, log=None):
    if log is None:
        log = lambda m: None
    base = {k: "" for k in KEYS}
    base["file_name"] = os.path.basename(path)
    try:
        text = extract_text(path)
    except Exception as e:
        log(f"    (could not read PDF: {e})")
        base["status"] = "Failed - could not read PDF"
        return base
    text = " ".join(text.split())
    if len(text.strip()) < 250:
        base["status"] = "Scanned - review manually"
        return base
    if len(text) > MAX_CHARS:
        text = text[:MAX_CHARS]
    try:
        raw = ask_llm(api_key, model, text, log=log)
        data = parse_json(raw)
        if isinstance(data, dict):
            for k in KEYS:
                if k in data and data[k] is not None:
                    base[k] = str(data[k]).strip()
        base["status"] = "Done"
    except Exception as e:
        log(f"    (AI error: {e})")
        base["status"] = "AI error - check manually"
    return base


def export_excel(results, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Papers"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    wrap = Alignment(wrap_text=True, vertical="top")

    for col, name in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    status_colors = {
        "Done": "C6EFCE",
        "Scanned - review manually": "FFEB9C",
        "AI error - check manually": "FFC7CE",
        "Failed - could not read PDF": "FFC7CE",
    }

    for row, item in enumerate(results, start=2):
        for col, key in enumerate(KEYS, start=1):
            cell = ws.cell(row=row, column=col, value=item.get(key, ""))
            cell.alignment = wrap
        color = status_colors.get(item.get("status", ""))
        if color:
            ws.cell(row=row, column=10).fill = PatternFill("solid", fgColor=color)

    widths = [28, 45, 25, 10, 55, 50, 50, 55, 50, 24]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A2"
    wb.save(path)


class App(ctk.CTk):
    def __init__(self):
        self.config_data = load_config()

        try:
            from tkinterdnd2 import TkinterDnD
            ctk.CTk.__init__(self)
            self.TkdndVersion = TkinterDnD._require(self)
            self.dnd_available = True
        except Exception:
            ctk.CTk.__init__(self)
            self.dnd_available = False

        self.title("Paper Extractor - PDF to Excel")
        self.geometry("860x680")
        self.minsize(760, 580)

        self.files = []
        self.last_results = []
        self.processing = False
        self.msg_queue = queue.Queue()

        self.api_key_var = ctk.StringVar(value=self.config_data.get("api_key", ""))
        self.model_var = ctk.StringVar(value=self.config_data.get("model", MODELS[0]))

        self._build_ui()
        if self.dnd_available:
            self._setup_dnd()

        self.after(150, self._poll_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, corner_radius=10)
        header.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 6))
        ctk.CTkLabel(header, text="Paper Extractor", font=ctk.CTkFont(size=24, weight="bold")).pack(side="left", padx=14, pady=8)
        ctk.CTkLabel(header, text="Drag PDFs here -> click Analyze -> Export to Excel",
                     font=ctk.CTkFont(size=13)).pack(side="left", padx=8)

        settings = ctk.CTkFrame(self, corner_radius=10)
        settings.grid(row=1, column=0, sticky="ew", padx=14, pady=6)
        settings.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(settings, text="OpenRouter key:").grid(row=0, column=0, padx=(14, 6), pady=8, sticky="w")
        self.key_entry = ctk.CTkEntry(settings, textvariable=self.api_key_var,
                                      placeholder_text="Paste your OpenRouter key (sk-or-...)", show="*")
        self.key_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=6, pady=8)

        ctk.CTkButton(settings, text="Get free key",
                      command=lambda: webbrowser.open("https://openrouter.ai/settings/keys"),
                      width=110, fg_color="#2e7d32", hover_color="#1b5e20").grid(row=0, column=3, padx=(6, 14), pady=8)

        ctk.CTkLabel(settings, text="AI model:").grid(row=1, column=0, padx=(14, 6), pady=8, sticky="w")
        ctk.CTkOptionMenu(settings, values=MODELS, variable=self.model_var, width=280).grid(row=1, column=1, padx=6, pady=8, sticky="w")
        ctk.CTkLabel(settings, text="Your key is stored only on this PC in a private folder.",
                     font=ctk.CTkFont(size=11), text_color="#e0a458").grid(row=1, column=3, padx=(6, 14), pady=8, sticky="e")

        self.file_list = ctk.CTkTextbox(self, corner_radius=10, wrap="none")
        self.file_list.grid(row=2, column=0, sticky="nsew", padx=14, pady=6)
        self.file_list.configure(state="disabled")
        self._set_list_placeholder()

        buttons = ctk.CTkFrame(self, corner_radius=10)
        buttons.grid(row=3, column=0, sticky="ew", padx=14, pady=(2, 6))

        ctk.CTkButton(buttons, text="+ Add PDFs", command=self.add_files, width=120).pack(side="left", padx=(14, 6), pady=10)
        ctk.CTkButton(buttons, text="Clear list", command=self.clear_list, width=100,
                      fg_color="#6d6d6d", hover_color="#4a4a4a").pack(side="left", padx=6, pady=10)
        ctk.CTkButton(buttons, text="Analyze  (AI)", command=self.process_all, width=140,
                      fg_color="#2e7d32", hover_color="#1b5e20").pack(side="left", padx=6, pady=10)
        ctk.CTkButton(buttons, text="Export to Excel", command=self.export_results, width=140,
                      fg_color="#1565c0", hover_color="#0d47a1").pack(side="left", padx=6, pady=10)

        self.progress = ctk.CTkProgressBar(buttons, width=170)
        self.progress.set(0)
        self.progress.pack(side="right", padx=14, pady=10)

        self.status_label = ctk.CTkLabel(self, text="Ready. Add PDFs to begin.", anchor="w")
        self.status_label.grid(row=4, column=0, sticky="ew", padx=16, pady=(0, 12))

    def _setup_dnd(self):
        for widget in [self.file_list, self.key_entry]:
            try:
                widget.drop_target_register("DND_Files")
                widget.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _on_drop(self, event):
        if self.processing:
            return
        try:
            files = self.tk.splitlist(event.data)
        except Exception:
            files = str(event.data).split()
        added = 0
        for f in files:
            f = f.strip().strip("{}")
            if f.lower().endswith(".pdf") and f not in self.files:
                self.files.append(f)
                added += 1
        if added:
            self._refresh_list()
            self.status_label.configure(text=f"Added {added} PDF(s). Click 'Analyze' to start.")

    def _set_list_placeholder(self):
        self.file_list.configure(state="normal")
        self.file_list.delete("1.0", "end")
        msg = "Drop your PDF files here, or click 'Add PDFs'.\n\nResults will appear here step by step, then you export them to Excel."
        self.file_list.insert("1.0", msg)
        self.file_list.configure(state="disabled")

    def _refresh_list(self):
        self.file_list.configure(state="normal")
        self.file_list.delete("1.0", "end")
        for i, f in enumerate(self.files, start=1):
            self.file_list.insert("end", f"{i}. {os.path.basename(f)}\n")
        self.file_list.configure(state="disabled")

    def add_files(self):
        if self.processing:
            return
        paths = filedialog.askopenfilenames(title="Choose PDF files", filetypes=[("PDF files", "*.pdf")])
        added = 0
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                added += 1
        if added:
            self._refresh_list()
            self.status_label.configure(text=f"Added {added} PDF(s). Click 'Analyze' to start.")

    def clear_list(self):
        if self.processing:
            return
        self.files = []
        self.last_results = []
        self._set_list_placeholder()
        self.status_label.configure(text="List cleared.")

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                self._handle_msg(kind, payload)
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def _handle_msg(self, kind, payload):
        if kind == "status":
            self.status_label.configure(text=payload)
        elif kind == "log":
            self._log(payload)
        elif kind == "progress":
            self.progress.set(payload)
        elif kind == "done":
            self.processing = False
            self.progress.set(1)
            self.status_label.configure(text=f"Finished. {payload} file(s) processed. Click 'Export to Excel'.")
            messagebox.showinfo("Done", f"Processing finished for {payload} file(s).\nClick 'Export to Excel' to save the results.")
        elif kind == "error":
            self.processing = False
            messagebox.showerror("Something went wrong", payload)
            self.status_label.configure(text="Error - see popup.")

    def _log(self, text):
        self.file_list.configure(state="normal")
        self.file_list.insert("end", text + "\n")
        self.file_list.see("end")
        self.file_list.configure(state="disabled")

    def msg(self, kind, payload):
        self.msg_queue.put((kind, payload))

    def process_all(self):
        if self.processing:
            return
        if not self.files:
            messagebox.showwarning("No files", "Add some PDF files first.")
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("API key needed",
                                   "Paste your OpenRouter key above first.\nOpening the page where you can get it...")
            webbrowser.open("https://openrouter.ai/settings/keys")
            self.key_entry.focus_set()
            return
        model = self.model_var.get()
        self.config_data = {"api_key": api_key, "model": model}
        save_config(self.config_data)

        files = list(self.files)
        self.last_results = []
        self.processing = True
        self.progress.set(0)
        self.file_list.configure(state="normal")
        self.file_list.delete("1.0", "end")
        self.file_list.insert("1.0", "Analyzing... results will appear below as each paper is done.\n\n")
        self.file_list.configure(state="disabled")

        def worker():
            total = len(files)
            results = []
            try:
                for i, f in enumerate(files):
                    self.msg("log", f"[{i + 1}/{total}] {os.path.basename(f)} ...")
                    res = process_pdf(f, api_key, model, log=lambda m: self.msg("log", m))
                    results.append(res)
                    self.msg("log", f"    -> {res['status']}")
                    self.msg("progress", (i + 1) / total)
                    if i < total - 1 and res["status"] == "Done":
                        time.sleep(2)
                self.last_results = results
                self.msg("done", len(results))
            except Exception as e:
                self.msg("error", f"Unexpected error:\n{e}")

        threading.Thread(target=worker, daemon=True).start()

    def export_results(self):
        if not self.last_results:
            messagebox.showwarning("Nothing to export", "Analyze your PDFs first, then export.")
            return
        default_name = f"Paper_Extractor_{date.today().isoformat()}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save Excel file", defaultextension=".xlsx",
            initialfile=default_name, filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            export_excel(self.last_results, path)
        except Exception as e:
            messagebox.showerror("Export failed", f"Could not save Excel file:\n{e}")
            return
        self.status_label.configure(text=f"Saved: {path}")
        if messagebox.askyesno("Saved", f"Excel file saved:\n{path}\n\nOpen it now?"):
            os.startfile(path)

    def _on_close(self):
        self.config_data = {"api_key": self.api_key_var.get().strip(), "model": self.model_var.get()}
        save_config(self.config_data)
        self.destroy()


if __name__ == "__main__":
    App().mainloop()