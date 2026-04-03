# =============================================================================
# Program Name : highlight_gui_v2.py
# Author       : Kaori Kashiwagi
# Date         : 2026-04-03
# Purpose      : Main GUI application for the PDF Keyword Highlight Tool.
#                Provides a dark-themed desktop interface where users can:
#                  - Select an Excel configuration file (highlight_master_v3.xlsx)
#                  - Choose between Single folder mode or All customers mode
#                  - Run keyword highlighting on PDF files
#                  - View real-time processing log with color-coded results
#                  - Open the shared Excel log file after processing
# =============================================================================

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys

from excel_master import load_master_rules
from processor import process_customer

# ── Color theme: Charcoal + Mint ──────────────────────────────────────────────
# Dark background colors for the GUI
C_BG        = "#1E1E1E"   # Main window background (dark charcoal)
C_SURFACE   = "#2A2A2A"   # Panel / card background
C_SURFACE2  = "#1a2e28"   # Active mode card background (dark green tint)
C_BORDER    = "#3A3A3A"   # Default border color
C_BORDER2   = "#2a4a3a"   # Active mode card border (green tint)
C_ACCENT    = "#5DCAA5"   # Mint green accent (buttons, highlights)
C_ACCENT_DK = "#1D9E75"   # Darker mint for hover effects
C_TEXT      = "#F0F0F0"   # Primary text color (near white)
C_MUTED     = "#888888"   # Secondary / muted text color

# Log message colors
C_OK        = "#5DCAA5"   # Green for successful processing
C_SKIP      = "#EF9F27"   # Amber for skipped files
C_ERR       = "#F09595"   # Red for errors

# Font definitions
FONT        = ("Segoe UI", 10)
FONT_SM     = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)   # Monospace for file paths and log

# ── Badge colors for each highlight color option ───────────────────────────────
# Format: { color_name: (badge_background, badge_text_color) }
COLOR_BADGES = {
    "blue":      ("#1a3a5c", "#6CB4E8"),
    "lightblue": ("#1a3a5c", "#6CB4E8"),
    "yellow":    ("#3a2e00", "#EF9F27"),
    "green":     ("#0a2e1a", "#5DCAA5"),
    "pink":      ("#3a1a2a", "#ED93B1"),
    "orange":    ("#3a1a00", "#EF9F27"),
    "purple":    ("#1a0a3a", "#AFA9EC"),
    "red":       ("#3a0a0a", "#F09595"),
}


class App(tk.Tk):
    """
    Main application window.
    Inherits from tk.Tk to create the root window directly.
    """

    def __init__(self):
        super().__init__()
        self.title("PDF Keyword Highlight Tool")
        self.geometry("860x800")
        self.configure(bg=C_BG)
        self.resizable(True, True)

        # ── Application state variables ────────────────────────────────────
        self.master_path = tk.StringVar()       # Path to highlight_master_v3.xlsx
        self.input_folder = tk.StringVar()      # Path to Input folder (Single folder mode)
        self.overwrite = tk.BooleanVar(value=False)  # Overwrite existing files in Done folder
        self.mode = "single"                    # Processing mode: "single" or "all"
        self.cancel_flag = {"cancel": False}    # Shared flag to stop processing mid-run
        self.log_path = None                    # Path to the shared Excel log file
        self.running = False                    # True while processing is running
        self._rules = []                        # List of CustomerRule objects loaded from Excel

        self._build_ui()

    # ── UI CONSTRUCTION ───────────────────────────────────────────────────────

    def _build_ui(self):
        """Build all UI elements from top to bottom."""

        # Header bar at the top
        hdr = tk.Frame(self, bg=C_SURFACE, padx=20, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="PDF Keyword Highlight",
                 font=("Segoe UI", 14, "bold"), bg=C_SURFACE, fg=C_ACCENT).pack(side="left")
        tk.Label(hdr, text="  Excel-driven · Color & page range · Auto-log",
                 font=FONT_SM, bg=C_SURFACE, fg=C_MUTED).pack(side="left")

        # Main body area with padding
        body = tk.Frame(self, bg=C_BG, padx=20, pady=16)
        body.pack(fill="both", expand=True)
        self._body = body

        # Step 01 — Excel configuration file selection
        self._section_label(body, "01", "Configuration file (highlight_master.xlsx)")
        row1 = tk.Frame(body, bg=C_BG)
        row1.pack(fill="x", pady=(4, 12))
        self._path_entry(row1, self.master_path, self._browse_excel)

        # Processing mode selector (Single folder / All customers)
        self._section_label(body, "", "Processing mode")
        self._mode_row = tk.Frame(body, bg=C_BG)
        self._mode_row.pack(fill="x", pady=(6, 0))
        self._build_mode_cards()

        # Dynamic area — switches between Step 02 (single) and customer list (all)
        self._dynamic = tk.Frame(body, bg=C_BG)
        self._dynamic.pack(fill="x")

        # Step 02 — Input folder selection (only shown in Single folder mode)
        self._step2_frame = tk.Frame(self._dynamic, bg=C_BG)
        self._section_label(self._step2_frame, "02",
                            "Input folder  —  auto-filled from Excel, or browse")
        row2 = tk.Frame(self._step2_frame, bg=C_BG)
        row2.pack(fill="x", pady=(4, 0))
        self._path_entry(row2, self.input_folder, self._browse_folder)

        # All customers panel (only shown in All customers mode)
        # Displays the list of customers, keywords, colors, and page ranges from Excel
        self._all_frame = tk.Frame(self._dynamic, bg=C_SURFACE2,
                                   highlightbackground=C_BORDER2, highlightthickness=1)
        tk.Label(self._all_frame, text="Customers to process — from Excel",
                 font=("Segoe UI", 9, "bold"), bg=C_SURFACE2, fg=C_ACCENT,
                 anchor="w", padx=12, pady=8).pack(fill="x")
        self._all_list = tk.Frame(self._all_frame, bg=C_SURFACE2)
        self._all_list.pack(fill="x", padx=12, pady=(0, 10))

        # Action buttons — Run, Stop, Open Log
        btn_row = tk.Frame(body, bg=C_BG)
        btn_row.pack(fill="x", pady=(12, 0))

        self.run_btn = tk.Button(btn_row, text="▶  Run — Start Processing",
                                 font=("Segoe UI", 11, "bold"),
                                 bg=C_ACCENT, fg="#1A1A1A",
                                 activebackground=C_ACCENT_DK, activeforeground="#fff",
                                 relief="flat", cursor="hand2", padx=20, pady=10,
                                 command=self._confirm_and_run)
        self.run_btn.pack(side="left", padx=(0, 8))

        self.cancel_btn = tk.Button(btn_row, text="Stop",
                                    font=FONT, bg=C_SURFACE, fg=C_MUTED,
                                    activebackground=C_BORDER, relief="flat",
                                    cursor="hand2", padx=14, pady=10,
                                    command=self._cancel, state="disabled")
        self.cancel_btn.pack(side="left", padx=(0, 8))

        self.open_log_btn = tk.Button(btn_row, text="Open Log",
                                      font=FONT, bg=C_SURFACE, fg=C_MUTED,
                                      activebackground=C_BORDER, relief="flat",
                                      cursor="hand2", padx=14, pady=10,
                                      command=self._open_log, state="disabled")
        self.open_log_btn.pack(side="left")

        # Overwrite checkbox — when checked, existing files in Done folder are deleted and reprocessed
        chk_frame = tk.Frame(body, bg=C_BG)
        chk_frame.pack(fill="x", pady=(8, 0))
        self._chk = tk.Checkbutton(
            chk_frame,
            text="Overwrite existing files in Destination folder",
            variable=self.overwrite,
            font=FONT_SM, bg=C_BG, fg=C_MUTED,
            selectcolor=C_SURFACE,
            activebackground=C_BG, activeforeground=C_TEXT,
            relief="flat", cursor="hand2")
        self._chk.pack(side="left")

        # Progress bar — shows overall processing progress
        prog_bg = tk.Frame(body, bg=C_BORDER, height=2)
        prog_bg.pack(fill="x", pady=(12, 0))
        self.prog_bar = tk.Frame(prog_bg, bg=C_ACCENT, height=2, width=0)
        self.prog_bar.place(x=0, y=0, relheight=1.0)

        # Summary stats cards — Completed / Skipped / Errors
        stats_frame = tk.Frame(body, bg=C_BG)
        stats_frame.pack(fill="x", pady=(12, 0))
        self.stat_ok   = self._stat_card(stats_frame, "0", "Completed", C_OK)
        self.stat_skip = self._stat_card(stats_frame, "0", "Skipped",   C_SKIP)
        self.stat_err  = self._stat_card(stats_frame, "0", "Errors",    C_ERR)

        # Processing log area — scrollable text box showing real-time results
        log_hdr = tk.Frame(body, bg=C_BG)
        log_hdr.pack(fill="x", pady=(12, 4))
        tk.Label(log_hdr, text="PROCESSING LOG", font=("Segoe UI", 9),
                 bg=C_BG, fg=C_MUTED).pack(side="left")
        tk.Button(log_hdr, text="Clear", font=FONT_SM, bg=C_BG, fg=C_MUTED,
                  relief="flat", cursor="hand2",
                  command=self._clear_log).pack(side="right")

        log_outer = tk.Frame(body, bg=C_SURFACE,
                             highlightbackground=C_BORDER, highlightthickness=1)
        log_outer.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_outer, bg=C_SURFACE, fg=C_TEXT,
                                font=FONT_MONO, relief="flat",
                                wrap="word", state="disabled",
                                padx=12, pady=10, cursor="arrow")
        scroll = tk.Scrollbar(log_outer, command=self.log_text.yview, bg=C_SURFACE)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        # Color tags for log messages
        self.log_text.tag_config("ok",   foreground=C_OK)    # Green for success
        self.log_text.tag_config("skip", foreground=C_SKIP)  # Amber for skipped
        self.log_text.tag_config("err",  foreground=C_ERR)   # Red for errors
        self.log_text.tag_config("info", foreground=C_MUTED) # Gray for info
        self.log_text.tag_config("head", foreground=C_ACCENT) # Mint for headers

        self._write("Ready. Select configuration file and mode, then click Run.\n", "info")

        # Watch for Excel file path changes and auto-load rules
        self.master_path.trace_add("write", lambda *_: self._on_excel_selected())

        # Show Single folder mode by default
        self._step2_frame.pack(fill="x", pady=(8, 4))

    # ── MODE CARD CONSTRUCTION ────────────────────────────────────────────────

    def _build_mode_cards(self):
        """Rebuild both mode selection cards (Single folder / All customers)."""
        for w in self._mode_row.winfo_children():
            w.destroy()
        self._mode_card(self._mode_row, "single", "Single folder",
                        "Select one Input folder manually.").pack(
            side="left", expand=True, fill="x", padx=(0, 8))
        self._mode_card(self._mode_row, "all", "All customers",
                        "Process all customers from Excel automatically.").pack(
            side="left", expand=True, fill="x")

    def _mode_card(self, parent, value, title, desc):
        """
        Create a clickable mode selection card.
        Active card is highlighted with mint border and dark green background.
        """
        active = (self.mode == value)
        bg = C_SURFACE2 if active else C_SURFACE
        border = C_ACCENT if active else C_BORDER
        f = tk.Frame(parent, bg=bg, padx=14, pady=10,
                     highlightbackground=border, highlightthickness=2)
        title_row = tk.Frame(f, bg=bg)
        title_row.pack(fill="x")
        # Radio button indicator dot
        dot_bg = C_ACCENT if active else C_BORDER
        dot_f = tk.Frame(title_row, bg=dot_bg, width=14, height=14)
        dot_f.pack(side="left", padx=(0, 6))
        dot_f.pack_propagate(False)
        if active:
            # Inner white dot for selected state
            tk.Frame(dot_f, bg=C_SURFACE2, width=6, height=6).place(
                relx=0.5, rely=0.5, anchor="center")
        tk.Label(title_row, text=title, font=("Segoe UI", 10, "bold"),
                 bg=bg, fg=C_TEXT).pack(side="left")
        tk.Label(f, text=desc, font=FONT_SM, bg=bg, fg=C_MUTED,
                 anchor="w", wraplength=300).pack(fill="x", pady=(4, 0))
        # Bind click event to all child widgets so the whole card is clickable
        for widget in [f, title_row, dot_f] + list(f.winfo_children()):
            widget.bind("<Button-1>", lambda e, v=value: self._set_mode(v))
        return f

    def _set_mode(self, value):
        """Switch between Single folder and All customers mode."""
        self.mode = value
        self._build_mode_cards()
        # Show/hide the appropriate panel
        if self.mode == "single":
            self._all_frame.pack_forget()
            self._step2_frame.pack(fill="x", pady=(8, 4))
            self.run_btn.config(text="▶  Run — Start Processing")
        else:
            self._step2_frame.pack_forget()
            self._all_frame.pack(fill="x", pady=(8, 4))
            self.run_btn.config(text="▶  Run — Process All Customers")
            self._refresh_all_list()

    def _refresh_all_list(self):
        """
        Rebuild the customer list panel shown in All customers mode.
        Displays each customer's name, keywords, color badge, and page range.
        """
        for w in self._all_list.winfo_children():
            w.destroy()
        if not self._rules:
            tk.Label(self._all_list,
                     text="Select highlight_master.xlsx to see customers.",
                     font=FONT_SM, bg=C_SURFACE2, fg=C_MUTED).pack(anchor="w", pady=4)
            return
        for r in self._rules:
            row = tk.Frame(self._all_list, bg=C_SURFACE2)
            row.pack(fill="x", pady=3)
            # Customer name
            tk.Label(row, text=r.customer, font=("Segoe UI", 9, "bold"),
                     bg=C_SURFACE2, fg=C_TEXT, width=14, anchor="w").pack(side="left")
            # Keywords list
            kws = ", ".join(r.keywords) if r.keywords else "(none)"
            tk.Label(row, text=kws, font=FONT_SM, bg=C_SURFACE2,
                     fg=C_MUTED, anchor="w").pack(side="left", expand=True, fill="x")
            # Color badge
            ck = r.color.lower()
            badge_bg, badge_fg = COLOR_BADGES.get(ck, ("#2a2a2a", "#888"))
            tk.Label(row, text=r.color.capitalize(),
                     font=("Segoe UI", 8, "bold"),
                     bg=badge_bg, fg=badge_fg, padx=8, pady=2).pack(side="left", padx=6)
            # Page range
            ep = str(r.end_page) if r.end_page else "end"
            tk.Label(row, text=f"p.{r.begin_page}–{ep}",
                     font=FONT_SM, bg=C_SURFACE2, fg=C_MUTED).pack(side="left")

    # ── HELPER WIDGETS ────────────────────────────────────────────────────────

    def _section_label(self, parent, num, text):
        """Create a section label with optional numbered badge."""
        f = tk.Frame(parent, bg=C_BG)
        f.pack(fill="x", pady=(8, 0))
        if num:
            tk.Label(f, text=num, font=("Segoe UI", 8, "bold"),
                     bg=C_ACCENT, fg="#1A1A1A", width=3, padx=4).pack(side="left")
            tk.Label(f, text="  ", bg=C_BG).pack(side="left")
        tk.Label(f, text=text, font=FONT, bg=C_BG, fg=C_MUTED).pack(side="left")

    def _path_entry(self, parent, var, cmd):
        """Create a path input field with a Browse button."""
        e = tk.Entry(parent, textvariable=var, font=FONT_MONO,
                     bg=C_SURFACE, fg=C_TEXT, insertbackground=C_TEXT,
                     relief="flat", highlightbackground=C_BORDER,
                     highlightthickness=1, highlightcolor=C_ACCENT)
        e.pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 8))
        tk.Button(parent, text="Browse", font=FONT,
                  bg=C_SURFACE, fg=C_ACCENT,
                  activebackground=C_BORDER, relief="flat",
                  cursor="hand2", padx=14, pady=6,
                  command=cmd).pack(side="left")

    def _stat_card(self, parent, num, label, color):
        """Create a summary stat card showing a large number and a label."""
        f = tk.Frame(parent, bg=C_SURFACE, padx=20, pady=10,
                     highlightbackground=C_BORDER, highlightthickness=1)
        f.pack(side="left", expand=True, fill="x", padx=(0, 8))
        lbl = tk.Label(f, text=num, font=("Segoe UI", 22, "bold"),
                       bg=C_SURFACE, fg=color)
        lbl.pack()
        tk.Label(f, text=label, font=FONT_SM, bg=C_SURFACE, fg=C_MUTED).pack()
        return lbl

    # ── FILE / FOLDER BROWSE ──────────────────────────────────────────────────

    def _browse_excel(self):
        """Open file dialog to select highlight_master_v3.xlsx."""
        path = filedialog.askopenfilename(
            title="Select highlight_master.xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.master_path.set(path)

    def _browse_folder(self):
        """Open folder dialog to select the Input folder."""
        path = filedialog.askdirectory(
            title="Select Input folder (containing PDF files)")
        if path:
            self.input_folder.set(path)

    def _on_excel_selected(self):
        """
        Called automatically when the Excel path changes.
        Loads rules from Excel and auto-fills the Input folder field
        with the first customer's CustomerFolder path.
        """
        xlsx = self.master_path.get().strip()
        if not xlsx or not os.path.exists(xlsx):
            self._rules = []
            self._refresh_all_list()
            return
        try:
            self._rules = load_master_rules(xlsx)
            # Auto-fill Input folder from first rule if not already set
            if self._rules and not self.input_folder.get():
                self.input_folder.set(self._rules[0].customer_folder)
            self._refresh_all_list()
        except Exception:
            self._rules = []

    # ── RUN CONFIRMATION ──────────────────────────────────────────────────────

    def _confirm_and_run(self):
        """
        Validate inputs and show a confirmation dialog before starting.
        Routes to _run_single or _run_all depending on the current mode.
        """
        xlsx = self.master_path.get().strip()
        if not xlsx or not os.path.exists(xlsx):
            messagebox.showerror("Error", "Please select a valid highlight_master.xlsx.")
            return

        if self.mode == "single":
            # Validate Input folder and check for PDFs
            folder = self.input_folder.get().strip()
            if not folder or not os.path.isdir(folder):
                messagebox.showerror("Error", "Please select a valid Input folder.")
                return
            pdfs = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
            if not pdfs:
                messagebox.showwarning("No PDFs", "No PDF files found in the selected folder.")
                return
            if messagebox.askyesno("Confirm",
                                   f"Ready to process {len(pdfs)} PDF file(s).\n\n"
                                   f"Input folder:\n{folder}\n\nContinue?"):
                self._run_single(xlsx, folder)
        else:
            # Validate all customer Input folders and count total PDFs
            if not self._rules:
                messagebox.showerror("Error", "No rules found in Excel.")
                return
            total = 0
            missing = []
            for r in self._rules:
                inp = os.path.join(r.customer_folder, "Input")
                if os.path.isdir(inp):
                    total += len([f for f in os.listdir(inp) if f.lower().endswith(".pdf")])
                else:
                    missing.append(r.customer)
            msg = (f"Ready to process all {len(self._rules)} customer(s).\n"
                   f"Total PDF files found: {total}\n")
            if missing:
                msg += "\nWarning — Input folder not found for:\n" + "\n".join(missing)
            msg += "\n\nContinue?"
            if messagebox.askyesno("Confirm — All Customers", msg):
                self._run_all(xlsx)

    # ── SINGLE FOLDER PROCESSING ──────────────────────────────────────────────

    def _run_single(self, xlsx, folder):
        """
        Process PDFs in a single selected Input folder.
        Finds the matching customer rule from Excel based on the folder path.
        Saves the log to C:\\PDFs\\pdf_highlight_log.xlsx (shared log).
        Runs in a background thread to keep the GUI responsive.
        """
        self._prepare_run()

        # Derive shared log path: go up two levels from Input folder
        # e.g. C:\PDFs\CustomerA\Input -> C:\PDFs\pdf_highlight_log.xlsx
        folder_norm = os.path.normpath(folder)
        pdfs_root = os.path.dirname(os.path.dirname(folder_norm))
        shared_log = os.path.join(pdfs_root, "pdf_highlight_log.xlsx")

        def worker():
            try:
                # Always reload Excel fresh to pick up any changes made since last run
                all_rules = load_master_rules(xlsx)

                # Find the rule whose CustomerFolder\Input matches the selected folder
                matched_rule = None
                for r in all_rules:
                    expected = os.path.normpath(os.path.join(r.customer_folder, "Input"))
                    if expected.lower() == folder_norm.lower():
                        matched_rule = r
                        break

                # Fall back to first rule if no exact match found
                if not matched_rule:
                    matched_rule = all_rules[0] if all_rules else None

                if not matched_rule:
                    self._write("No matching rule found in Excel.\n", "err")
                    return

                ow = "ON" if self.overwrite.get() else "OFF"
                self._write(f"Single folder mode — processing: {matched_rule.customer}\n", "head")
                self._write(f"Log: {shared_log}   Overwrite: {ow}\n", "info")
                counts = {"ok": 0, "skip": 0, "err": 0}

                self._write_rule_header(matched_rule)

                # Run the actual PDF processing
                results, lp = process_customer(
                    matched_rule, folder,
                    cancel_flag=self.cancel_flag,
                    progress_callback=lambda n, c, t: self.after(
                        0, lambda w=int(c/t*760): self.prog_bar.config(width=w)),
                    shared_log_path=shared_log,
                    overwrite=self.overwrite.get())
                self.log_path = lp
                self._process_results(results, counts)
                self._finish_run(counts)

            except Exception as e:
                self._write(f"\nFatal error: {e}\n", "err")
            finally:
                self._cleanup_run()

        threading.Thread(target=worker, daemon=True).start()

    # ── ALL CUSTOMERS PROCESSING ──────────────────────────────────────────────

    def _run_all(self, xlsx):
        """
        Process all customers listed in Excel sequentially.
        Each customer's CustomerFolder\\Input\\ is used automatically.
        All results are saved to a single shared log file in C:\\PDFs\\.
        Runs in a background thread to keep the GUI responsive.
        """
        self._prepare_run()

        # Determine shared log location from first rule's parent folder
        # e.g. C:\PDFs\CustomerA -> C:\PDFs\pdf_highlight_log.xlsx
        if self._rules:
            pdfs_root = os.path.dirname(self._rules[0].customer_folder.rstrip("/\\"))
        else:
            pdfs_root = "C:\\PDFs"
        shared_log = os.path.normpath(os.path.join(pdfs_root, "pdf_highlight_log.xlsx"))

        def worker():
            try:
                # Always reload Excel fresh to pick up any changes made since last run
                rules = load_master_rules(xlsx)
                ow = "ON" if self.overwrite.get() else "OFF"
                self._write(f"All customers mode — {len(rules)} customer(s).\n", "head")
                self._write(f"Shared log: {shared_log}   Overwrite: {ow}\n", "info")
                counts = {"ok": 0, "skip": 0, "err": 0}
                total_r = len(rules)

                # Process each customer in order
                for i, rule in enumerate(rules):
                    if self.cancel_flag["cancel"]:
                        break

                    # Input folder is always CustomerFolder\Input\
                    inp = os.path.join(rule.customer_folder, "Input")
                    ep = str(rule.end_page) if rule.end_page else "end"
                    self._write(f"\n── {rule.customer} ──  {inp}\n", "head")
                    self._write(f"   Keywords : {', '.join(rule.keywords)}\n", "info")
                    self._write(f"   Color    : {rule.color.capitalize()}   "
                                f"Pages: {rule.begin_page}–{ep}\n", "info")
                    self._write(f"   Dest     : {rule.destination_folder}\n", "info")

                    # Skip if Input folder does not exist
                    if not os.path.isdir(inp):
                        self._write(f"   Input folder not found — skipping.\n", "err")
                        counts["err"] += 1
                        self.after(0, lambda o=counts["ok"], s=counts["skip"],
                                   e=counts["err"]: self._update_stats(o, s, e))
                        continue

                    # Progress callback calculates overall progress across all customers
                    def progress(name, current, total, idx=i):
                        overall = (idx / total_r) + (current / total / total_r)
                        self.after(0, lambda w=int(overall*760): self.prog_bar.config(width=w))

                    # Run the actual PDF processing for this customer
                    results, lp = process_customer(
                        rule, inp,
                        cancel_flag=self.cancel_flag,
                        progress_callback=progress,
                        shared_log_path=shared_log,
                        overwrite=self.overwrite.get())
                    self.log_path = lp
                    self._process_results(results, counts)

                self._finish_run(counts)

            except Exception as e:
                self._write(f"\nFatal error: {e}\n", "err")
            finally:
                self._cleanup_run()

        threading.Thread(target=worker, daemon=True).start()

    # ── RUN STATE MANAGEMENT ──────────────────────────────────────────────────

    def _prepare_run(self):
        """Disable buttons and reset stats before starting a run."""
        self.running = True
        self.cancel_flag["cancel"] = False
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.open_log_btn.config(state="disabled", fg=C_MUTED)
        self.stat_ok.config(text="0")
        self.stat_skip.config(text="0")
        self.stat_err.config(text="0")
        self.prog_bar.config(width=0)

    def _cleanup_run(self):
        """Re-enable Run button and disable Stop button after run completes."""
        self.running = False
        self.after(0, lambda: self.run_btn.config(state="normal"))
        self.after(0, lambda: self.cancel_btn.config(state="disabled"))

    def _write_rule_header(self, rule):
        """Write a customer rule summary to the processing log."""
        ep = str(rule.end_page) if rule.end_page else "end"
        self._write(f"\n── {rule.customer} ──\n", "head")
        self._write(f"   Keywords : {', '.join(rule.keywords)}\n", "info")
        self._write(f"   Color    : {rule.color.capitalize()}   "
                    f"Pages: {rule.begin_page}–{ep}\n", "info")
        self._write(f"   Dest     : {rule.destination_folder}\n", "info")

    def _process_results(self, results, counts):
        """
        Display each file's processing result in the log and update stat counters.
        OK = green, SKIP = amber, ERROR = red.
        """
        for pdf_name, status in results:
            if status.startswith("OK"):
                tag, counts["ok"] = "ok", counts["ok"] + 1
                label = f"  OK    {pdf_name}  — {status.replace('OK_HITS_','')} hits"
            elif "SKIP" in status:
                tag, counts["skip"] = "skip", counts["skip"] + 1
                label = f"  SKIP  {pdf_name}  — already exists"
            elif "CANCEL" in status:
                tag = "info"
                label = f"  --    {pdf_name}  — cancelled"
            else:
                tag, counts["err"] = "err", counts["err"] + 1
                label = f"  ERR   {pdf_name}  — {status}"
            self._write(label + "\n", tag)
            # Update stat cards on the main thread
            self.after(0, lambda o=counts["ok"], s=counts["skip"],
                       e=counts["err"]: self._update_stats(o, s, e))

    def _finish_run(self, counts):
        """Show final summary in the log and enable Open Log button."""
        if self.cancel_flag["cancel"]:
            self._write("\nStopped by user.\n", "skip")
        else:
            self._write(
                f"\nDone. {counts['ok']} completed  |  "
                f"{counts['skip']} skipped  |  {counts['err']} errors\n"
                f"Log saved to: {self.log_path}\n", "head")
            # Activate Open Log button with accent color
            self.after(0, lambda: self.open_log_btn.config(state="normal", fg=C_ACCENT))
            self.after(0, lambda: self.prog_bar.config(width=760))

    def _update_stats(self, ok, skip, err):
        """Update the three stat cards with current counts."""
        self.stat_ok.config(text=str(ok))
        self.stat_skip.config(text=str(skip))
        self.stat_err.config(text=str(err))

    def _cancel(self):
        """Set cancel flag — processing stops after the current file finishes."""
        self.cancel_flag["cancel"] = True
        self.cancel_btn.config(state="disabled")
        self._write("Stopping after current file...\n", "skip")

    def _open_log(self):
        """Open the shared Excel log file using the default application."""
        if self.log_path and os.path.exists(self.log_path):
            if sys.platform == "win32":
                os.startfile(self.log_path)   # Windows: open with default app
            else:
                subprocess.call(["open", self.log_path])  # macOS fallback
        else:
            messagebox.showinfo("Log", "Log file not found.")

    def _clear_log(self):
        """Clear all text from the processing log area."""
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _write(self, msg, tag=""):
        """
        Thread-safe method to append a message to the processing log.
        Uses self.after() to ensure UI updates happen on the main thread.
        """
        def _do():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg, tag)
            self.log_text.see("end")   # Auto-scroll to latest message
            self.log_text.config(state="disabled")
        self.after(0, _do)


# ── ENTRY POINT ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
