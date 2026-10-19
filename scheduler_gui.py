#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scheduler_gui.py
=================
Tkinter desktop GUI. All optimization logic lives in
``scheduler_engine``; this module is presentation-only so it can change
(or be swapped for another toolkit) without touching the math.

Layout follows one rule: the workbook holds the conference data, this
window holds the decisions made per run -- what to prioritise (weights),
how hard to search (gap, time limit) and how many alternatives to return.

Notes on a few deliberate choices:
  - The input file is validated the moment it is chosen, so a format
    problem is reported before the user presses Run, not after a failed
    solve.
  - The result panel shows the score *broken into its four parts*. A bare
    "-35.0" means nothing on its own; the breakdown is what lets a user
    see which weight did the work.
  - The solution quality line distinguishes a proven optimum from a
    schedule the time limit cut short. CBC calls both "Optimal".
  - Errors are written to the log in full. They used to be truncated to
    90 characters in the banner and nowhere else, which hid exactly the
    part naming the row to fix.
  - Language: the "日本語 | EN" switch in the header rebuilds the window
    in the other language on the spot. What the window shows (file
    summary, last result, log) is kept as data and redrawn, so nothing
    is lost. The switch is locked while a run is in progress so the log
    of one run is never in two languages.
"""

from __future__ import annotations

import os
import platform
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

import scheduler_engine as engine
import scheduler_messages as msg
from scheduler_messages import t


class QueueStream:
    """Thread-safe stream that redirects writes to a queue, optionally
    mirroring to the original stream and/or a log file."""

    def __init__(self, log_queue: "queue.Queue", original_stream=None, log_file=None):
        self.log_queue = log_queue
        self.original_stream = original_stream
        self.log_file = log_file

    def write(self, text):
        if not text:
            return
        self.log_queue.put(text)
        for sink in (self.original_stream, self.log_file):
            if sink:
                try:
                    sink.write(text)
                    sink.flush()
                except Exception:
                    pass

    def flush(self):
        for sink in (self.original_stream, self.log_file):
            if sink:
                try:
                    sink.flush()
                except Exception:
                    pass


class ConferenceSchedulerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Conference Presentation Scheduler")
        self.root.geometry("960x820")
        self.root.minsize(840, 640)

        self.log_queue: "queue.Queue" = queue.Queue()
        self.is_running = False
        self.run_start_time = None
        self.last_output_path = None
        self.last_timetables: dict = {}

        # Both start empty. Guessing from the working directory used to
        # depend on how the program was started: from the .app it is "/",
        # which put the output at "/output.xlsx", where nothing can be
        # written. Choosing the input fills in the output beside it.
        self.input_file_var = tk.StringVar(value="")
        self.output_file_var = tk.StringVar(value="")
        self.input_summary_var = tk.StringVar(value="")
        self.elapsed_var = tk.StringVar(value="")

        self.w_continuity_var = tk.DoubleVar(value=1.0)
        self.w_conflict_var = tk.DoubleVar(value=1.0)
        self.w_empty_var = tk.DoubleVar(value=1.0)
        self.w_continuity_text = tk.StringVar(value="1.0")
        self.w_conflict_text = tk.StringVar(value="1.0")
        self.w_empty_text = tk.StringVar(value="1.0")

        self.mip_gap_var = tk.StringVar(value="")     # blank => exact solve
        self.time_limit_var = tk.StringVar(value="")  # blank => no limit
        self.solutions_var = tk.StringVar(value="1")
        self.save_log_var = tk.BooleanVar(value=False)

        self.result_headline_var = tk.StringVar(value="")
        self.result_quality_var = tk.StringVar(value="")
        self.result_breakdown_var = tk.StringVar(value="")

        # What the window is showing, kept as data so it can be redrawn
        # in the other language.
        self._summary_state = None   # ("ok", info) / ("invalid", text) / ("unreadable", text)
        self._last_result = None
        self._last_error = None
        self._last_elapsed = None

        self._setup_styles()
        self._build_ui()
        self._append_log(f"Host OS: {platform.system()} ({platform.machine()})\n")
        self._append_log(f"Base Directory: {engine.get_base_dir()}\n")
        self._append_log("-" * 60 + "\n")
        self._refresh_input_summary()

        self.root.after(50, self._poll_log_queue)

    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------
    def _setup_styles(self):
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("Header.TFrame", background="#1e293b")
        style.configure("HeaderTitle.TLabel", background="#1e293b", foreground="#ffffff",
                        font=("Helvetica", 14, "bold"))
        # White cards on the grey window, with the input rows in grey. The
        # frame and its heading stay grey and a white body fills the border
        # (see _card): a white frame would also paint the strip beside the
        # heading white. Text labels inside a body share its white, otherwise
        # the empty file summary shows as a small grey bar.
        card_white = "#ffffff"
        style.configure("Card.TLabelframe", padding=0)
        style.configure("Card.TLabelframe.Label", font=("Helvetica", 10, "bold"),
                        foreground="#1e293b")
        style.configure("CardBody.TFrame", background=card_white)
        style.configure("Accent.TButton", font=("Helvetica", 11, "bold"), padding=(15, 6))
        style.configure("Action.TButton", font=("Helvetica", 9), padding=(8, 4))
        style.configure("Hint.TLabel", font=("Helvetica", 8), foreground="#64748b",
                        background=card_white)
        style.configure("Summary.TLabel", font=("Helvetica", 9), foreground="#334155",
                        background=card_white)
        style.configure("Elapsed.TLabel", font=("Helvetica", 9, "italic"), foreground="#475569")
        style.configure("Success.TLabel", font=("Helvetica", 11, "bold"), foreground="#15803d")
        style.configure("Warn.TLabel", font=("Helvetica", 11, "bold"), foreground="#b45309")
        style.configure("Error.TLabel", font=("Helvetica", 11, "bold"), foreground="#b91c1c")
        style.configure("Detail.TLabel", font=("Helvetica", 9), foreground="#475569")
        # Language switch: the active language is drawn light on dark.
        style.configure("LangOn.TButton", font=("Helvetica", 9, "bold"), padding=(8, 2),
                        background="#f8fafc", foreground="#1e293b")
        style.configure("LangOff.TButton", font=("Helvetica", 9), padding=(8, 2),
                        background="#334155", foreground="#cbd5e1")
        style.map("LangOff.TButton", background=[("active", "#475569")])

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    @staticmethod
    def _card(parent, title: str, **pack_options):
        """A titled card; returns the white body its contents go in."""
        frame = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe")
        frame.pack(**pack_options)
        body = ttk.Frame(frame, style="CardBody.TFrame", padding=10)
        body.pack(fill=tk.BOTH, expand=True)
        return body

    def _build_ui(self):
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(15, 10))
        header.pack(fill=tk.X, side=tk.TOP)
        ttk.Label(header, text="Conference Presentation Scheduler",
                  style="HeaderTitle.TLabel").pack(side=tk.LEFT)
        switch = ttk.Frame(header, style="Header.TFrame")
        switch.pack(side=tk.RIGHT)
        self.lang_buttons = {}
        for code, label in (("ja", "日本語"), ("en", "EN")):
            button = ttk.Button(
                switch, text=label,
                style="LangOn.TButton" if msg.get_language() == code else "LangOff.TButton",
                command=lambda c=code: self._switch_language(c),
            )
            button.pack(side=tk.LEFT)
            self.lang_buttons[code] = button
        self.header = header

        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)
        self.main = main

        # --- Files -----------------------------------------------------
        files = self._card(main, t("gui.files"), fill=tk.X, pady=(0, 8))

        row = ttk.Frame(files)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text=t("gui.input"), width=6).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.input_file_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(row, text=t("gui.browse"), style="Action.TButton",
                   command=self._browse_input_file).pack(side=tk.RIGHT)

        row = ttk.Frame(files)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text=t("gui.output"), width=6).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.output_file_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(row, text=t("gui.change"), style="Action.TButton",
                   command=self._browse_output_file).pack(side=tk.RIGHT)

        ttk.Label(files, textvariable=self.input_summary_var, style="Summary.TLabel",
                  wraplength=880, justify=tk.LEFT).pack(anchor=tk.W, pady=(4, 0))

        # --- Weights ---------------------------------------------------
        # Sliders rather than text boxes. The three terms are normalised
        # inside the engine, so a weight is a relative priority on a known
        # 0-10 range -- not a number whose scale the user has to guess.
        weights = self._card(main, t("gui.weights"), fill=tk.X, pady=(0, 8))
        for label, variable, display in (
            (t("gui.w_continuity"), self.w_continuity_var, self.w_continuity_text),
            (t("gui.w_conflict"), self.w_conflict_var, self.w_conflict_text),
            (t("gui.w_empty"), self.w_empty_var, self.w_empty_text),
        ):
            row = ttk.Frame(weights)
            row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=label, width=16).pack(side=tk.LEFT)
            ttk.Scale(
                row, from_=0.0, to=10.0, orient=tk.HORIZONTAL, length=280, variable=variable,
                command=lambda raw, d=display: d.set(f"{round(float(raw), 1):.1f}"),
            ).pack(side=tk.LEFT, padx=(4, 8))
            ttk.Label(row, textvariable=display, width=5).pack(side=tk.LEFT)
        ttk.Label(weights, text=t("gui.weights_hint"),
                  style="Hint.TLabel").pack(anchor=tk.W, pady=(4, 0))

        # --- Run settings ----------------------------------------------
        options = self._card(main, t("gui.run_settings"), fill=tk.X, pady=(0, 8))

        row = ttk.Frame(options)
        row.pack(fill=tk.X)
        ttk.Label(row, text=t("gui.mip_gap")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(row, textvariable=self.mip_gap_var, width=6).pack(side=tk.LEFT)
        ttk.Label(row, text="%").pack(side=tk.LEFT, padx=(2, 18))
        ttk.Label(row, text=t("gui.time_limit")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(row, textvariable=self.time_limit_var, width=6).pack(side=tk.LEFT)
        ttk.Label(row, text=t("gui.seconds_unit")).pack(side=tk.LEFT, padx=(2, 18))
        ttk.Label(row, text=t("gui.solutions")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Combobox(row, textvariable=self.solutions_var,
                     values=["1", str(engine.MAX_SOLUTIONS)],
                     width=3, state="readonly").pack(side=tk.LEFT)

        ttk.Label(options, text=t("gui.settings_hint", max=engine.MAX_SOLUTIONS),
                  style="Hint.TLabel", wraplength=880, justify=tk.LEFT).pack(anchor=tk.W, pady=(4, 0))
        ttk.Checkbutton(options, text=t("gui.save_log"),
                        variable=self.save_log_var).pack(anchor=tk.W, pady=(4, 0))

        # --- Run bar ----------------------------------------------------
        control = ttk.Frame(main)
        control.pack(fill=tk.X, pady=(0, 8))
        self.run_btn = ttk.Button(control, text=t("gui.run"), style="Accent.TButton",
                                  command=self._start_optimization_thread)
        self.run_btn.pack(side=tk.LEFT)
        self.progress_bar = ttk.Progressbar(control, mode="indeterminate", length=160)
        self.progress_bar.pack(side=tk.RIGHT)
        ttk.Label(control, textvariable=self.elapsed_var,
                  style="Elapsed.TLabel").pack(side=tk.RIGHT, padx=8)

        # --- Result ------------------------------------------------------
        result = ttk.Frame(main, padding=(4, 0))
        result.pack(fill=tk.X, pady=(0, 6))
        self.result_headline = ttk.Label(result, textvariable=self.result_headline_var,
                                         style="Success.TLabel", wraplength=900, justify=tk.LEFT)
        self.result_headline.pack(anchor=tk.W)
        ttk.Label(result, textvariable=self.result_quality_var, style="Detail.TLabel",
                  wraplength=900, justify=tk.LEFT).pack(anchor=tk.W)
        ttk.Label(result, textvariable=self.result_breakdown_var, style="Detail.TLabel",
                  wraplength=900, justify=tk.LEFT).pack(anchor=tk.W)

        # --- Footer ---------------------------------------------------
        # Packed before the log: pack() hands out space in order, so a
        # footer packed last got none whenever the log's requested height
        # filled the window, and its buttons were hidden until the window
        # was enlarged. Packed first, it always keeps its row and the log
        # takes whatever height is left.
        footer = ttk.Frame(main)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        self.preview_btn = ttk.Button(footer, text=t("gui.preview"), style="Action.TButton",
                                      command=self._show_timetable_preview, state=tk.DISABLED)
        self.preview_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.open_file_btn = ttk.Button(footer, text=t("gui.open_excel"), style="Action.TButton",
                                        command=self._open_output_file, state=tk.DISABLED)
        self.open_file_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.open_folder_btn = ttk.Button(footer, text=t("gui.open_folder"), style="Action.TButton",
                                          command=self._open_output_folder, state=tk.DISABLED)
        self.open_folder_btn.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(footer, text=t("gui.quit"), style="Action.TButton",
                   command=self.root.quit).pack(side=tk.RIGHT)

        # --- Log ---------------------------------------------------------
        log_card = self._card(main, t("gui.log"), fill=tk.BOTH, expand=True, pady=(0, 8))
        log_font = ("Menlo", 10) if platform.system() == "Darwin" else ("Consolas", 10)
        self.log_console = scrolledtext.ScrolledText(
            log_card, wrap=tk.WORD, bg="#0f172a", fg="#f8fafc", insertbackground="#ffffff",
            font=log_font, relief=tk.FLAT, padx=10, pady=10,
        )
        self.log_console.pack(fill=tk.BOTH, expand=True)

        # Result buttons follow what is available, including after a
        # language switch rebuilt them.
        if self.last_timetables:
            self.preview_btn.config(state=tk.NORMAL)
        if self.last_output_path:
            self.open_file_btn.config(state=tk.NORMAL)
            self.open_folder_btn.config(state=tk.NORMAL)

    # ------------------------------------------------------------------
    # Language
    # ------------------------------------------------------------------
    def _switch_language(self, language: str):
        if self.is_running or language == msg.get_language():
            return
        msg.set_language(language)
        log_text = self.log_console.get("1.0", "end-1c")
        self.header.destroy()
        self.main.destroy()
        self._build_ui()
        self.log_console.insert(tk.END, log_text)
        self.log_console.see(tk.END)
        # Re-check the file quietly so its messages follow the new language.
        self._refresh_input_summary(log=False)
        self._render_result()
        if self._last_elapsed is not None:
            self.elapsed_var.set(t("gui.elapsed", n=f"{self._last_elapsed:.0f}"))

    def _set_language_switch_enabled(self, enabled: bool):
        for button in self.lang_buttons.values():
            button.config(state=tk.NORMAL if enabled else tk.DISABLED)

    # ------------------------------------------------------------------
    # Input summary / validation on selection
    # ------------------------------------------------------------------
    def _refresh_input_summary(self, log: bool = True):
        path = self.input_file_var.get().strip()
        if not path or not os.path.exists(path):
            self._summary_state = None
        else:
            try:
                self._summary_state = ("ok", engine.summarize_input(path))
            except engine.InputValidationError as exc:
                self._summary_state = ("invalid", str(exc))
                if log:
                    self._append_log(t("gui.summary_log_header") + str(exc) + "\n")
            except Exception as exc:
                self._summary_state = ("unreadable", str(exc))
        self._render_input_summary()

    def _render_input_summary(self):
        if self._summary_state is None:
            self.input_summary_var.set("")
            return
        kind, payload = self._summary_state
        if kind == "invalid":
            first = payload.strip().splitlines()[0].rstrip(":：")
            self.input_summary_var.set(t("gui.summary_invalid", first=first))
            return
        if kind == "unreadable":
            self.input_summary_var.set(t("gui.summary_unreadable", error=payload))
            return

        info = payload
        parts = [
            t("gui.n_presentations", n=info["presentations"]),
            t("gui.n_rooms", n=info["rooms"]),
            t("gui.n_timeslots", n=info["timeslots"]),
            t("gui.n_topics", n=info["topics"]),
        ]
        if info["speakers"]:
            parts.append(t("gui.n_speakers", n=info["speakers"]))
        if info["conflict_pairs"]:
            parts.append(t("gui.n_conflicts", n=info["conflict_pairs"]))
        summary = " / ".join(parts) + t("gui.slots_available", n=info["slots_available"])
        self.input_summary_var.set(summary)

    # ------------------------------------------------------------------
    # File dialogs
    # ------------------------------------------------------------------
    def _browse_input_file(self):
        selected = filedialog.askopenfilename(
            title=t("gui.choose_input"),
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
            initialdir=os.path.dirname(self.input_file_var.get()) or os.path.expanduser("~"),
            parent=self.root,
        )
        if selected:
            self.input_file_var.set(selected)
            self.output_file_var.set(os.path.join(os.path.dirname(selected), "output.xlsx"))
            # The result on screen belongs to the previous file; leaving it
            # there next to a new file name invites reading it as this one's.
            # (The saved output workbook is untouched.)
            if not self.is_running:
                self._clear_result()
            self._refresh_input_summary()

    def _browse_output_file(self):
        selected = filedialog.asksaveasfilename(
            title=t("gui.choose_output"),
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialdir=os.path.dirname(self.output_file_var.get()) or os.path.expanduser("~"),
            parent=self.root,
        )
        if selected:
            self.output_file_var.set(selected)

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------
    def _append_log(self, text):
        self.log_console.insert(tk.END, text)
        self.log_console.see(tk.END)

    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple):
                    event_type, payload = item
                    if event_type == "DONE":
                        self._on_optimization_success(payload)
                    elif event_type == "ERROR":
                        self._on_optimization_error(payload)
                else:
                    self._append_log(str(item))
        except queue.Empty:
            pass
        finally:
            if self.is_running and self.run_start_time is not None:
                self.elapsed_var.set(t("gui.elapsed", n=f"{time.time() - self.run_start_time:.0f}"))
            self.root.after(50, self._poll_log_queue)

    # ------------------------------------------------------------------
    # Field parsing
    # ------------------------------------------------------------------
    _INVALID = object()

    def _parse_optional_number(self, variable, label: str, blank_note: str, cast):
        """Read an optional numeric field. Blank returns None.

        A bad value warns and returns the _INVALID sentinel; it is never
        silently treated as "unset", which used to hide typing mistakes.
        """
        raw = variable.get().strip()
        if not raw:
            return None
        try:
            value = float(raw)
        except ValueError:
            messagebox.showwarning(t("gui.input_error_title"),
                                   t("gui.not_number", label=label, note=blank_note),
                                   parent=self.root)
            return self._INVALID
        if value < 0:
            messagebox.showwarning(t("gui.input_error_title"),
                                   t("gui.negative", label=label, note=blank_note),
                                   parent=self.root)
            return self._INVALID
        return cast(value)

    # ------------------------------------------------------------------
    # Run
    # ------------------------------------------------------------------
    def _start_optimization_thread(self):
        if self.is_running:
            return

        input_path = self.input_file_var.get().strip()
        output_path = self.output_file_var.get().strip()

        if not input_path:
            messagebox.showwarning(t("gui.input_error_title"), t("gui.input_empty"),
                                   parent=self.root)
            return
        if not os.path.exists(input_path):
            messagebox.showerror(t("gui.file_missing_title"),
                                 t("gui.input_missing", path=input_path), parent=self.root)
            return
        if not output_path:
            messagebox.showwarning(t("gui.output_error_title"), t("gui.output_missing"),
                                   parent=self.root)
            return

        # Sliders are bounded and snap to 0.1, so no validation is needed.
        weights = [round(v.get(), 1) for v in
                   (self.w_continuity_var, self.w_conflict_var, self.w_empty_var)]

        gap = self._parse_optional_number(
            self.mip_gap_var, t("gui.mip_gap"), t("gui.gap_note"), lambda v: v / 100.0)
        if gap is self._INVALID:
            return
        time_limit = self._parse_optional_number(
            self.time_limit_var, t("gui.time_limit"), t("gui.limit_note"), int)
        if time_limit is self._INVALID:
            return

        try:
            solutions = int(self.solutions_var.get())
        except ValueError:
            solutions = 1

        # Running twice in a row used to overwrite the previous result
        # without a word, which matters for a tool meant to be run
        # repeatedly while trying different weights.
        if os.path.exists(output_path):
            keep = messagebox.askyesno(
                t("gui.overwrite_title"),
                t("gui.overwrite_body", name=os.path.basename(output_path)),
                parent=self.root,
            )
            if not keep:
                base, ext = os.path.splitext(output_path)
                output_path = f"{base}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
                self.output_file_var.set(output_path)

        log_file_handle = None
        if self.save_log_var.get():
            log_dir = os.path.join(os.path.dirname(os.path.abspath(output_path)) or ".", "logs")
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(log_dir, f"run_{time.strftime('%Y%m%d_%H%M%S')}.log")
            try:
                log_file_handle = open(log_path, "w", encoding="utf-8")
                self._append_log(t("gui.log_saving", path=log_path))
            except Exception as exc:
                self._append_log(t("gui.log_open_failed", error=exc))

        self.is_running = True
        self.run_start_time = time.time()
        self.run_btn.config(state=tk.DISABLED)
        self._set_language_switch_enabled(False)
        for button in (self.preview_btn, self.open_file_btn, self.open_folder_btn):
            button.config(state=tk.DISABLED)
        self._last_result = None
        self._last_error = None
        self.last_timetables = {}
        self.last_output_path = None
        self.elapsed_var.set(t("gui.elapsed", n=0))
        self.result_headline.config(style="Success.TLabel")
        self.result_headline_var.set(t("gui.running"))
        self.result_quality_var.set("")
        self.result_breakdown_var.set("")
        self.progress_bar.start(10)

        self._append_log("\n" + "=" * 60 + "\n")
        self._append_log(t("gui.run_started", time=time.strftime("%Y-%m-%d %H:%M:%S")))

        threading.Thread(
            target=self._run_worker,
            args=(input_path, output_path, weights, gap, time_limit, solutions, log_file_handle),
            daemon=True,
        ).start()

    def _run_worker(self, input_path, output_path, weights, gap, time_limit,
                    solutions, log_file_handle):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = QueueStream(self.log_queue, old_stdout, log_file_handle)
        sys.stderr = QueueStream(self.log_queue, old_stderr, log_file_handle)
        try:
            result = engine.run_optimization(
                input_path,
                output_path,
                weight_continuity=weights[0],
                weight_conflict=weights[1],
                weight_empty_seats=weights[2],
                mip_gap=gap if gap is not None else engine.DEFAULT_MIP_GAP,
                time_limit=time_limit,
                solutions=solutions,
            )
            self.log_queue.put(("DONE", result))
        except Exception as exc:
            self.log_queue.put(("ERROR", str(exc)))
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr
            if log_file_handle:
                try:
                    log_file_handle.close()
                except Exception:
                    pass

    # ------------------------------------------------------------------
    # Result handling
    # ------------------------------------------------------------------
    def _clear_result(self):
        """Forget the last result or error and grey out the buttons that use it."""
        self._last_result = None
        self._last_error = None
        self._last_elapsed = None
        self.last_timetables = {}
        self.last_output_path = None
        self.elapsed_var.set("")
        self._render_result()
        for button in (self.preview_btn, self.open_file_btn, self.open_folder_btn):
            button.config(state=tk.DISABLED)

    def _finish_run(self):
        self.is_running = False
        if self.run_start_time is not None:
            self._last_elapsed = time.time() - self.run_start_time
        self.progress_bar.stop()
        self.run_btn.config(state=tk.NORMAL)
        self._set_language_switch_enabled(True)

    def _on_optimization_success(self, result):
        self._finish_run()
        self.last_output_path = result["output_file"]
        self.open_file_btn.config(state=tk.NORMAL)
        self.open_folder_btn.config(state=tk.NORMAL)

        self.last_timetables = {}
        try:
            all_sheets = pd.read_excel(self.last_output_path, sheet_name=None)
            self.last_timetables = {
                name: frame for name, frame in all_sheets.items()
                if name.startswith("Result_Timetable")
            }
            if self.last_timetables:
                self.preview_btn.config(state=tk.NORMAL)
        except Exception as exc:
            self._append_log(t("gui.preview_load_failed", error=exc))

        self._last_result = result
        self._last_error = None
        self._render_result()

    def _on_optimization_error(self, error_message):
        self._finish_run()
        # The engine's messages name the offending row, presentation or
        # speaker, and that detail lives past the 90th character.
        self._append_log(t("gui.aborted") + str(error_message) + "\n")
        self._last_result = None
        self._last_error = str(error_message)
        self._render_result()

    def _render_result(self):
        """Draw the result panel from the last result or error."""
        if self._last_error is not None:
            text = self._last_error.strip()
            first_line = text.splitlines()[0] if text else t("gui.unknown_error")
            self.result_headline.config(style="Error.TLabel")
            self.result_headline_var.set(first_line)
            self.result_quality_var.set(t("gui.see_log"))
            self.result_breakdown_var.set("")
            self.root.update_idletasks()
            return

        result = self._last_result
        if result is None:
            self.result_headline_var.set("")
            self.result_quality_var.set("")
            self.result_breakdown_var.set("")
            return

        quality = result.get("solution_quality", "")
        truncated = "time limit" in quality.lower()
        self.result_headline.config(style="Warn.TLabel" if truncated else "Success.TLabel")

        returned = result.get("solutions_returned", 1)
        headline = t("gui.headline", score=result["objective_score"],
                     time=result["calc_time"], count=result["assigned_count"])
        if returned > 1:
            headline += t("gui.headline_candidates", n=returned)
        self.result_headline_var.set(headline)

        if truncated:
            quality_text = t("gui.quality_truncated")
        elif quality.startswith("Optimal within"):
            percent = quality.replace("Optimal within ", "").replace(" gap", "")
            quality_text = t("gui.quality_gap", percent=percent)
        else:
            quality_text = t("gui.quality_exact")
        self.result_quality_var.set(quality_text)

        candidates = result.get("candidates") or []
        if candidates:
            best = candidates[0]
            self.result_breakdown_var.set(t(
                "gui.breakdown",
                a_count=best["continuity_bonuses"], a=best["score_continuity"],
                b1_count=best["cross_topic_conflicts"], b1=best["score_conflict"],
                b2_count=best["same_topic_overlaps"], b2=best["score_same_topic"],
                c_count=best["empty_seats"], c=best["score_empty_seats"],
            ))
        self.root.update_idletasks()

    # ------------------------------------------------------------------
    # Preview
    # ------------------------------------------------------------------
    def _show_timetable_preview(self):
        if not self.last_timetables:
            messagebox.showwarning(t("gui.warning_title"), t("gui.no_result"), parent=self.root)
            return

        win = tk.Toplevel(self.root)
        win.title(t("gui.timetable_title"))
        win.geometry("960x440")

        if len(self.last_timetables) == 1:
            self._fill_timetable_tree(win, next(iter(self.last_timetables.values())))
            return

        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True)
        for name in sorted(self.last_timetables):
            page = ttk.Frame(notebook)
            notebook.add(page, text=t("gui.candidate_tab", n=name.replace("Result_Timetable_", "")))
            self._fill_timetable_tree(page, self.last_timetables[name])

    @staticmethod
    def _fill_timetable_tree(parent, frame: pd.DataFrame):
        columns = list(frame.columns)
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=140, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for _, row in frame.iterrows():
            tree.insert("", tk.END, values=[row[c] for c in columns])

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    # Opening results
    # ------------------------------------------------------------------
    def _open_output_file(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            try:
                engine.open_file_or_folder(self.last_output_path)
            except Exception as exc:
                messagebox.showerror(t("gui.error_title"), t("gui.open_file_failed", error=exc),
                                     parent=self.root)
        else:
            messagebox.showwarning(t("gui.warning_title"), t("gui.output_not_found"), parent=self.root)

    def _open_output_folder(self):
        target = self.last_output_path or self.output_file_var.get()
        folder = os.path.dirname(os.path.abspath(target)) if target else os.getcwd()
        if os.path.exists(folder):
            try:
                engine.open_file_or_folder(folder)
            except Exception as exc:
                messagebox.showerror(t("gui.error_title"), t("gui.open_folder_failed", error=exc),
                                     parent=self.root)
        else:
            messagebox.showwarning(t("gui.warning_title"), t("gui.folder_not_found", folder=folder),
                                   parent=self.root)


def launch_gui(language: str | None = None):
    """Open the window. ``language`` ("ja"/"en") overrides the system setting."""
    msg.set_language(language or msg.detect_language())
    root = tk.Tk()
    ConferenceSchedulerApp(root)
    root.mainloop()
