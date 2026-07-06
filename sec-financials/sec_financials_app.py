#!/usr/bin/env python3
"""
SEC Financials — desktop app
============================
A tiny window: type a ticker, click Generate, and an Excel workbook with the
company's income statement, balance sheet, and cash flow (from its latest 10-K)
is created next to this app and opened automatically.

For end users: just double-click the app — no Python or command line needed.
(Built into a single .exe with build_exe.bat; see that file.)
"""

import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

from sec_financials import generate, IDENTITY

EMAIL_HINT = "Your name and email — e.g. Jane Smith jane@company.com"


def _add_placeholder(entry, text):
    """Show grey hint text in an Entry until the user types."""
    entry.insert(0, text)
    entry.config(fg="grey")

    def focus_in(_):
        if entry.get() == text:
            entry.delete(0, "end")
            entry.config(fg="black")

    def focus_out(_):
        if not entry.get().strip():
            entry.insert(0, text)
            entry.config(fg="grey")

    entry.bind("<FocusIn>", focus_in)
    entry.bind("<FocusOut>", focus_out)


def app_dir():
    """Folder to save output into: next to the .exe (frozen) or the script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class App:
    def __init__(self, root):
        self.root = root
        root.title("SEC Financials")
        root.geometry("500x400")
        root.resizable(False, False)

        pad = {"padx": 14, "pady": 6}
        tk.Label(root, text="SEC Financial Statement Puller",
                 font=("Segoe UI", 14, "bold"), fg="#1F3864").pack(pady=(16, 2))
        tk.Label(root, text="Type a stock ticker and click Generate.",
                 font=("Segoe UI", 9), fg="#555").pack()

        frm = tk.Frame(root); frm.pack(**pad)
        tk.Label(frm, text="Ticker:", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="e", padx=4)
        self.ticker = tk.Entry(frm, font=("Segoe UI", 12), width=14, justify="center")
        self.ticker.grid(row=0, column=1, padx=4)
        self.ticker.insert(0, "GTM")
        self.ticker.bind("<Return>", lambda e: self.on_generate())

        tk.Label(frm, text="Name & email:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="e", padx=4, pady=(6, 0))
        self.email = tk.Entry(frm, font=("Segoe UI", 9), width=34)
        self.email.grid(row=1, column=1, padx=4, pady=(6, 0))
        _add_placeholder(self.email, EMAIL_HINT)

        self.btn = tk.Button(root, text="Generate", font=("Segoe UI", 11, "bold"),
                             bg="#1F3864", fg="white", width=16, command=self.on_generate)
        self.btn.pack(pady=6)

        self.log = tk.Text(root, height=9, width=58, font=("Consolas", 9),
                           state="disabled", bg="#F5F6FA", relief="flat")
        self.log.pack(padx=14, pady=(6, 4))

        tk.Label(root, text="SEC requires your email on requests. Data source: SEC EDGAR (free).",
                 font=("Segoe UI", 8), fg="#888").pack()

    def say(self, msg):
        self.root.after(0, self._append, msg)

    def _append(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def on_generate(self):
        ticker = self.ticker.get().strip().upper()
        if not ticker:
            messagebox.showwarning("SEC Financials", "Please type a ticker (e.g. GTM).")
            return
        email = self.email.get().strip()
        if email == EMAIL_HINT:
            email = ""
        if "@" not in email:
            messagebox.showwarning("SEC Financials",
                                   "Please enter your name and email.\nSEC requires it on each request, "
                                   "e.g. 'Jane Smith jane@company.com'.")
            self.email.focus_set()
            return
        self.btn.configure(state="disabled", text="Working…")
        self.log.configure(state="normal"); self.log.delete("1.0", "end"); self.log.configure(state="disabled")
        threading.Thread(target=self._worker, args=(ticker, email), daemon=True).start()

    def _worker(self, ticker, email):
        try:
            path = generate(ticker, out_dir=app_dir(), identity=email or IDENTITY, status=self.say)
            self.say("Done.")
            self.root.after(0, self._done, path)
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, self._error, str(e))

    def _done(self, path):
        self.btn.configure(state="normal", text="Generate")
        try:
            os.startfile(path)          # opens the workbook (Windows)
        except Exception:
            pass
        messagebox.showinfo("SEC Financials", f"Created:\n{os.path.basename(path)}\n\nSaved in:\n{os.path.dirname(path)}")

    def _error(self, msg):
        self.btn.configure(state="normal", text="Generate")
        messagebox.showerror("SEC Financials", "Could not generate the workbook.\n\n" + msg)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
