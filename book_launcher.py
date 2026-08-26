"""
Book Tools Launcher v2.1
-------------------------
Single window that runs all reading-workflow scripts with configurable settings
and the ability to stop running processes.
"""

from __future__ import annotations

import glob
import json
import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
# SCRIPT_DIR holds the code (and launcher_config.json); DATA_DIR holds the
# spreadsheets and logs.  They are the same folder in the old layout and
# separate once the code is cloned out of OneDrive.
SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(r"C:\Users\megha\OneDrive\Documents\Reading")

CONFIG_FILE = SCRIPT_DIR / "launcher_config.json"

TOOLS = [
    {
        "label": "Process Screenshots",
        "subtitle": "Run books.py — extract titles from screenshots",
        "script": "books.py",
        "packages": ["openai", "pandas", "openpyxl", "pillow"],
        "post_install": [],
    },
    {
        "label": "Update Amazon Owned",
        "subtitle": "Run amazon_owned_books.py — mark owned books",
        "script": "amazon_owned_books.py",
        "packages": ["playwright", "pandas", "openpyxl"],
        "post_install": [[sys.executable, "-m", "playwright", "install", "chromium"]],
    },
    {
        "label": "Push to StoryGraph",
        "subtitle": "Run storygraph_to_read.py — add to to-read list",
        "script": "storygraph_to_read.py",
        "packages": ["pandas", "openpyxl", "playwright"],
        "post_install": [[sys.executable, "-m", "playwright", "install", "chromium"]],
        "force_console": True,  # Always show console for progress monitoring
    },
]

SELECTOR_HTML = "my-book-selector.html"
APPLY_SCRIPT = "apply_reading_log.py"

# Default settings
DEFAULT_CONFIG = {
    "books_py": {
        "openai_api_key": "",
        "model": "gpt-4o-mini",
        "max_workers": 4,
        "confidence_threshold": 0.75,
        "test_mode": False,
        "test_limit": 10,
        "enable_metadata_enrichment": True,
        "enable_asin_lookup": True,
    },
    "amazon_owned_books": {
        "start_page": 1,
        "max_pages": 0,  # 0 = all pages
    },
    "storygraph": {
        "max_books": 0,  # 0 = all books
    },
}

# Global process tracking
RUNNING_PROCESSES: dict[str, subprocess.Popen] = {}
PROCESS_LOCK = threading.Lock()


# --------------------------------------------------------------------------- #
# Settings Management
# --------------------------------------------------------------------------- #
def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(config: dict) -> None:
    """Save config to JSON file with error handling."""
    try:
        # Ensure parent directory exists
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(config, indent=2), encoding="utf-8")
        print(f"Config saved to: {CONFIG_FILE}")
    except Exception as e:
        print(f"ERROR saving config: {e}")
        raise


# --------------------------------------------------------------------------- #
# Process Management
# --------------------------------------------------------------------------- #
def track_process(name: str, proc: subprocess.Popen) -> None:
    """Add a process to the tracking dict."""
    with PROCESS_LOCK:
        RUNNING_PROCESSES[name] = proc


def untrack_process(name: str) -> None:
    """Remove a process from tracking."""
    with PROCESS_LOCK:
        RUNNING_PROCESSES.pop(name, None)


def stop_all_processes() -> None:
    """Kill all tracked processes."""
    with PROCESS_LOCK:
        if not RUNNING_PROCESSES:
            messagebox.showinfo("No Running Scripts", "No scripts are currently running.")
            return
        
        count = len(RUNNING_PROCESSES)
        script_names = ", ".join(RUNNING_PROCESSES.keys())
        
        if not messagebox.askyesno(
            "Stop All Scripts",
            f"Stop {count} running script(s)?\n\n{script_names}\n\nThis will terminate them immediately."
        ):
            return
        
        stopped = 0
        for name, proc in list(RUNNING_PROCESSES.items()):
            try:
                proc.terminate()
                proc.wait(timeout=2)
                stopped += 1
            except Exception:
                try:
                    proc.kill()
                    stopped += 1
                except Exception:
                    pass
        
        RUNNING_PROCESSES.clear()
        messagebox.showinfo("Scripts Stopped", f"Stopped {stopped} script(s).")


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def find_python() -> str:
    venv = SCRIPT_DIR / "venv" / "Scripts" / "python.exe"
    if venv.exists():
        return str(venv)
    if shutil.which("py"):
        return "py"
    if shutil.which("python"):
        return "python"
    return sys.executable


def build_runner_command(tool: dict, config: dict) -> str:
    py = find_python()
    script_path = SCRIPT_DIR / tool["script"]
    packages = " ".join(tool["packages"])

    pieces = []
    
    # Set environment variables for books.py
    if tool["script"] == "books.py" and config.get("books_py", {}).get("openai_api_key"):
        api_key = config["books_py"]["openai_api_key"]
        pieces.append(f'set OPENAI_API_KEY={api_key}')
    
    # Force unbuffered output so console shows progress in real-time
    pieces.append('set PYTHONUNBUFFERED=1')
    
    pieces.extend([
        f'cd /d "{SCRIPT_DIR}"',
        f'echo ============================================================',
        f'echo  {tool["label"]}',
        f'echo ============================================================',
        f'"{py}" -m pip install --upgrade pip',
        f'"{py}" -m pip install {packages}',
    ])
    
    for extra in tool["post_install"]:
        pieces.append(" ".join(f'"{p}"' if " " in p else p for p in extra))
    
    # Run with -u flag for unbuffered output
    pieces.append(f'"{py}" -u "{script_path}"')
    pieces.append("echo.")
    pieces.append("echo Finished. Press any key to close...")
    pieces.append("pause >nul")
    return " && ".join(pieces)


def launch_in_console(tool: dict, config: dict, status_label: tk.Label) -> None:
    script_path = SCRIPT_DIR / tool["script"]
    if not script_path.exists():
        messagebox.showerror(
            "Script not found",
            f"Couldn't find:\n{script_path}\n\nMake sure the launcher lives in the "
            f"same folder as your scripts.",
        )
        return
    
    script_name = tool["script"]
    
    # Check if already running
    with PROCESS_LOCK:
        if script_name in RUNNING_PROCESSES:
            messagebox.showwarning(
                "Already Running",
                f"{tool['label']} is already running.\n\nUse 'Stop All Scripts' to terminate it."
            )
            return
    
    cmd = build_runner_command(tool, config)
    
    # Launch and track the process
    proc = subprocess.Popen(
        f'start "{tool["label"]}" cmd /k {cmd}',
        shell=True,
        cwd=str(SCRIPT_DIR),
    )
    
    track_process(script_name, proc)
    update_status_label(status_label)
    
    # Monitor process in background thread
    def monitor():
        proc.wait()
        untrack_process(script_name)
        update_status_label(status_label)
    
    threading.Thread(target=monitor, daemon=True).start()


def update_status_label(status_label: tk.Label) -> None:
    """Update the status label showing running scripts."""
    with PROCESS_LOCK:
        count = len(RUNNING_PROCESSES)
        if count == 0:
            status_label.config(text="No scripts running")
        elif count == 1:
            name = list(RUNNING_PROCESSES.keys())[0]
            status_label.config(text=f"Running: {name}")
        else:
            status_label.config(text=f"Running: {count} scripts")


def open_selector() -> None:
    webbrowser.open("https://meghanareed.github.io/my-book-selector/")


def open_folder() -> None:
    if sys.platform.startswith("win"):
        os.startfile(SCRIPT_DIR)  # noqa: SIM115
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(SCRIPT_DIR)])
    else:
        subprocess.Popen(["xdg-open", str(SCRIPT_DIR)])


def apply_decisions(status_var: tk.StringVar, log_widget: tk.Text, config: dict) -> None:
    apply_path = SCRIPT_DIR / APPLY_SCRIPT
    if not apply_path.exists():
        messagebox.showerror(
            "apply_reading_log.py not found",
            f"Expected:\n{apply_path}",
        )
        return

    downloads = Path.home() / "Downloads"
    initial_dir = str(downloads) if downloads.exists() else str(SCRIPT_DIR)
    json_path = filedialog.askopenfilename(
        title="Pick book-selector-decisions JSON",
        initialdir=initial_dir,
        filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
    )
    if not json_path:
        return

    py = find_python()

    def worker() -> None:
        status_var.set("Applying decisions…")
        log_widget.configure(state="normal")
        log_widget.delete("1.0", tk.END)
        try:
            proc = subprocess.Popen(
                [py, str(apply_path), json_path],
                cwd=str(SCRIPT_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )
            assert proc.stdout is not None
            for line in proc.stdout:
                log_widget.insert(tk.END, line)
                log_widget.see(tk.END)
            proc.wait()
            if proc.returncode == 0:
                status_var.set("Decisions applied [OK]")
            else:
                status_var.set(f"Failed (exit {proc.returncode})")
        except Exception as exc:  # noqa: BLE001
            log_widget.insert(tk.END, f"\nERROR: {exc}\n")
            status_var.set("Failed")
        finally:
            log_widget.configure(state="disabled")

    threading.Thread(target=worker, daemon=True).start()


def run_piped(label: str, script: str, args: list[str],
              status_var: tk.StringVar, log_widget: tk.Text,
              status_label: tk.Label, config: dict) -> None:
    """
    Run a Python script and pipe its stdout/stderr directly into the
    launcher's log widget — same mechanism as apply_decisions.
    Also sets OPENAI_API_KEY from config if present.
    """
    script_path = SCRIPT_DIR / script
    if not script_path.exists():
        messagebox.showerror("Script not found", f"Expected:\n{script_path}")
        return

    py  = find_python()
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"

    # Inject API key if stored in config
    api_key = config.get("books_py", {}).get("openai_api_key", "")
    if api_key:
        env["OPENAI_API_KEY"] = api_key

    def worker() -> None:
        track_process(script, None)  # placeholder so status label updates
        update_status_label(status_label)

        log_widget.configure(state="normal")
        log_widget.delete("1.0", tk.END)
        log_widget.insert(tk.END, f">> {label}\n{'='*50}\n")
        log_widget.see(tk.END)

        try:
            proc = subprocess.Popen(
                [py, "-u", str(script_path)] + args,
                cwd=str(SCRIPT_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=env,
            )
            track_process(script, proc)
            update_status_label(status_label)

            assert proc.stdout is not None
            for line in proc.stdout:
                log_widget.insert(tk.END, line)
                log_widget.see(tk.END)
                log_widget.update_idletasks()   # keep UI responsive while lines stream in

            proc.wait()
            status_var.set(
                f"{label} finished [OK]" if proc.returncode == 0
                else f"{label} failed (exit {proc.returncode})"
            )
            log_widget.insert(tk.END, f"\n{'='*50}\nDone\n")

        except Exception as exc:
            log_widget.insert(tk.END, f"\nERROR: {exc}\n")
            status_var.set(f"{label} — error")
        finally:
            log_widget.configure(state="disabled")
            untrack_process(script)
            update_status_label(status_label)

    threading.Thread(target=worker, daemon=True).start()
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, config: dict):
        super().__init__(parent)
        self.title("Settings")
        self.geometry("600x500")
        self.config = config.copy()
        self.result = None
        
        self.build_ui()
        self.center_window()
        self.transient(parent)
        self.grab_set()
    
    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
    
    def build_ui(self):
        # Notebook for tabs
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Books.py tab
        books_frame = ttk.Frame(notebook, padding=10)
        notebook.add(books_frame, text="📸 Books.py")
        self.build_books_settings(books_frame)
        
        # Amazon tab
        amazon_frame = ttk.Frame(notebook, padding=10)
        notebook.add(amazon_frame, text="📦 Amazon")
        self.build_amazon_settings(amazon_frame)
        
        # StoryGraph tab
        sg_frame = ttk.Frame(notebook, padding=10)
        notebook.add(sg_frame, text="📚 StoryGraph")
        self.build_storygraph_settings(sg_frame)
        
        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side="right")
        ttk.Button(btn_frame, text="Reset to Defaults", command=self.reset_defaults).pack(side="left")
    
    def build_books_settings(self, parent):
        self.books_vars = {}
        
        # API Key
        ttk.Label(parent, text="OpenAI API Key:", font=("", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.books_vars["openai_api_key"] = tk.StringVar(value=self.config.get("books_py", {}).get("openai_api_key", ""))
        api_entry = ttk.Entry(parent, textvariable=self.books_vars["openai_api_key"], width=50, show="*")
        api_entry.grid(row=0, column=1, sticky="ew", pady=5)
        
        # Model
        ttk.Label(parent, text="Model:", font=("", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        self.books_vars["model"] = tk.StringVar(value=self.config.get("books_py", {}).get("model", "gpt-4o-mini"))
        model_combo = ttk.Combobox(parent, textvariable=self.books_vars["model"], 
                                    values=["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-3.5-turbo"])
        model_combo.grid(row=1, column=1, sticky="w", pady=5)
        
        # Max Workers
        ttk.Label(parent, text="Max Workers (threads):", font=("", 9, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        self.books_vars["max_workers"] = tk.IntVar(value=self.config.get("books_py", {}).get("max_workers", 4))
        ttk.Spinbox(parent, from_=1, to=16, textvariable=self.books_vars["max_workers"], width=10).grid(row=2, column=1, sticky="w", pady=5)
        
        # Confidence Threshold
        ttk.Label(parent, text="Confidence Threshold:", font=("", 9, "bold")).grid(row=3, column=0, sticky="w", pady=5)
        self.books_vars["confidence_threshold"] = tk.DoubleVar(value=self.config.get("books_py", {}).get("confidence_threshold", 0.75))
        ttk.Spinbox(parent, from_=0.0, to=1.0, increment=0.05, textvariable=self.books_vars["confidence_threshold"], width=10).grid(row=3, column=1, sticky="w", pady=5)
        
        # Test Mode
        ttk.Label(parent, text="Test Mode:", font=("", 9, "bold")).grid(row=4, column=0, sticky="w", pady=5)
        self.books_vars["test_mode"] = tk.BooleanVar(value=self.config.get("books_py", {}).get("test_mode", False))
        ttk.Checkbutton(parent, text="Enable (processes only test_limit images)", 
                       variable=self.books_vars["test_mode"]).grid(row=4, column=1, sticky="w", pady=5)
        
        # Test Limit
        ttk.Label(parent, text="Test Limit (when test mode on):", font=("", 9, "bold")).grid(row=5, column=0, sticky="w", pady=5)
        self.books_vars["test_limit"] = tk.IntVar(value=self.config.get("books_py", {}).get("test_limit", 10))
        ttk.Spinbox(parent, from_=1, to=100, textvariable=self.books_vars["test_limit"], width=10).grid(row=5, column=1, sticky="w", pady=5)
        
        # Metadata Enrichment
        ttk.Label(parent, text="Metadata Enrichment:", font=("", 9, "bold")).grid(row=6, column=0, sticky="w", pady=5)
        self.books_vars["enable_metadata_enrichment"] = tk.BooleanVar(value=self.config.get("books_py", {}).get("enable_metadata_enrichment", True))
        ttk.Checkbutton(parent, text="Enable", variable=self.books_vars["enable_metadata_enrichment"]).grid(row=6, column=1, sticky="w", pady=5)
        
        # ASIN Lookup
        self.books_vars["enable_asin_lookup"] = tk.BooleanVar(value=self.config.get("books_py", {}).get("enable_asin_lookup", True))
        ttk.Checkbutton(parent, text="Enable ASIN Lookup", variable=self.books_vars["enable_asin_lookup"]).grid(row=7, column=1, sticky="w", pady=2)
        
        parent.columnconfigure(1, weight=1)
    
    def build_amazon_settings(self, parent):
        self.amazon_vars = {}
        
        ttk.Label(parent, text="Starting Page Number:", font=("", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.amazon_vars["start_page"] = tk.IntVar(value=self.config.get("amazon_owned_books", {}).get("start_page", 1))
        ttk.Spinbox(parent, from_=1, to=999, textvariable=self.amazon_vars["start_page"], width=10).grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(parent, text="Max Pages to Process:", font=("", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        self.amazon_vars["max_pages"] = tk.IntVar(value=self.config.get("amazon_owned_books", {}).get("max_pages", 0))
        ttk.Spinbox(parent, from_=0, to=999, textvariable=self.amazon_vars["max_pages"], width=10).grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(parent, text="(0 = all pages)", font=("", 8), foreground="gray").grid(row=1, column=2, sticky="w", padx=5)
        
        parent.columnconfigure(1, weight=1)
    
    def build_storygraph_settings(self, parent):
        self.sg_vars = {}
        
        ttk.Label(parent, text="Max Books to Add:", font=("", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.sg_vars["max_books"] = tk.IntVar(value=self.config.get("storygraph", {}).get("max_books", 0))
        ttk.Spinbox(parent, from_=0, to=999, textvariable=self.sg_vars["max_books"], width=10).grid(row=0, column=1, sticky="w", pady=5)
        ttk.Label(parent, text="(0 = all books)", font=("", 8), foreground="gray").grid(row=0, column=2, sticky="w", padx=5)
        
        parent.columnconfigure(1, weight=1)
    
    def save(self):
        # Collect all settings
        self.config["books_py"] = {
            k: v.get() for k, v in self.books_vars.items()
        }
        self.config["amazon_owned_books"] = {
            k: v.get() for k, v in self.amazon_vars.items()
        }
        self.config["storygraph"] = {
            k: v.get() for k, v in self.sg_vars.items()
        }
        
        self.result = self.config
        self.destroy()
    
    def cancel(self):
        self.result = None
        self.destroy()
    
    def reset_defaults(self):
        if messagebox.askyesno("Reset Settings", "Reset all settings to defaults?"):
            self.config = DEFAULT_CONFIG.copy()
            # Rebuild UI with defaults
            for widget in self.winfo_children():
                widget.destroy()
            self.build_ui()


# --------------------------------------------------------------------------- #
# Main UI
# --------------------------------------------------------------------------- #
def build_ui(config: dict) -> tk.Tk:
    root = tk.Tk()
    root.title("Book Tools Launcher v2.1")
    root.geometry("980x620")
    root.minsize(800, 500)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    elif "clam" in style.theme_names():
        style.theme_use("clam")

    style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
    style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#555")
    style.configure("Tool.TButton", font=("Segoe UI", 11), padding=(10, 8))
    style.configure("Stop.TButton", font=("Segoe UI", 10, "bold"))
    style.configure("LogHead.TLabel", font=("Segoe UI", 10, "bold"), foreground="#333")

    outer = ttk.Frame(root, padding=14)
    outer.pack(fill="both", expand=True)

    # ── Header (full width) ───────────────────────────────────────────────────
    header_frame = ttk.Frame(outer)
    header_frame.pack(fill="x", pady=(0, 10))

    ttk.Label(header_frame, text="📚 Book Tools", style="Title.TLabel").pack(side="left")

    status_label = ttk.Label(header_frame, text="No scripts running",
                             font=("Segoe UI", 8), foreground="#666")
    status_label.pack(side="left", padx=15)

    def open_settings():
        dialog = SettingsDialog(root, config)
        root.wait_window(dialog)
        if dialog.result:
            config.update(dialog.result)
            save_config(config)
            messagebox.showinfo("Settings Saved", "Settings have been saved!")

    ttk.Button(header_frame, text="Stop All Scripts",
               style="Stop.TButton",
               command=stop_all_processes).pack(side="right", padx=5)
    ttk.Button(header_frame, text="Settings", command=open_settings).pack(side="right")

    ttk.Label(outer, text=f"Folder: {SCRIPT_DIR}", style="Sub.TLabel").pack(
        anchor="w", pady=(0, 8))

    # ── Two-panel body ────────────────────────────────────────────────────────
    body = ttk.Frame(outer)
    body.pack(fill="both", expand=True)
    body.columnconfigure(0, weight=0, minsize=340)   # left: buttons, fixed width
    body.columnconfigure(1, weight=1)                 # right: log, expands
    body.rowconfigure(0, weight=1)

    # Left panel — buttons
    left = ttk.Frame(body)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

    # Right panel — log
    right = ttk.Frame(body, relief="sunken", borderwidth=1)
    right.grid(row=0, column=1, sticky="nsew")
    right.rowconfigure(1, weight=1)
    right.columnconfigure(0, weight=1)

    status_var = tk.StringVar(value="")
    ttk.Label(right, text="Output Log", style="LogHead.TLabel").grid(
        row=0, column=0, sticky="w", padx=8, pady=(6, 2))
    ttk.Label(right, textvariable=status_var, style="Sub.TLabel").grid(
        row=0, column=1, sticky="e", padx=8, pady=(6, 2))

    log_widget = tk.Text(right, wrap="word", state="disabled",
                         font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
                         insertbackground="white", relief="flat")
    log_widget.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=4, pady=(0, 4))

    log_scroll = ttk.Scrollbar(right, orient="vertical", command=log_widget.yview)
    log_scroll.grid(row=1, column=2, sticky="ns", pady=(0, 4))
    log_widget.configure(yscrollcommand=log_scroll.set)

    # Clear log button
    def clear_log():
        log_widget.configure(state="normal")
        log_widget.delete("1.0", tk.END)
        log_widget.configure(state="disabled")
        status_var.set("")
    ttk.Button(right, text="Clear", width=7, command=clear_log).grid(
        row=2, column=0, sticky="w", padx=6, pady=(2, 6))

    # ── Left panel contents ───────────────────────────────────────────────────
    container = left  # alias so all existing code below just works

    # Tool buttons
    for tool in TOOLS:
        frame = ttk.Frame(container)
        frame.pack(fill="x", pady=4)
        ttk.Button(
            frame,
            text=tool["label"],
            style="Tool.TButton",
            width=24,
            command=lambda t=tool: launch_in_console(t, config, status_label),
        ).pack(side="left")
        ttk.Label(
            frame,
            text=tool["subtitle"],
            style="Sub.TLabel",
            wraplength=200,
            justify="left",
        ).pack(side="left", padx=8)

    ttk.Separator(container, orient="horizontal").pack(fill="x", pady=10)

    # Selector + reading log
    sel_frame = ttk.Frame(container)
    sel_frame.pack(fill="x", pady=4)
    ttk.Button(
        sel_frame,
        text="Open Book Selector",
        style="Tool.TButton",
        width=24,
        command=open_selector,
    ).pack(side="left")
    ttk.Label(
        sel_frame,
        text="Open web selector",
        style="Sub.TLabel",
    ).pack(side="left", padx=8)

    apply_frame = ttk.Frame(container)
    apply_frame.pack(fill="x", pady=4)
    ttk.Button(
        apply_frame,
        text="Apply Selector Decisions",
        style="Tool.TButton",
        width=24,
        command=lambda: apply_decisions(status_var, log_widget, config),
    ).pack(side="left")
    ttk.Label(
        apply_frame,
        text="Import decisions JSON",
        style="Sub.TLabel",
    ).pack(side="left", padx=8)

    reenrich_frame = ttk.Frame(container)
    reenrich_frame.pack(fill="x", pady=4)
    ttk.Button(
        reenrich_frame,
        text="Fill Missing Metadata",
        style="Tool.TButton",
        width=24,
        command=lambda: run_piped(
            "Fill Missing Metadata",
            "reenrich_existing.py",
            [],
            status_var,
            log_widget,
            status_label,
            config,
        ),
    ).pack(side="left")
    ttk.Label(
        reenrich_frame,
        text="Fill PageCount, Genre, Tropes etc.",
        style="Sub.TLabel",
    ).pack(side="left", padx=8)

    folder_frame = ttk.Frame(container)
    folder_frame.pack(fill="x", pady=4)
    ttk.Button(
        folder_frame,
        text="Open Reading Folder",
        style="Tool.TButton",
        width=24,
        command=open_folder,
    ).pack(side="left")
    ttk.Label(
        folder_frame,
        text="Open folder in Explorer",
        style="Sub.TLabel",
    ).pack(side="left", padx=8)

    logs_frame = ttk.Frame(container)
    logs_frame.pack(fill="x", pady=4)

    def open_logs():
        """Open the most recent log file from the logs/ subfolder."""
        import glob

        # The scripts write logs beside the spreadsheets (DATA_DIR), not beside
        # the code.  Fall back to SCRIPT_DIR for the old shared-folder layout.
        logs_dir = DATA_DIR / "logs"
        if not logs_dir.exists():
            logs_dir = SCRIPT_DIR / "logs"
        amazon_pattern = str(logs_dir / "amazon_scrape_*.log")
        storygraph_pattern = str(logs_dir / "storygraph_*.log")
        reenrich_pattern = str(logs_dir / "reenrich_*.log")

        amazon_logs = glob.glob(amazon_pattern)
        storygraph_logs = glob.glob(storygraph_pattern)
        reenrich_logs = glob.glob(reenrich_pattern)
        all_logs = amazon_logs + storygraph_logs + reenrich_logs

        if all_logs:
            newest_log = Path(max(all_logs, key=os.path.getmtime))
            if sys.platform.startswith("win"):
                os.startfile(newest_log)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(newest_log)])
            else:
                subprocess.Popen(["xdg-open", str(newest_log)])
        else:
            messagebox.showinfo(
                "No Logs Found",
                f"No log files found in:\n{logs_dir}\n\n"
                "Log files are created when you run:\n"
                "• Update Amazon Owned (creates amazon_scrape_*.log)\n"
                "• Push to StoryGraph (creates storygraph_*.log)\n"
                "• Fill Missing Metadata (creates reenrich_*.log)\n\n"
                "Run one of these scripts first, then check logs.",
                parent=root
            )

    ttk.Button(
        logs_frame,
        text="View Latest Log",
        style="Tool.TButton",
        width=24,
        command=open_logs,
    ).pack(side="left")
    ttk.Label(
        logs_frame,
        text="Open most recent log file",
        style="Sub.TLabel",
    ).pack(side="left", padx=8)

    return root


def main() -> None:
    config = load_config()
    root = build_ui(config)
    root.mainloop()


if __name__ == "__main__":
    main()
