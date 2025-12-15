import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional
import threading
import json
from pathlib import Path
import time
import requests

# --- HTTP control (use httpx Digest Auth to talk to DLI) ---
try:
    import httpx
    HAS_HTTPX = True
except Exception:
    httpx = None
    HAS_HTTPX = False

# === Defaults for DLI device config ===
DEFAULT_DLI_HOST = "192.168.0.100"
DEFAULT_DLI_USER = "admin"
DEFAULT_DLI_PASS = "1234"
DLI_TIMEOUT = 5  # seconds
# The DLI REST endpoint in the user's example uses 0-based channels.
OUTLET_BASE = 0

# === Config file location ===
CONFIG_PATH = Path.home() / ".cue_switchboard.json"
SEQUENCES_DIR = Path(__file__).parent / "sequences"

COLUMNS = ["Cue", "Order", "Sequence", "Delay"] + [f"Switch{i}" for i in range(1, 9)]
SWITCH_COLUMNS = set(COLUMNS[4:])  # Switch1..Switch8


# === Domain callback that receives the selected cue as a dict ===
def on_cue_change(cue: dict):
    print("Cue changed:", cue)


class CueEditorDialog(tk.Toplevel):
    """Used for both Add and Edit. Pass `initial` for edit mode."""
    def __init__(self, master, title: str, initial: Optional[Dict] = None, suggested_order: Optional[int] = None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        self.var_cue = tk.StringVar(value=(initial.get("Cue") if initial else ""))
        default_order = (
            initial.get("Order")
            if initial and initial.get("Order") not in ("", None)
            else (suggested_order if suggested_order is not None else "")
        )
        self.var_order = tk.StringVar(value=str(default_order) if default_order != "" else "")
        self.var_sequence = tk.StringVar(value=(initial.get("Sequence") if initial else ""))
        self.switch_vars = []
        for i in range(8):
            val = bool(initial.get(f"Switch{i+1}", False)) if initial else False
            self.switch_vars.append(tk.BooleanVar(value=val))

        frm = ttk.Frame(self, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Cue name:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=4)
        cue_entry = ttk.Entry(frm, textvariable=self.var_cue, width=28)
        cue_entry.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Cue order:").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.var_order, width=10).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Sequence:").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.var_sequence, width=40).grid(row=2, column=1, sticky="w", pady=4)

        # Delay field (milliseconds)
        self.var_delay = tk.StringVar(value=str(initial.get("Delay") if initial and initial.get("Delay") is not None else ""))
        ttk.Label(frm, text="Delay (ms):").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.var_delay, width=12).grid(row=3, column=1, sticky="w", pady=4)

        sw_frame = ttk.LabelFrame(frm, text="Cue switches", padding=(10, 8))
        sw_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        for i in range(8):
            ttk.Checkbutton(sw_frame, text=f"Switch{i+1}", variable=self.switch_vars[i])\
            .grid(row=i // 4, column=i % 4, sticky="w", padx=6, pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=self._on_cancel).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(btns, text="OK", command=self._on_ok).grid(row=0, column=1)

        self.result = None
        self.bind("<Return>", lambda e: self._on_ok())
        self.bind("<Escape>", lambda e: self._on_cancel())
        cue_entry.focus()

    def _on_ok(self):
        cue = self.var_cue.get().strip()
        if not cue:
            messagebox.showwarning("Missing Cue", "Please enter a cue name.")
            return

        order_text = self.var_order.get().strip()
        order_val = ""
        if order_text:
            try:
                order_val = int(order_text)
            except ValueError:
                messagebox.showwarning("Invalid Order", "Cue order must be an integer (or leave blank).")
                return

        # parse delay (optional integer)
        delay_text = self.var_delay.get().strip() if hasattr(self, 'var_delay') else ""
        delay_val = ""
        if delay_text:
            try:
                delay_val = int(delay_text)
            except ValueError:
                messagebox.showwarning("Invalid Delay", "Delay must be an integer (milliseconds) or leave blank.")
                return

        cue_dict = {"Cue": cue, "Order": order_val, "Sequence": self.var_sequence.get().strip(), "Delay": delay_val}
        for i, var in enumerate(self.switch_vars, 1):
            cue_dict[f"Switch{i}"] = bool(var.get())

        self.result = cue_dict
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


class StepEditorDialog(tk.Toplevel):
    """Dialog to add/edit a single sequence step: switch (int), position (bool), delay (int)."""
    def __init__(self, master, title: str, initial: Optional[Dict] = None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        self.var_switch = tk.StringVar(value=str(initial.get("switch") if initial else ""))
        self.var_position = tk.BooleanVar(value=bool(initial.get("position", False)) if initial else False)
        self.var_delay = tk.StringVar(value=str(initial.get("delay") if initial else ""))

        frm = ttk.Frame(self, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Switch:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.var_switch, width=12).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Position:").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Checkbutton(frm, variable=self.var_position, text="On/True").grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Delay (ms):").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.var_delay, width=12).grid(row=2, column=1, sticky="w", pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Cancel", command=self._on_cancel).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(btns, text="OK", command=self._on_ok).grid(row=0, column=1)

        self.result = None
        self.bind("<Return>", lambda e: self._on_ok())
        self.bind("<Escape>", lambda e: self._on_cancel())

    def _on_ok(self):
        # Validate switch and delay as integers
        sw_text = self.var_switch.get().strip()
        delay_text = self.var_delay.get().strip()
        try:
            sw = int(sw_text)
        except Exception:
            messagebox.showwarning("Invalid Switch", "Switch must be an integer.")
            return
        try:
            delay = int(delay_text) if delay_text != "" else 0
        except Exception:
            messagebox.showwarning("Invalid Delay", "Delay must be an integer (milliseconds).")
            return

        self.result = {"switch": sw, "position": bool(self.var_position.get()), "delay": delay}
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


class SequenceEditorDialog(tk.Toplevel):
    """Dialog to create or edit a sequence file (collection of steps).

    Returns a tuple (filename, steps) in `self.result` on OK, where `steps` is a
    list of dicts: {"switch":int, "position":bool, "delay":int}.
    """
    def __init__(self, master, title: str, initial_file: Optional[str] = None, initial_steps: Optional[list] = None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        self.var_filename = tk.StringVar(value=(initial_file if initial_file else "new_sequence"))

        frm = ttk.Frame(self, padding=8)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="File name:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(frm, textvariable=self.var_filename, width=36).grid(row=0, column=1, sticky="w", pady=4)

        # Steps Tree
        cols = ("Switch", "Position", "Delay")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse", height=8)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=80, anchor="center")
        self.tree.column("Switch", width=80, anchor="center")
        self.tree.column("Position", width=80, anchor="center")
        self.tree.column("Delay", width=100, anchor="e")
        self.tree.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(4, 0))

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(btn_frame, text="Add Step", command=self._on_add_step).grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="Edit Step", command=self._on_edit_step).grid(row=0, column=1, padx=4)
        ttk.Button(btn_frame, text="Remove Step", command=self._on_remove_step).grid(row=0, column=2, padx=4)

        # Save/Cancel
        ctl = ttk.Frame(frm)
        ctl.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(ctl, text="Cancel", command=self._on_cancel).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(ctl, text="Save", command=self._on_save).grid(row=0, column=1)

        self.result = None

        # Load initial steps
        self._load_steps(initial_steps or [])

        self.bind("<Return>", lambda e: self._on_save())
        self.bind("<Escape>", lambda e: self._on_cancel())

    def _load_steps(self, steps: list):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for s in steps:
            sw = s.get("switch")
            pos = s.get("position")
            delay = s.get("delay")
            vals = [str(sw), "True" if pos else "False", str(delay)]
            self.tree.insert("", "end", values=vals)

    def _gather_steps(self):
        out = []
        for iid in self.tree.get_children():
            v = self.tree.item(iid, "values")
            try:
                sw = int(v[0])
            except Exception:
                sw = 0
            pos = str(v[1]).strip().lower() in {"1", "true", "t", "yes", "y", "on"}
            try:
                delay = int(v[2])
            except Exception:
                delay = 0
            out.append({"switch": sw, "position": pos, "delay": delay})
        return out

    def _on_add_step(self):
        dlg = StepEditorDialog(self, "Add Step")
        self.wait_window(dlg)
        if dlg.result is None:
            return
        step = dlg.result
        vals = [str(step["switch"]), "True" if step["position"] else "False", str(step["delay"])]
        self.tree.insert("", "end", values=vals)

    def _on_edit_step(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Edit Step", "Please select a step to edit.")
            return
        iid = sel[0]
        vals = self.tree.item(iid, "values")
        initial = {"switch": int(vals[0]) if vals[0] else 0, "position": vals[1].lower() in {"true","1","t","y"}, "delay": int(vals[2]) if vals[2] else 0}
        dlg = StepEditorDialog(self, "Edit Step", initial=initial)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        step = dlg.result
        new_vals = [str(step["switch"]), "True" if step["position"] else "False", str(step["delay"])]
        self.tree.item(iid, values=new_vals)

    def _on_remove_step(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Remove Step", "Please select a step to remove.")
            return
        for iid in sel:
            self.tree.delete(iid)

    def _on_save(self):
        name = self.var_filename.get().strip()
        if not name:
            messagebox.showwarning("Missing Name", "Please enter a filename for the sequence.")
            return
        # strip extension
        if name.lower().endswith('.xml'):
            name = name[:-4]
        steps = self._gather_steps()
        self.result = (name, steps)
        self.destroy()

    def _on_cancel(self):
        self.result = None
        self.destroy()


class CueTableApp(tk.Tk):
    def __init__(self, cue_change_callback=on_cue_change, rows=None):
        super().__init__()
        self.title("Cue Switchboard")
        self.geometry("1150x560")

        self.cue_change_callback = cue_change_callback
        self.row_data_by_iid: Dict[str, Dict] = {}
        self.current_file: Optional[str] = None
        self._inline_editor: Optional[tk.Entry] = None  # for in-place cell edits

        # HTTP/DLI status
        self.dli_status = tk.StringVar(
            value="DLI (HTTP): ready" if HAS_HTTPX else "DLI (HTTP): httpx not installed"
        )

        # --- Debounce/async state ---
        self._pending_send_id: Optional[str] = None
        self._send_seq: int = 0  # increments per send; used to ignore stale results
        self._is_closing: bool = False  # guard to stop scheduling after close begins

        # --- Load config (creates defaults if missing) ---
        self.config_data = self._load_config()
        # Device settings from config
        self.device_host = self.config_data["device"]["host"]
        self.device_user = self.config_data["device"]["user"]
        self.device_pass = self.config_data["device"]["pass"]
        # Last selected index (persisted)
        self._last_selected_index: Optional[int] = self.config_data.get("last_selected_index", 0)

        # Menus
        self._build_menu()

        # Treeview (grid)
        self.tree = ttk.Treeview(self, columns=COLUMNS, show="headings", selectmode="browse")
        self.tree.configure(takefocus=True)  # allow keyboard focus
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbars
        y_scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        # Headings & widths
        for col in COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center", stretch=True)
        self.tree.column("Cue", width=240, anchor="w")
        self.tree.column("Order", width=90, anchor="e")
        self.tree.column("Sequence", width=160, anchor="center")
        self.tree.column("Delay", width=90, anchor="center")

        # Layout stretch
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # Selection and editing bindings
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_double_click_cell)    # edit Cue/Order
        self.tree.bind("<Button-1>", self._on_single_click_cell)    # toggle switches

        # Status bar (DLI status)
        status_bar = ttk.Frame(self)
        status_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.status_label = ttk.Label(status_bar, anchor="w", textvariable=self.dli_status)
        self.status_label.pack(fill="x")

        # ---------------- No demo rows ----------------
        # If config has a last file path and it exists, auto-load it and try to restore selection index.
        last = self.config_data.get("last_xml_path")
        if last and Path(last).exists():
            try:
                rows_from_file = self._read_xml(last)
                self.load_rows(rows_from_file)
                self.current_file = last
                self._update_title()
                self._select_row_by_index(self._last_selected_index)
                # Ensure keyboard focus lands on the tree after UI settles
                self.after(0, self.tree.focus_set)
            except Exception as e:
                messagebox.showwarning("Load Last File", f"Could not load last XML:\n{e}")

        # Hook window close
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ------------------ Config load/save ------------------
    def _default_config(self) -> Dict:
        return {
            "device": {
                "host": DEFAULT_DLI_HOST,
                "user": DEFAULT_DLI_USER,
                "pass": DEFAULT_DLI_PASS
            },
            "last_xml_path": None,
            "last_selected_index": 0
        }

    def _load_config(self) -> Dict:
        cfg = self._default_config()
        try:
            if CONFIG_PATH.exists():
                with CONFIG_PATH.open("r", encoding="utf-8") as f:
                    on_disk = json.load(f)
                if isinstance(on_disk, dict):
                    dev = on_disk.get("device", {})
                    cfg["device"]["host"] = dev.get("host", cfg["device"]["host"])
                    cfg["device"]["user"] = dev.get("user", cfg["device"]["user"])
                    cfg["device"]["pass"] = dev.get("pass", cfg["device"]["pass"])
                    cfg["last_xml_path"] = on_disk.get("last_xml_path", cfg["last_xml_path"])
                    cfg["last_selected_index"] = on_disk.get("last_selected_index", cfg["last_selected_index"])
        except Exception as e:
            messagebox.showwarning("Config", f"Error reading config; using defaults.\n{e}")

        # Ensure file exists on first run
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            if not CONFIG_PATH.exists():
                with CONFIG_PATH.open("w", encoding="utf-8") as f:
                    json.dump(cfg, f, indent=2)
        except Exception:
            pass

        return cfg

    def _save_config(self):
        data = {
            "device": {
                "host": self.device_host,
                "user": self.device_user,
                "pass": self.device_pass
            },
            "last_xml_path": self.current_file,
            "last_selected_index": self._last_selected_index if self._last_selected_index is not None else 0
        }
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with CONFIG_PATH.open("w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            messagebox.showwarning("Config", f"Could not save config:\n{e}")

    # ------------------ Safe Tk scheduling ------------------
    def _safe_after(self, ms: int, func):
        """Schedule func if the app is still alive; ignore if closing/destroyed."""
        if self._is_closing or not self.winfo_exists():
            return
        try:
            return self.after(ms, func)
        except Exception:
            # Happens if mainloop is exiting; ignore
            return

    # ------------------ Busy cursor helpers ------------------
    def _set_busy(self, busy: bool):
        """Set a wait cursor during network activity."""
        cursor = "watch" if busy else ""
        try:
            self.configure(cursor=cursor)
            self.tree.configure(cursor=cursor)
        except Exception:
            pass
        self.update_idletasks()


    # ------------------ HTTP send (async) ------------------
    def _build_pairs_from_cue(self, cue: Dict):
        """Return list of [channel, state] per Switch1..Switch8."""
        pairs = []
        for i in range(1, 9):
            tag = f"Switch{i}"
            state = bool(cue.get(tag, False))
            channel = OUTLET_BASE + (i - 1)
            pairs.append([channel, state])
        return pairs

    def _post_cue_sync(self, cue: Dict) -> str:
        """Blocking HTTP POST; returns short status message. Raises on failure."""
        if not HAS_HTTPX:
            raise RuntimeError("httpx is not installed. Cannot send request.")

        # Gathers the states from the datatable
        payload = [self._build_pairs_from_cue(cue)]  # [[[ch,state]...]] as per spec
        
        # Also capture the sequence name/string from the table for use by callers
        sequence = (cue.get("Sequence") or "").strip()
        # If a sequence is specified, resolve and execute the sequence file instead of the single payload
        if sequence:
            seq_filename = sequence if sequence.lower().endswith('.xml') else sequence + '.xml'
            seq_path = SEQUENCES_DIR / seq_filename
            if seq_path.exists():
                execute_sequence_from_file(
                    str(seq_path),
                    base_url=f"http://{self.device_host}",
                    username=self.device_user,
                    password=self.device_pass,
                    timeout=DLI_TIMEOUT,
                )
            else:
                raise FileNotFoundError(f"Sequence file not found: {seq_path}")


        # Proceed with the states for this row
        url = f"http://{self.device_host}/restapi/relay/set_outlet_transient_states/"
        headers = {"X-CSRF": "x", "Content-Type": "application/json"}
        auth = httpx.DigestAuth(self.device_user, self.device_pass)

        resp = httpx.post(url, headers=headers, json=payload, auth=auth, timeout=DLI_TIMEOUT)
        return f"{resp.status_code} — {resp.text[:120]}"

    def _begin_send_selected(self, iid: str):
        """Start background send for the given iid, with busy cursor and stale-response ignore."""
        if self._is_closing or not self.winfo_exists():
            return

        cue = self.row_data_by_iid.get(iid)
        if not cue:
            return

        self._set_busy(True)
        self._pending_send_id = None
        self._send_seq += 1
        my_seq = self._send_seq

        def worker():
            try:
                msg = self._post_cue_sync(cue)
                ok = True
            except Exception as e:
                msg = f"{e}"
                ok = False
            # marshal back to main thread (safely)
            self._safe_after(0, lambda: self._finish_send(my_seq, ok, msg, cue))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_send(self, seq: int, ok: bool, msg: str, cue: Dict):
        """Finalize UI after background send. Ignore stale/out-of-order replies."""
        if self._is_closing or not self.winfo_exists():
            return
        if seq != self._send_seq:
            return
        name = cue.get("Cue", "")
        self.dli_status.set((f"DLI (HTTP): {msg} — {name}") if ok else (f"DLI (HTTP) error: {msg} — {name}"))
        self._set_busy(False)

    # ------------------ Debounced selection -> send ------------------
    def _debounced_send_for_iid(self, iid: str, delay_ms: int = 120):
        """Schedule a send a moment after selection so UI can update first."""
        if self._pending_send_id is not None:
            try:
                self.after_cancel(self._pending_send_id)
            except Exception:
                pass
            self._pending_send_id = None
        self._pending_send_id = self._safe_after(delay_ms, lambda: self._begin_send_selected(iid))

    def on_closing(self):
        # Prevent new callbacks from being scheduled; cancel pending debounce
        self._is_closing = True
        try:
            if self._pending_send_id:
                self.after_cancel(self._pending_send_id)
        except Exception:
            pass
        # Save config (device + last file path + last index) before exit
        self._save_config()
        try:
            self.destroy()
        except Exception:
            pass

    # ------------------ Menu ------------------
    def _build_menu(self):
        menubar = tk.Menu(self)

        # File
        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="Open…", command=self._menu_open, accelerator="Ctrl+O")
        file_menu.add_command(label="Save…", command=self._menu_save_as, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._menu_exit, accelerator="Ctrl+Q")
        menubar.add_cascade(label="File", menu=file_menu)

        # Cue
        cue_menu = tk.Menu(menubar, tearoff=False)
        cue_menu.add_command(label="Add New Cue…", command=self._menu_add_new, accelerator="Ctrl+N")
        cue_menu.add_command(label="Edit Selected Cue…", command=self._menu_edit_selected, accelerator="Ctrl+E")
        menubar.add_cascade(label="Cue", menu=cue_menu)

        # Sequence
        seq_menu = tk.Menu(menubar, tearoff=False)
        seq_menu.add_command(label="Add New Sequence…", command=self._menu_add_new_sequence)
        seq_menu.add_command(label="Edit Sequence File…", command=self._menu_edit_sequence)
        menubar.add_cascade(label="Sequence", menu=seq_menu)

        # Run command on menubar
        menubar.add_command(label="Run", command=self._menu_run, accelerator="F5")

        self.config(menu=menubar)

        # Shortcuts
        self.bind_all("<Control-o>", lambda e: self._menu_open())
        self.bind_all("<Control-s>", lambda e: self._menu_save_as())
        self.bind_all("<Control-q>", lambda e: self._menu_exit())
        self.bind_all("<Control-n>", lambda e: self._menu_add_new())
        self.bind_all("<Control-e>", lambda e: self._menu_edit_selected())
        self.bind_all("<F2>",        lambda e: self._menu_edit_selected())
        self.bind_all("<F5>",        lambda e: self._menu_run())

    # ------------------ File actions ------------------
    def _menu_open(self):
        path = filedialog.askopenfilename(
            title="Open Cues XML",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            rows = self._read_xml(path)
        except Exception as e:
            messagebox.showerror("Open Failed", f"Could not open file:\n{e}")
            return

        self.clear_rows()
        self.load_rows(rows)
        self.current_file = path
        self.config_data["last_xml_path"] = path  # update in-memory config; will persist on close
        self._update_title()

        # After manual open, reset the remembered index to the first row and ensure keyboard focus.
        self._last_selected_index = 0
        self._select_row_by_index(self._last_selected_index)
        self.after(0, self.tree.focus_set)

    def _menu_save_as(self):
        initialfile = self.current_file.split("/")[-1] if self.current_file else None
        path = filedialog.asksaveasfilename(
            title="Save Cues XML",
            defaultextension=".xml",
            initialfile=initialfile or "cues.xml",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            rows = self._gather_rows()
            self._write_xml(path, rows)
            self.current_file = path
            self.config_data["last_xml_path"] = path  # update in-memory config; will persist on close
            self._update_title()
        except Exception as e:
            messagebox.showerror("Save Failed", f"Could not save file:\n{e}")

    def _menu_exit(self):
        self.on_closing()

    # ------------------ Sequence menu handlers ------------------
    def _ensure_sequences_dir(self):
        try:
            SEQUENCES_DIR.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

    def _menu_add_new_sequence(self):
        """Create a new sequence file using the SequenceEditorDialog."""
        self._ensure_sequences_dir()
        dlg = SequenceEditorDialog(self, "New Sequence", initial_file=None)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        filename, steps = dlg.result
        path = SEQUENCES_DIR / (filename if filename.endswith('.xml') else filename + '.xml')
        try:
            self._write_sequence_xml(str(path), steps)
            messagebox.showinfo("Sequence Saved", f"Sequence saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Save Failed", f"Could not save sequence:\n{e}")

    def _menu_edit_sequence(self):
        self._ensure_sequences_dir()
        path = filedialog.askopenfilename(title="Open Sequence File", initialdir=str(SEQUENCES_DIR), filetypes=[("XML files","*.xml"), ("All files","*.*")])
        if not path:
            return
        try:
            steps = self._read_sequence_xml(path)
        except Exception as e:
            messagebox.showerror("Open Failed", f"Could not open sequence:\n{e}")
            return
        initial_file = Path(path).name
        dlg = SequenceEditorDialog(self, f"Edit Sequence — {initial_file}", initial_file=initial_file, initial_steps=steps)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        filename, steps = dlg.result
        save_path = SEQUENCES_DIR / (filename if filename.endswith('.xml') else filename + '.xml')
        try:
            self._write_sequence_xml(str(save_path), steps)
            messagebox.showinfo("Sequence Saved", f"Sequence saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Save Failed", f"Could not save sequence:\n{e}")

    # ------------------ Run menu handler ------------------
    def _menu_run(self):
        """Build a <Cues> root from current grid and execute it via execute_sequence.

        Runs in a background thread so the UI stays responsive.
        """
        rows = self._gather_rows()
        if not rows:
            messagebox.showinfo("Run Sequence", "No cues to run.")
            return

        # Build XML root
        cues_el = ET.Element("Cues")
        for r in rows:
            cue_el = ET.SubElement(cues_el, "Cue")
            cue_el.set("name", str(r.get("Cue", "")))
            order = r.get("Order")
            cue_el.set("order", str(order if order is not None else ""))
            cue_el.set("sequence", str(r.get("Sequence", "")))
            cue_el.set("delay", str(r.get("Delay", "")))
            for i in range(1, 9):
                s = ET.SubElement(cue_el, f"Switch{i}")
                s.text = "true" if bool(r.get(f"Switch{i}", False)) else "false"

        # Run in background thread
        def worker():
            self._set_busy(True)
            try:
                execute_sequence(
                    cues_el,
                    base_url=f"http://{self.device_host}",
                    username=self.device_user,
                    password=self.device_pass,
                    timeout=DLI_TIMEOUT,
                )
            except Exception as e:
                msg = str(e)
                # self._safe_after(0, lambda: messagebox.showerror("Run Failed", f"Sequence execution failed:\n{msg}"))
                self._safe_after(0, lambda m=msg: self.dli_status.set(f"DLI (HTTP) error: {m}"))
            else:
                # self._safe_after(0, lambda: messagebox.showinfo("Run Complete", "Sequence executed successfully."))
                self._safe_after(0, lambda: self.dli_status.set("DLI (HTTP): sequence run complete"))
            finally:
                self._safe_after(0, lambda: self._set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    # ------------------ Sequence XML I/O ------------------
    def _read_sequence_xml(self, path: str) -> list:
        tree = ET.parse(path)
        root = tree.getroot()
        if root.tag != "Sequence":
            # allow root 'Sequence' or read step elements regardless
            pass
        steps = []
        for step_el in root.findall("Step"):
            sw_raw = step_el.get("switch") or step_el.get("Switch") or step_el.get("channel")
            try:
                sw = int(sw_raw) if sw_raw is not None else 0
            except Exception:
                sw = 0
            pos_raw = step_el.get("position") or step_el.get("Position") or step_el.get("state")
            pos = str(pos_raw).strip().lower() in {"1", "true", "t", "yes", "y", "on"}
            delay_raw = step_el.get("delay") or step_el.get("Delay") or step_el.text
            try:
                delay = int(delay_raw) if delay_raw is not None else 0
            except Exception:
                delay = 0
            steps.append({"switch": sw, "position": pos, "delay": delay})
        return steps

    def _write_sequence_xml(self, path: str, steps: list):
        root = ET.Element("Sequence")
        for s in steps:
            el = ET.SubElement(root, "Step")
            el.set("switch", str(s.get("switch", 0)))
            el.set("position", "true" if bool(s.get("position", False)) else "false")
            el.set("delay", str(int(s.get("delay", 0))))
        ET.ElementTree(root).write(path, encoding="utf-8", xml_declaration=True)

    def _update_title(self):
        suffix = f" — {self.current_file}" if self.current_file else ""
        self.title(f"Cue Switchboard{suffix}")

    # ------------------ Cue actions ------------------
    def _menu_add_new(self):
        next_order = self._next_order()
        dlg = CueEditorDialog(self, "Add New Cue", initial=None, suggested_order=next_order)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        if dlg.result.get("Order") in ("", None):
            dlg.result["Order"] = next_order if next_order is not None else ""
        iid = self.add_row(dlg.result)
        if iid:
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)
            self.tree.focus_set()
            self._emit_for_iid(iid)

    def _menu_edit_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Edit Cue", "Please select a cue to edit.")
            return
        iid = sel[0]
        current = self.row_data_by_iid.get(iid, {})
        dlg = CueEditorDialog(self, "Edit Cue", initial=current)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        self.row_data_by_iid[iid] = {k: dlg.result.get(k) for k in COLUMNS}
        self._refresh_item_values(iid)

    # ------------------ Inline editing / toggling ------------------
    def _on_single_click_cell(self, event):
        # Let Treeview process selection first
        self.after(1, lambda: self._toggle_if_switch_cell(event))

    def _toggle_if_switch_cell(self, event):
        item = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item or not column_id:
            return
        col_index = int(column_id.replace('#', '')) - 1
        if col_index < 0:
            return
        col_name = COLUMNS[col_index]
        if col_name in SWITCH_COLUMNS:
            cue = self.row_data_by_iid.get(item)
            if cue is None:
                return
            new_val = not bool(cue.get(col_name, False))
            cue[col_name] = new_val
            vals = list(self.tree.item(item, "values"))
            vals[col_index] = "True" if new_val else "False"
            self.tree.item(item, values=vals)

    def _on_double_click_cell(self, event):
        item = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item or not column_id:
            return
        col_index = int(column_id.replace('#', '')) - 1
        if col_index < 0:
            return
        col_name = COLUMNS[col_index]
        if col_name in SWITCH_COLUMNS:
            return  # switches toggle on single click

        bbox = self.tree.bbox(item, column_id)
        if not bbox:
            return
        x, y, width, height = bbox

        current_vals = self.tree.item(item, "values")
        current_text = current_vals[col_index] if col_index < len(current_vals) else ""

        if self._inline_editor and self._inline_editor.winfo_exists():
            self._inline_editor.destroy()
        self._inline_editor = tk.Entry(self.tree)
        self._inline_editor.insert(0, current_text)
        self._inline_editor.select_range(0, tk.END)
        self._inline_editor.focus()
        self._inline_editor.place(x=x, y=y, width=width, height=height)

        def commit():
            new_text = self._inline_editor.get().strip()
            self._inline_editor.destroy()
            new_vals = list(current_vals)
            new_vals[col_index] = new_text
            self.tree.item(item, values=new_vals)

            cue = self.row_data_by_iid.get(item, {})
            if col_name in ("Order", "Delay"):
                try:
                    cue[col_name] = int(new_text) if new_text != "" else ""
                except ValueError:
                    cue[col_name] = new_text
            else:
                cue[col_name] = new_text

        self._inline_editor.bind("<Return>", lambda e: commit())
        self._inline_editor.bind("<Escape>", lambda e: self._inline_editor.destroy())
        self._inline_editor.bind("<FocusOut>", lambda e: commit())

    # ------------------ Helpers ------------------
    def _next_order(self) -> Optional[int]:
        orders = []
        for r in self._gather_rows():
            v = r.get("Order")
            if isinstance(v, int):
                orders.append(v)
            else:
                try:
                    orders.append(int(v))
                except Exception:
                    pass
        return (max(orders) + 1) if orders else 1

    def add_row(self, row: Dict) -> Optional[str]:
        clean = {k: row.get(k) for k in COLUMNS}
        display_values = [self._to_cell_text(clean[c]) for c in COLUMNS]
        iid = self.tree.insert("", "end", values=display_values)
        self.row_data_by_iid[iid] = clean
        return iid

    def _refresh_item_values(self, iid: str):
        cue = self.row_data_by_iid.get(iid, {})
        vals = [self._to_cell_text(cue.get(c)) for c in COLUMNS]
        self.tree.item(iid, values=vals)

    # ------------------ XML I/O ------------------
    def _read_xml(self, path: str) -> List[Dict]:
        tree = ET.parse(path)
        root = tree.getroot()
        if root.tag != "Cues":
            raise ValueError("Root element must be <Cues>")

        rows: List[Dict] = []
        for cue_el in root.findall("Cue"):
            name = cue_el.get("name", "")
            order_raw = cue_el.get("order", "")
            order = self._safe_int(order_raw)
            sequence = cue_el.get("sequence", "")
            delay_raw = cue_el.get("delay", "")
            delay = self._safe_int(delay_raw)

            cue = {"Cue": name, "Order": order, "Sequence": sequence, "Delay": delay}
            for i in range(1, 9):
                tag = f"Switch{i}"
                val_el = cue_el.find(tag)
                cue[tag] = self._parse_bool(val_el.text) if val_el is not None and val_el.text is not None else False
            rows.append(cue)

        rows.sort(key=lambda r: (r["Order"] if isinstance(r["Order"], int) else float("inf")))
        return rows

    def _write_xml(self, path: str, rows: List[Dict]):
        cues_el = ET.Element("Cues")
        for r in rows:
            cue_el = ET.SubElement(cues_el, "Cue")
            cue_el.set("name", str(r.get("Cue", "")))
            order = r.get("Order")
            cue_el.set("order", str(order if order is not None else ""))
            cue_el.set("sequence", str(r.get("Sequence", "")))
            cue_el.set("delay", str(r.get("Delay", "")))

            for i in range(1, 9):
                s = ET.SubElement(cue_el, f"Switch{i}")
                s.text = "true" if bool(r.get(f"Switch{i}", False)) else "false"

        ET.ElementTree(cues_el).write(path, encoding="utf-8", xml_declaration=True)

    @staticmethod
    def _parse_bool(text: str) -> bool:
        return str(text).strip().lower() in {"1", "true", "t", "yes", "y", "on"}

    @staticmethod
    def _safe_int(v):
        try:
            return int(v)
        except (ValueError, TypeError):
            return v


    

    # ------------------ Table ops ------------------
    def clear_rows(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.row_data_by_iid.clear()

    def load_rows(self, rows: List[Dict]):
        for row in rows:
            self.add_row(row)

    def _gather_rows(self) -> List[Dict]:
        return [self.row_data_by_iid[iid] for iid in self.tree.get_children()]

    def _to_cell_text(self, val):
        if isinstance(val, bool):
            return "True" if val else "False"
        return "" if val is None else str(val)

    # ------------------ Selection handling (debounced) ------------------
    def _on_tree_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]

        # Track and remember selected index for persistence
        try:
            children = list(self.tree.get_children())
            self._last_selected_index = children.index(iid)
        except Exception:
            self._last_selected_index = 0

        # 1) Immediately notify domain callback so any UI/state reacts right away
        self._emit_for_iid(iid)

        # 2) Debounce & defer the HTTP send so the highlight renders first
        self._debounced_send_for_iid(iid, delay_ms=120)

    def _select_row_by_index(self, index: Optional[int]):
        """Attempt to select the row at 'index'; fall back to first row. Also ensures keyboard focus."""
        children = list(self.tree.get_children())
        if not children:
            return
        try:
            idx = int(index) if index is not None else 0
            if idx < 0 or idx >= len(children):
                idx = 0
        except Exception:
            idx = 0

        iid = children[idx]
        self.tree.selection_set(iid)
        self.tree.focus(iid)       # item focus (Treeview supports .focus)
        self.tree.see(iid)
        self.tree.focus_set()      # keyboard focus to the Treeview
        self._emit_for_iid(iid)    # HTTP send still debounced via <<TreeviewSelect>>

    def _emit_for_iid(self, iid):
        cue = self.row_data_by_iid.get(iid)
        if cue is not None and self.cue_change_callback:
            self.cue_change_callback(cue)


def execute_sequence_from_file(xml_path, base_url="http://192.168.0.100", username="admin", password="1234", timeout=5):
    url = f"{base_url.rstrip('/')}/restapi/relay/set_outlet_transient_states/"
    auth = requests.auth.HTTPDigestAuth(username, password)

    headers = {
        "X-CSRF": "x",
        "Content-Type": "application/json",
    }

    tree = ET.parse(xml_path)
    root = tree.getroot()

    if root.tag != "Cues":
        raise ValueError(f"Expected root <Cues>, got <{root.tag}>")

    cues = []
    for cue in root.findall("Cue"):
        order_raw = cue.get("order")
        if order_raw is None:
            raise ValueError("Cue missing required 'order' attribute")
        try:
            order = int(order_raw)
        except ValueError as e:
            raise ValueError(f"Invalid Cue order='{order_raw}' (must be integer)") from e

        cues.append((order, cue))

    cues.sort(key=lambda t: t[0])

    for _, cue in cues:
        steps = []

        for child in list(cue):
            tag = child.tag  # e.g., "Switch1"
            if not tag.startswith("Switch"):
                continue

            num_raw = tag[len("Switch"):]  # "1"
            if not num_raw.isdigit():
                continue

            switch_1based = int(num_raw)
            index_0based = switch_1based - 1

            text = (child.text or "").strip().lower()
            if text not in ("true", "false"):
                raise ValueError(f"Invalid value for <{tag}>: '{child.text}' (must be true/false)")

            steps.append([index_0based, text == "true"])

        if not steps:
            raise ValueError("Cue produced no Switch settings (no <SwitchN> elements found)")

        payload = [steps]

        resp = requests.post(
            url,
            auth=auth,
            headers=headers,
            json=payload,
            timeout=timeout,
        )
        resp.raise_for_status()

        # Optional per-cue delay (milliseconds)
        delay_raw = cue.get("delay")
        delay_ms = 0
        
        if delay_raw is not None:
            delay_raw = delay_raw.strip()
            if delay_raw != "":
                try:
                    delay_ms = int(delay_raw)
                except ValueError as e:
                    raise ValueError(f"Invalid Cue delay='{delay_raw}' (must be integer milliseconds)") from e
        
        if delay_ms > 0:
            time.sleep(delay_ms / 1000.0)


def execute_sequence(
    root,
    base_url="http://192.168.0.100",
    username="admin",
    password="1234",
    timeout=5,
):
    """
    Execute a relay sequence defined by a <Cues> XML root Element.
    """
    url = f"{base_url.rstrip('/')}/restapi/relay/set_outlet_transient_states/"
    auth = requests.auth.HTTPDigestAuth(username, password)

    headers = {
        "X-CSRF": "x",
        "Content-Type": "application/json",
    }

    # Collect and sort cues by order
    cues = []
    for cue in root.findall("Cue"):
        order_raw = cue.get("order")
        if order_raw is None:
            raise ValueError("Cue missing required 'order' attribute")

        try:
            order = int(order_raw)
        except ValueError as e:
            raise ValueError(
                f"Invalid Cue order='{order_raw}' (must be integer)"
            ) from e

        cues.append((order, cue))

    cues.sort(key=lambda t: t[0])

    # Execute each cue
    for _, cue in cues:
        steps = []

        for child in list(cue):
            tag = child.tag
            if not tag.startswith("Switch"):
                continue

            num_raw = tag[len("Switch"):]
            if not num_raw.isdigit():
                continue

            switch_1based = int(num_raw)
            index_0based = switch_1based - 1

            text = (child.text or "").strip().lower()
            if text not in ("true", "false"):
                raise ValueError(
                    f"Invalid value for <{tag}>: '{child.text}' (must be true/false)"
                )

            steps.append([index_0based, text == "true"])

        if not steps:
            raise ValueError(
                "Cue produced no Switch settings (no <SwitchN> elements found)"
            )

        payload = [steps]

        resp = requests.post(
            url,
            auth=auth,
            headers=headers,
            json=payload,
            timeout=timeout,
        )
        resp.raise_for_status()

        # Optional per-cue delay (milliseconds)
        delay_raw = cue.get("delay")
        delay_ms = 0

        if delay_raw is not None:
            delay_raw = delay_raw.strip()
            if delay_raw != "":
                try:
                    delay_ms = int(delay_raw)
                except ValueError as e:
                    raise ValueError(
                        f"Invalid Cue delay='{delay_raw}' (must be integer milliseconds)"
                    ) from e

        if delay_ms > 0:
            time.sleep(delay_ms / 1000.0)


if __name__ == "__main__":
    app = CueTableApp(cue_change_callback=on_cue_change)
    app.mainloop()
