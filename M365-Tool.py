#!/usr/bin/env python3
"""
Microsoft 365 Admin Tool v3.0
Erstellt für Kaulich IT Systems GmbH

Features:
- Postfachberechtigungen (FullAccess, SendAs, AutoMapping)
- Microsoft 365 Gruppen / Teams-Gruppen (Mitglieder/Besitzer)
- Verteilerlisten (Distribution Groups)
- Sicherheitsgruppen (Mail-Enabled Security Groups)
- Automatische Modulprüfung & Installation
- Dark Theme UI

Voraussetzungen:
- Windows mit PowerShell
- ExchangeOnlineManagement Modul (wird automatisch installiert)
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import json
import queue
import time
from datetime import datetime


# ============================================================
# FARBSCHEMA - Dark Theme
# ============================================================
COLORS = {
    'bg':               '#1e1e2e',
    'bg_panel':         '#252536',
    'bg_input':         '#2e2e42',
    'bg_header':        '#16161e',
    'accent':           '#7aa2f7',
    'accent_hover':     '#89b4fa',
    'success':          '#9ece6a',
    'warning':          '#e0af68',
    'error':            '#f7768e',
    'text':             '#c0caf5',
    'text_dim':         '#565f89',
    'text_muted':       '#414868',
    'border':           '#3b3b54',
    'button_danger':    '#f7768e',
    'button_success':   '#9ece6a',
    'purple':           '#bb9af7',
    'cyan':             '#7dcfff',
    'orange':           '#ff9e64',
    'green':            '#73daca',
}

MODULES = {
    'postfach':      '📧 Postfachberechtigungen',
    'teams':         '👥 Microsoft 365 / Teams Gruppen',
    'verteiler':     '📨 Verteilerlisten',
    'security':      '🔒 Sicherheitsgruppen',
}


class PowerShellSession:
    """Persistente PowerShell-Session für Exchange Online"""

    def __init__(self):
        self.process = None
        self.output_queue = queue.Queue()

    def start(self):
        if self.process is not None:
            return
        self.process = subprocess.Popen(
            ["powershell", "-NoLogo", "-NoExit", "-Command", "-"],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, text=True,
            encoding='utf-8', errors='replace', bufsize=1
        )
        self.reader_thread = threading.Thread(target=self._read_output, daemon=True)
        self.reader_thread.start()
        self._send_command('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8')

    def _read_output(self):
        while self.process and self.process.poll() is None:
            try:
                line = self.process.stdout.readline()
                if line:
                    self.output_queue.put(line)
            except:
                break

    def _send_command(self, command):
        if self.process and self.process.poll() is None:
            self.process.stdin.write(command + "\n")
            self.process.stdin.flush()

    def execute(self, command, timeout=120):
        if not self.process or self.process.poll() is not None:
            return False, "", "PowerShell-Session nicht aktiv"

        while not self.output_queue.empty():
            try: self.output_queue.get_nowait()
            except: break

        end_marker = f"###END_{int(time.time() * 1000)}###"
        error_marker = f"###ERROR_{int(time.time() * 1000)}###"

        wrapped = f"""
try {{
    {command}
}} catch {{
    Write-Output "{error_marker}$($_.Exception.Message)"
}}
Write-Output "{end_marker}"
"""
        self._send_command(wrapped)

        output_lines = []
        error_lines = []
        start_time = time.time()

        while True:
            if time.time() - start_time > timeout:
                return False, "", "Timeout"
            try:
                line = self.output_queue.get(timeout=0.5)
                stripped = line.rstrip()
                if end_marker in stripped:
                    break
                elif error_marker in stripped:
                    error_lines.append(stripped.split(error_marker, 1)[1] if error_marker in stripped else stripped)
                else:
                    output_lines.append(stripped)
            except queue.Empty:
                continue

        output = "\n".join(output_lines)
        errors = "\n".join(error_lines)
        return (not bool(error_lines)), output, errors

    def stop(self):
        if self.process:
            try:
                self._send_command("exit")
                self.process.terminate()
            except:
                pass
            self.process = None


class StyledButton(tk.Canvas):
    """Flat Button mit Hover"""

    def __init__(self, parent, text, command=None, bg=COLORS['accent'],
                 fg='#ffffff', width=160, height=34, font_size=10, **kwargs):
        super().__init__(parent, width=width, height=height,
                         bg=parent.cget('bg'), highlightthickness=0, **kwargs)
        self.command = command
        self.bg_color = bg
        self.fg_color = fg
        self.hover_color = self._adjust(bg, 25)
        self.disabled_color = COLORS['text_muted']
        self.text = text
        self._width = width
        self._height = height
        self.font_size = font_size
        self.enabled = True
        self._draw(self.bg_color)
        self.bind('<Enter>', lambda e: self._draw(self.hover_color) if self.enabled else None)
        self.bind('<Leave>', lambda e: self._draw(self.bg_color if self.enabled else self.disabled_color))
        self.bind('<Button-1>', lambda e: self.command() if self.enabled and self.command else None)

    def _adjust(self, color, amount):
        try:
            r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
            return f'#{min(255,r+amount):02x}{min(255,g+amount):02x}{min(255,b+amount):02x}'
        except:
            return color

    def _draw(self, color):
        self.delete('all')
        r = 6
        w, h = self._width, self._height
        self.create_arc(0, 0, r*2, r*2, start=90, extent=90, fill=color, outline=color)
        self.create_arc(w-r*2, 0, w, r*2, start=0, extent=90, fill=color, outline=color)
        self.create_arc(0, h-r*2, r*2, h, start=180, extent=90, fill=color, outline=color)
        self.create_arc(w-r*2, h-r*2, w, h, start=270, extent=90, fill=color, outline=color)
        self.create_rectangle(r, 0, w-r, h, fill=color, outline=color)
        self.create_rectangle(0, r, w, h-r, fill=color, outline=color)
        tc = self.fg_color if self.enabled else COLORS['text_dim']
        self.create_text(w/2, h/2, text=self.text, fill=tc,
                         font=('Segoe UI', self.font_size, 'bold'))

    def configure(self, **kw):
        if 'state' in kw:
            self.enabled = kw['state'] != tk.DISABLED
            self._draw(self.bg_color if self.enabled else self.disabled_color)
        if 'text' in kw:
            self.text = kw['text']
            self._draw(self.bg_color if self.enabled else self.disabled_color)
        if 'bg' in kw:
            self.bg_color = kw['bg']
            self.hover_color = self._adjust(kw['bg'], 25)
            self._draw(self.bg_color)


class M365AdminTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Microsoft 365 Admin Tool v3.0 — Kaulich IT Systems GmbH")
        self.root.geometry("750x820")
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS['bg'])

        self.ps = PowerShellSession()
        self.ps.start()

        self.connected = False
        self.all_mailboxes = []
        self.all_groups = {
            'teams': [],
            'verteiler': [],
            'security': [],
        }
        self.current_module = tk.StringVar(value='postfach')

        self.build_ui()
        self.log("🚀 M365 Admin Tool gestartet", COLORS['success'])
        self.root.after(500, self.check_module)

    # ============================================================
    # UI BUILDING
    # ============================================================

    def build_ui(self):
        # Header
        header = tk.Frame(self.root, bg=COLORS['bg_header'], height=50)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(header, text="⚡ Microsoft 365 Admin Tool",
                 font=('Segoe UI', 16, 'bold'), fg=COLORS['accent'],
                 bg=COLORS['bg_header']).pack(side=tk.LEFT, padx=15)

        tk.Label(header, text="v3.0 — Kaulich IT Systems GmbH",
                 font=('Segoe UI', 9), fg=COLORS['text_dim'],
                 bg=COLORS['bg_header']).pack(side=tk.RIGHT, padx=15)

        main = tk.Frame(self.root, bg=COLORS['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # === VERBINDUNG ===
        conn_frame = self._section(main, "🔐 Verbindung")
        conn_inner = tk.Frame(conn_frame, bg=COLORS['bg_panel'])
        conn_inner.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(conn_inner, text="Admin-UPN:", font=('Segoe UI', 10),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        self.admin_entry = tk.Entry(conn_inner, width=35, font=('Segoe UI', 10),
                                    bg=COLORS['bg_input'], fg=COLORS['text'],
                                    insertbackground=COLORS['text'], relief=tk.FLAT,
                                    highlightthickness=1, highlightbackground=COLORS['border'],
                                    highlightcolor=COLORS['accent'])
        self.admin_entry.grid(row=0, column=1, padx=(0, 10), ipady=5)

        self.connect_btn = StyledButton(conn_inner, "🔌 Verbinden",
                                        command=self.connect, width=130)
        self.connect_btn.grid(row=0, column=2)

        self.module_label = tk.Label(conn_frame, text="⏳ Prüfe Module...",
                                     font=('Segoe UI', 9), fg=COLORS['warning'],
                                     bg=COLORS['bg_panel'])
        self.module_label.pack(fill=tk.X, padx=10, pady=(0, 5))

        hint = tk.Frame(conn_frame, bg=COLORS['bg_panel'])
        hint.pack(fill=tk.X, padx=10, pady=(0, 8))
        tk.Label(hint, text="⚠️ Session wird beim Beenden automatisch getrennt",
                 font=('Segoe UI', 9), fg=COLORS['text_dim'],
                 bg=COLORS['bg_panel']).pack(anchor=tk.W)

        # === MODUL-AUSWAHL ===
        mod_frame = self._section(main, "🔧 Funktion")
        mod_inner = tk.Frame(mod_frame, bg=COLORS['bg_panel'])
        mod_inner.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(mod_inner, text="Modul:", font=('Segoe UI', 10, 'bold'),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0, 10))

        self.module_combo = ttk.Combobox(mod_inner, width=45, state="readonly",
                                         font=('Segoe UI', 10),
                                         values=list(MODULES.values()))
        self.module_combo.pack(side=tk.LEFT)
        self.module_combo.current(0)
        self.module_combo.bind('<<ComboboxSelected>>', self.on_module_change)

        # === DYNAMISCHER BEREICH ===
        self.dynamic_frame = tk.Frame(main, bg=COLORS['bg'])
        self.dynamic_frame.pack(fill=tk.BOTH, expand=True)

        # Alle Module erstellen (übereinander, nur eins sichtbar)
        self.module_frames = {}
        self._build_postfach_module()
        self._build_group_module('teams', "Teams-Gruppe")
        self._build_group_module('verteiler', "Verteilerliste")
        self._build_group_module('security', "Sicherheitsgruppe")

        # === PROTOKOLL ===
        log_frame = self._section(main, "📋 Protokoll")
        log_inner = tk.Frame(log_frame, bg=COLORS['bg_panel'])
        log_inner.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.log_text = tk.Text(log_inner, height=7, font=('Consolas', 9),
                                bg=COLORS['bg_input'], fg=COLORS['text'],
                                relief=tk.FLAT, wrap=tk.WORD,
                                highlightthickness=1,
                                highlightbackground=COLORS['border'])

        scrollbar = tk.Scrollbar(log_inner, orient=tk.VERTICAL,
                                 command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for tag, color in [('success', COLORS['success']), ('error', COLORS['error']),
                           ('warning', COLORS['warning']), ('info', COLORS['text']),
                           ('accent', COLORS['accent']), ('dim', COLORS['text_dim'])]:
            self.log_text.tag_configure(tag, foreground=color)

        # Combobox Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TCombobox',
                        fieldbackground=COLORS['bg_input'],
                        background=COLORS['bg_panel'],
                        foreground=COLORS['text'],
                        arrowcolor=COLORS['text'],
                        bordercolor=COLORS['border'],
                        lightcolor=COLORS['border'],
                        darkcolor=COLORS['border'])

        # Erstes Modul anzeigen
        self.show_module('postfach')

    def _section(self, parent, title):
        container = tk.Frame(parent, bg=COLORS['bg'])
        container.pack(fill=tk.X, pady=(0, 8))
        tk.Label(container, text=title, font=('Segoe UI', 11, 'bold'),
                 fg=COLORS['accent'], bg=COLORS['bg']).pack(anchor=tk.W, pady=(0, 4))
        frame = tk.Frame(container, bg=COLORS['bg_panel'],
                         highlightbackground=COLORS['border'], highlightthickness=1)
        frame.pack(fill=tk.X)
        return frame

    def _make_combo_row(self, parent, label_text, hint_text=None):
        """Erstellt ein Label + Combobox Paar"""
        row = tk.Frame(parent, bg=COLORS['bg_panel'])
        row.pack(fill=tk.X, padx=10, pady=(5, 0))
        tk.Label(row, text=label_text, font=('Segoe UI', 10),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0, 10))
        combo = ttk.Combobox(row, width=50, state="disabled", font=('Segoe UI', 9))
        combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        combo.set("-- Erst verbinden --")
        if hint_text:
            hint_row = tk.Frame(parent, bg=COLORS['bg_panel'])
            hint_row.pack(fill=tk.X, padx=10, pady=(1, 0))
            tk.Label(hint_row, text=hint_text, font=('Segoe UI', 8),
                     fg=COLORS['text_dim'], bg=COLORS['bg_panel']).pack(anchor=tk.E)
        return combo

    def _make_search(self, parent, callback):
        """Erstellt ein Suchfeld"""
        row = tk.Frame(parent, bg=COLORS['bg_panel'])
        row.pack(fill=tk.X, padx=10, pady=(8, 8))
        tk.Label(row, text="🔍", font=('Segoe UI', 10),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0, 5))
        entry = tk.Entry(row, width=55, font=('Segoe UI', 10),
                         bg=COLORS['bg_input'], fg=COLORS['text'],
                         insertbackground=COLORS['text'], relief=tk.FLAT,
                         highlightthickness=1, highlightbackground=COLORS['border'],
                         highlightcolor=COLORS['accent'])
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        entry.bind('<KeyRelease>', callback)
        return entry

    # ============================================================
    # MODUL: POSTFACH
    # ============================================================

    def _build_postfach_module(self):
        frame = tk.Frame(self.dynamic_frame, bg=COLORS['bg'])
        self.module_frames['postfach'] = frame

        sec = self._section_in(frame, "📧 Postfachberechtigungen verwalten")

        self.mb_target_combo = self._make_combo_row(sec, "Ziel-Postfach:",
                                                     "(Postfach, auf das zugegriffen werden soll)")
        self.mb_user_combo = self._make_combo_row(sec, "Benutzer:      ",
                                                    "(Benutzer, der Zugriff erhalten/verlieren soll)")
        self.mb_search = self._make_search(sec, self.filter_mailboxes)

        # Filter
        filter_row = tk.Frame(sec, bg=COLORS['bg_panel'])
        filter_row.pack(fill=tk.X, padx=10, pady=(0, 5))
        tk.Label(filter_row, text="Typ:", font=('Segoe UI', 9, 'bold'),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0, 8))
        self.mb_filter = tk.StringVar(value="all")
        for text, val in [("Alle", "all"), ("👤 Benutzer", "user"), ("👥 Shared", "shared")]:
            tk.Radiobutton(filter_row, text=text, variable=self.mb_filter, value=val,
                           font=('Segoe UI', 9), fg=COLORS['text'], bg=COLORS['bg_panel'],
                           selectcolor=COLORS['bg_input'], activebackground=COLORS['bg_panel'],
                           command=self.apply_mb_filter).pack(side=tk.LEFT, padx=(0, 12))

        # Berechtigungen
        perm_row = tk.Frame(sec, bg=COLORS['bg_panel'])
        perm_row.pack(fill=tk.X, padx=10, pady=(5, 5))

        self.fullaccess_var = tk.BooleanVar(value=True)
        self.automapping_var = tk.BooleanVar(value=True)
        self.sendas_var = tk.BooleanVar(value=True)

        tk.Checkbutton(perm_row, text="📂 Vollzugriff", variable=self.fullaccess_var,
                        font=('Segoe UI', 10), fg=COLORS['text'], bg=COLORS['bg_panel'],
                        selectcolor=COLORS['bg_input'], activebackground=COLORS['bg_panel'],
                        command=self._toggle_automap).pack(side=tk.LEFT)

        self.automap_cb = tk.Checkbutton(perm_row, text="🔗 AutoMapping", variable=self.automapping_var,
                                          font=('Segoe UI', 10), fg=COLORS['text'], bg=COLORS['bg_panel'],
                                          selectcolor=COLORS['bg_input'], activebackground=COLORS['bg_panel'])
        self.automap_cb.pack(side=tk.LEFT, padx=(20, 0))

        tk.Checkbutton(perm_row, text="✉️ Senden als", variable=self.sendas_var,
                        font=('Segoe UI', 10), fg=COLORS['text'], bg=COLORS['bg_panel'],
                        selectcolor=COLORS['bg_input'], activebackground=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(20, 0))

        # Buttons
        btn_row = tk.Frame(sec, bg=COLORS['bg_panel'])
        btn_row.pack(fill=tk.X, padx=10, pady=(5, 10))

        StyledButton(btn_row, "✅ Berechtigung hinzufügen", command=self.add_mb_permission,
                     bg=COLORS['button_success'], fg='#1e1e2e', width=200).pack(side=tk.LEFT, padx=(0, 10))
        StyledButton(btn_row, "❌ Berechtigung entfernen", command=self.remove_mb_permission,
                     bg=COLORS['button_danger'], width=200).pack(side=tk.LEFT, padx=(0, 10))
        self.mb_disconnect_btn = StyledButton(btn_row, "🔌 Trennen", command=self.disconnect,
                                               bg=COLORS['bg_input'], width=100)
        self.mb_disconnect_btn.pack(side=tk.RIGHT)

    # ============================================================
    # MODUL: GRUPPEN (Teams / Verteiler / Security)
    # ============================================================

    def _build_group_module(self, key, label):
        frame = tk.Frame(self.dynamic_frame, bg=COLORS['bg'])
        self.module_frames[key] = frame

        sec = self._section_in(frame, f"{'👥' if key=='teams' else '📨' if key=='verteiler' else '🔒'} {label} verwalten")

        group_combo = self._make_combo_row(sec, f"{label}:", f"(Die Gruppe, zu der hinzugefügt/entfernt werden soll)")
        user_combo = self._make_combo_row(sec, "Benutzer:     ", "(Der Benutzer)")
        search = self._make_search(sec, lambda e, k=key: self.filter_groups(k))

        # Rolle (nur bei Teams)
        role_var = tk.StringVar(value="Member")
        if key == 'teams':
            role_row = tk.Frame(sec, bg=COLORS['bg_panel'])
            role_row.pack(fill=tk.X, padx=10, pady=(0, 5))
            tk.Label(role_row, text="Rolle:", font=('Segoe UI', 9, 'bold'),
                     fg=COLORS['text'], bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0, 8))
            for text, val in [("👤 Mitglied", "Member"), ("👑 Besitzer", "Owner")]:
                tk.Radiobutton(role_row, text=text, variable=role_var, value=val,
                               font=('Segoe UI', 9), fg=COLORS['text'], bg=COLORS['bg_panel'],
                               selectcolor=COLORS['bg_input'], activebackground=COLORS['bg_panel']
                               ).pack(side=tk.LEFT, padx=(0, 12))

        # Mitglieder anzeigen Button + Liste
        members_frame = tk.Frame(sec, bg=COLORS['bg_panel'])
        members_frame.pack(fill=tk.X, padx=10, pady=(0, 5))

        members_text = tk.Text(members_frame, height=4, font=('Consolas', 9),
                               bg=COLORS['bg_input'], fg=COLORS['text'],
                               relief=tk.FLAT, wrap=tk.WORD, state=tk.DISABLED,
                               highlightthickness=1, highlightbackground=COLORS['border'])
        members_text.pack(fill=tk.X, pady=(5, 0))

        show_members_btn = StyledButton(members_frame, "📋 Mitglieder anzeigen",
                                         command=lambda k=key: self.show_members(k),
                                         bg=COLORS['bg_input'], width=170, height=28, font_size=9)
        show_members_btn.pack(anchor=tk.W, pady=(5, 0))

        # Buttons
        btn_row = tk.Frame(sec, bg=COLORS['bg_panel'])
        btn_row.pack(fill=tk.X, padx=10, pady=(5, 10))

        StyledButton(btn_row, "✅ Hinzufügen", command=lambda k=key: self.add_group_member(k),
                     bg=COLORS['button_success'], fg='#1e1e2e', width=160).pack(side=tk.LEFT, padx=(0, 10))
        StyledButton(btn_row, "❌ Entfernen", command=lambda k=key: self.remove_group_member(k),
                     bg=COLORS['button_danger'], width=160).pack(side=tk.LEFT, padx=(0, 10))

        disconnect_btn = StyledButton(btn_row, "🔌 Trennen", command=self.disconnect,
                                       bg=COLORS['bg_input'], width=100)
        disconnect_btn.pack(side=tk.RIGHT)

        # Widgets speichern
        setattr(self, f'{key}_group_combo', group_combo)
        setattr(self, f'{key}_user_combo', user_combo)
        setattr(self, f'{key}_search', search)
        setattr(self, f'{key}_role_var', role_var)
        setattr(self, f'{key}_members_text', members_text)

    def _section_in(self, parent, title):
        """Section innerhalb eines Modul-Frames"""
        container = tk.Frame(parent, bg=COLORS['bg'])
        container.pack(fill=tk.X, pady=(0, 5))
        tk.Label(container, text=title, font=('Segoe UI', 11, 'bold'),
                 fg=COLORS['purple'], bg=COLORS['bg']).pack(anchor=tk.W, pady=(0, 4))
        frame = tk.Frame(container, bg=COLORS['bg_panel'],
                         highlightbackground=COLORS['border'], highlightthickness=1)
        frame.pack(fill=tk.X)
        return frame

    # ============================================================
    # MODULE SWITCHING
    # ============================================================

    def on_module_change(self, event=None):
        selected_text = self.module_combo.get()
        for key, label in MODULES.items():
            if label == selected_text:
                self.show_module(key)
                break

    def show_module(self, key):
        for k, f in self.module_frames.items():
            f.pack_forget()
        self.module_frames[key].pack(fill=tk.BOTH, expand=True, before=self._get_log_parent())
        self.current_module.set(key)

    def _get_log_parent(self):
        """Findet den Protokoll-Frame"""
        for child in self.root.winfo_children():
            for sub in child.winfo_children():
                if isinstance(sub, tk.Frame):
                    for w in sub.winfo_children():
                        if isinstance(w, tk.Frame) and hasattr(self, 'log_text'):
                            try:
                                if self.log_text.winfo_parent() == str(w.winfo_children()[0]) if w.winfo_children() else False:
                                    return sub
                            except:
                                pass
        return None

    # ============================================================
    # VERBINDUNG & MODULE
    # ============================================================

    def check_module(self):
        self.connect_btn.configure(state=tk.DISABLED)
        self.log("🔍 Prüfe ExchangeOnlineManagement...", COLORS['warning'])

        def do_check():
            cmd = '$m = Get-Module -ListAvailable -Name ExchangeOnlineManagement; if ($m) { Write-Output "INSTALLED:$($m.Version)" } else { Write-Output "NOT_INSTALLED" }'
            success, stdout, _ = self.ps.execute(cmd, timeout=30)
            if "INSTALLED:" in stdout:
                ver = stdout.split("INSTALLED:")[1].strip().split("\n")[0]
                self.root.after(0, lambda: self._module_ok(ver))
            else:
                self.root.after(0, self._module_missing)

        threading.Thread(target=do_check, daemon=True).start()

    def _module_ok(self, ver):
        self.module_label.configure(text=f"✅ ExchangeOnlineManagement v{ver}", fg=COLORS['success'])
        self.connect_btn.configure(state=tk.NORMAL)
        self.log(f"✅ Modul v{ver} gefunden — bereit", COLORS['success'])

    def _module_missing(self):
        self.module_label.configure(text="❌ ExchangeOnlineManagement fehlt!", fg=COLORS['error'])
        self.log("❌ Modul fehlt!", COLORS['error'])

        if messagebox.askyesno("Modul fehlt",
                               "ExchangeOnlineManagement ist nicht installiert.\n\nJetzt installieren?"):
            self._install_module()
        else:
            self.log("💡 Manuell: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser", COLORS['accent'])

    def _install_module(self):
        self.log("📦 Installiere Modul... (1-2 Min.)", COLORS['warning'])
        self.module_label.configure(text="⏳ Installiere...", fg=COLORS['warning'])

        def do_install():
            cmd = """
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
$m = Get-Module -ListAvailable -Name ExchangeOnlineManagement
if ($m) { Write-Output "INSTALL_OK:$($m.Version)" } else { Write-Output "INSTALL_FAILED" }
"""
            _, stdout, stderr = self.ps.execute(cmd, timeout=300)
            if "INSTALL_OK:" in stdout:
                ver = stdout.split("INSTALL_OK:")[1].strip().split("\n")[0]
                self.root.after(0, lambda: self._module_ok(ver))
                self.root.after(0, lambda: messagebox.showinfo("Erfolg", f"Modul v{ver} installiert!"))
            else:
                self.root.after(0, lambda: self.log(f"❌ Installation fehlgeschlagen: {stderr}", COLORS['error']))

        threading.Thread(target=do_install, daemon=True).start()

    def connect(self):
        admin = self.admin_entry.get().strip()
        if not admin:
            messagebox.showwarning("Fehlt", "Bitte Admin-UPN eingeben!")
            return

        self.log("🔄 Verbinde zu Exchange Online...", COLORS['warning'])
        self.connect_btn.configure(state=tk.DISABLED)

        def do_connect():
            cmd = f'Connect-ExchangeOnline -UserPrincipalName "{admin}" -ShowBanner:$false'
            success, _, stderr = self.ps.execute(cmd, timeout=180)
            if success:
                verify_cmd = 'Get-OrganizationConfig | Select-Object -ExpandProperty Name'
                v_ok, v_out, _ = self.ps.execute(verify_cmd, timeout=30)
                org = v_out.strip().split("\n")[0] if v_ok and v_out.strip() else ""
                self.root.after(0, lambda: self._connected(org))
            else:
                self.root.after(0, lambda: self._connect_failed(stderr))

        threading.Thread(target=do_connect, daemon=True).start()

    def _connected(self, org):
        self.connected = True
        self.connect_btn.configure(text="✅ Verbunden", bg=COLORS['success'])
        name = f" mit {org}" if org else ""
        self.log(f"✅ Verbunden{name}!", COLORS['success'])
        messagebox.showinfo("Verbunden", f"Erfolgreich verbunden{name}!\n\nDaten werden geladen...")
        self.load_all_data()

    def _connect_failed(self, error):
        self.connect_btn.configure(state=tk.NORMAL)
        self.log(f"❌ Verbindung fehlgeschlagen: {error}", COLORS['error'])
        messagebox.showerror("Fehler", f"Verbindung fehlgeschlagen:\n\n{error}")

    def disconnect(self):
        self.log("🔌 Trenne...", COLORS['warning'])
        self.ps.execute("Disconnect-ExchangeOnline -Confirm:$false", timeout=30)
        self.connected = False
        self.all_mailboxes = []
        for k in self.all_groups:
            self.all_groups[k] = []
        self.connect_btn.configure(text="🔌 Verbinden", bg=COLORS['accent'])
        self.connect_btn.configure(state=tk.NORMAL)

        # Alle Combos zurücksetzen
        for combo in [self.mb_target_combo, self.mb_user_combo]:
            combo.configure(values=[], state="disabled")
            combo.set("-- Erst verbinden --")
        for key in ['teams', 'verteiler', 'security']:
            for attr in ['group_combo', 'user_combo']:
                c = getattr(self, f'{key}_{attr}')
                c.configure(values=[], state="disabled")
                c.set("-- Erst verbinden --")

        self.log("✅ Getrennt.", COLORS['success'])

    # ============================================================
    # DATEN LADEN
    # ============================================================

    def load_all_data(self):
        self.log("📥 Lade Postfächer & Gruppen...", COLORS['warning'])

        def do_load():
            # Postfächer
            cmd_mb = 'Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails | ConvertTo-Json -Compress'
            ok_mb, out_mb, _ = self.ps.execute(cmd_mb, timeout=180)

            # Unified Groups (Teams / M365)
            cmd_ug = 'Get-UnifiedGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, GroupType | ConvertTo-Json -Compress'
            ok_ug, out_ug, _ = self.ps.execute(cmd_ug, timeout=180)

            # Verteilerlisten
            cmd_dl = 'Get-DistributionGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, GroupType | ConvertTo-Json -Compress'
            ok_dl, out_dl, _ = self.ps.execute(cmd_dl, timeout=180)

            # Mail-enabled Security Groups
            # Sicherheitsgruppen sind Distribution Groups mit GroupType "Universal, SecurityEnabled"
            # Wir filtern sie aus den Distribution Groups heraus
            self.root.after(0, lambda: self._data_loaded(
                ok_mb, out_mb, ok_ug, out_ug, ok_dl, out_dl))

        threading.Thread(target=do_load, daemon=True).start()

    def _parse_json(self, output):
        if not output or not output.strip():
            return []
        try:
            start = output.find('[')
            start_obj = output.find('{')
            if start == -1 and start_obj == -1:
                return []
            if start == -1 or (start_obj != -1 and start_obj < start):
                start = start_obj
            data = json.loads(output[start:])
            if isinstance(data, dict):
                data = [data]
            return data
        except:
            return []

    def _data_loaded(self, ok_mb, out_mb, ok_ug, out_ug, ok_dl, out_dl):
        # Postfächer
        if ok_mb:
            data = self._parse_json(out_mb)
            self.all_mailboxes = []
            for mb in data:
                email = mb.get('PrimarySmtpAddress', '')
                name = mb.get('DisplayName', '')
                mtype = mb.get('RecipientTypeDetails', 'UserMailbox')
                if email:
                    is_shared = 'Shared' in mtype
                    prefix = "👥" if is_shared else "👤"
                    self.all_mailboxes.append({
                        'display': f"{prefix} {name} <{email}>",
                        'email': email, 'name': name,
                        'type': 'shared' if is_shared else 'user'
                    })
            self.all_mailboxes.sort(key=lambda x: x['display'])
            self.apply_mb_filter()
            self.log(f"  📧 {len(self.all_mailboxes)} Postfächer geladen", COLORS['text_dim'])

        # Teams / M365 Gruppen
        if ok_ug:
            data = self._parse_json(out_ug)
            self.all_groups['teams'] = []
            for g in data:
                email = g.get('PrimarySmtpAddress', '')
                name = g.get('DisplayName', '')
                if email:
                    self.all_groups['teams'].append({
                        'display': f"👥 {name} <{email}>",
                        'email': email, 'name': name
                    })
            self.all_groups['teams'].sort(key=lambda x: x['display'])
            self._update_group_combos('teams')
            self.log(f"  👥 {len(self.all_groups['teams'])} Teams/M365 Gruppen geladen", COLORS['text_dim'])

        # Verteilerlisten & Sicherheitsgruppen
        if ok_dl:
            data = self._parse_json(out_dl)
            self.all_groups['verteiler'] = []
            self.all_groups['security'] = []
            for g in data:
                email = g.get('PrimarySmtpAddress', '')
                name = g.get('DisplayName', '')
                gtype = str(g.get('GroupType', ''))
                if email:
                    entry = {
                        'display': f"{'🔒' if 'Security' in gtype else '📨'} {name} <{email}>",
                        'email': email, 'name': name
                    }
                    if 'Security' in gtype:
                        self.all_groups['security'].append(entry)
                    else:
                        self.all_groups['verteiler'].append(entry)

            self.all_groups['verteiler'].sort(key=lambda x: x['display'])
            self.all_groups['security'].sort(key=lambda x: x['display'])
            self._update_group_combos('verteiler')
            self._update_group_combos('security')
            self.log(f"  📨 {len(self.all_groups['verteiler'])} Verteilerlisten geladen", COLORS['text_dim'])
            self.log(f"  🔒 {len(self.all_groups['security'])} Sicherheitsgruppen geladen", COLORS['text_dim'])

        total = len(self.all_mailboxes) + sum(len(v) for v in self.all_groups.values())
        self.log(f"✅ Insgesamt {total} Objekte geladen!", COLORS['success'])

    def _update_group_combos(self, key):
        items = [g['display'] for g in self.all_groups[key]]
        group_combo = getattr(self, f'{key}_group_combo')
        user_combo = getattr(self, f'{key}_user_combo')
        group_combo.configure(values=items, state="normal")
        group_combo.set("")

        # User-Combo bekommt immer alle Postfächer (Benutzer)
        user_items = [mb['display'] for mb in self.all_mailboxes]
        user_combo.configure(values=user_items, state="normal")
        user_combo.set("")

    # ============================================================
    # POSTFACH FUNKTIONEN
    # ============================================================

    def _toggle_automap(self):
        self.automap_cb.configure(state=tk.NORMAL if self.fullaccess_var.get() else tk.DISABLED)

    def apply_mb_filter(self):
        ft = self.mb_filter.get()
        if ft == "all":
            filtered = self.all_mailboxes
        else:
            filtered = [m for m in self.all_mailboxes if m['type'] == ft]
        items = [m['display'] for m in filtered]
        self.mb_target_combo.configure(values=items, state="normal")
        self.mb_user_combo.configure(values=items, state="normal")

    def filter_mailboxes(self, event=None):
        search = self.mb_search.get().lower()
        ft = self.mb_filter.get()
        base = self.all_mailboxes if ft == "all" else [m for m in self.all_mailboxes if m['type'] == ft]
        filtered = [m for m in base if search in m['display'].lower()] if search else base
        items = [m['display'] for m in filtered]
        self.mb_target_combo.configure(values=items)
        self.mb_user_combo.configure(values=items)

    def filter_groups(self, key):
        search = getattr(self, f'{key}_search').get().lower()
        base = self.all_groups[key]
        filtered = [g for g in base if search in g['display'].lower()] if search else base
        getattr(self, f'{key}_group_combo').configure(values=[g['display'] for g in filtered])

    def _get_email(self, selection):
        if '<' in selection and '>' in selection:
            return selection.split('<')[1].split('>')[0]
        return selection

    def _validate_mb(self):
        mb = self.mb_target_combo.get().strip()
        user = self.mb_user_combo.get().strip()
        if not mb or mb.startswith("--"):
            messagebox.showwarning("Fehlt", "Bitte Ziel-Postfach auswählen!")
            return False
        if not user or user.startswith("--"):
            messagebox.showwarning("Fehlt", "Bitte Benutzer auswählen!")
            return False
        if not self.fullaccess_var.get() and not self.sendas_var.get():
            messagebox.showwarning("Fehlt", "Mindestens eine Berechtigung auswählen!")
            return False
        return True

    def add_mb_permission(self):
        if not self._validate_mb():
            return
        mailbox = self._get_email(self.mb_target_combo.get())
        user = self._get_email(self.mb_user_combo.get())

        msg = f"Berechtigungen HINZUFÜGEN?\n\n📬 {mailbox}\n👤 {user}\n\n"
        msg += f"Vollzugriff: {'✅' if self.fullaccess_var.get() else '❌'}\n"
        if self.fullaccess_var.get():
            msg += f"AutoMapping: {'✅' if self.automapping_var.get() else '❌'}\n"
        msg += f"Senden als: {'✅' if self.sendas_var.get() else '❌'}"
        if not messagebox.askyesno("Bestätigen", msg):
            return

        def do_add():
            errors = []
            if self.fullaccess_var.get():
                self.root.after(0, lambda: self.log("  📂 Vollzugriff...", COLORS['warning']))
                am = "$true" if self.automapping_var.get() else "$false"
                ok, _, err = self.ps.execute(f'Add-MailboxPermission -Identity "{mailbox}" -User "{user}" -AccessRights FullAccess -AutoMapping {am}')
                if ok: self.root.after(0, lambda: self.log("  ✅ Vollzugriff hinzugefügt", COLORS['success']))
                else: errors.append(err)
            if self.sendas_var.get():
                self.root.after(0, lambda: self.log("  ✉️ Senden als...", COLORS['warning']))
                ok, _, err = self.ps.execute(f'Add-RecipientPermission -Identity "{mailbox}" -Trustee "{user}" -AccessRights SendAs -Confirm:$false')
                if ok: self.root.after(0, lambda: self.log("  ✅ Senden als hinzugefügt", COLORS['success']))
                else: errors.append(err)
            self.root.after(0, lambda: self._action_done(errors, "hinzugefügt"))

        threading.Thread(target=do_add, daemon=True).start()

    def remove_mb_permission(self):
        if not self._validate_mb():
            return
        mailbox = self._get_email(self.mb_target_combo.get())
        user = self._get_email(self.mb_user_combo.get())

        if not messagebox.askyesno("⚠️ Warnung",
                                   f"Berechtigungen ENTFERNEN?\n\n📬 {mailbox}\n👤 {user}\n\n⚠️ Nicht rückgängig!",
                                   icon="warning"):
            return

        def do_remove():
            errors = []
            if self.fullaccess_var.get():
                self.root.after(0, lambda: self.log("  📂 Entferne Vollzugriff...", COLORS['warning']))
                ok, _, err = self.ps.execute(f'Remove-MailboxPermission -Identity "{mailbox}" -User "{user}" -AccessRights FullAccess -Confirm:$false')
                if ok: self.root.after(0, lambda: self.log("  ✅ Vollzugriff entfernt", COLORS['success']))
                else: errors.append(err)
            if self.sendas_var.get():
                self.root.after(0, lambda: self.log("  ✉️ Entferne Senden als...", COLORS['warning']))
                ok, _, err = self.ps.execute(f'Remove-RecipientPermission -Identity "{mailbox}" -Trustee "{user}" -AccessRights SendAs -Confirm:$false')
                if ok: self.root.after(0, lambda: self.log("  ✅ Senden als entfernt", COLORS['success']))
                else: errors.append(err)
            self.root.after(0, lambda: self._action_done(errors, "entfernt"))

        threading.Thread(target=do_remove, daemon=True).start()

    # ============================================================
    # GRUPPEN FUNKTIONEN
    # ============================================================

    def _validate_group(self, key):
        group = getattr(self, f'{key}_group_combo').get().strip()
        user = getattr(self, f'{key}_user_combo').get().strip()
        if not group or group.startswith("--"):
            messagebox.showwarning("Fehlt", "Bitte Gruppe auswählen!")
            return False
        if not user or user.startswith("--"):
            messagebox.showwarning("Fehlt", "Bitte Benutzer auswählen!")
            return False
        return True

    def add_group_member(self, key):
        if not self._validate_group(key):
            return
        group = self._get_email(getattr(self, f'{key}_group_combo').get())
        user = self._get_email(getattr(self, f'{key}_user_combo').get())
        role = getattr(self, f'{key}_role_var').get()

        label = MODULES[key]
        if not messagebox.askyesno("Bestätigen",
                                   f"Benutzer hinzufügen?\n\n{label}\n📋 {group}\n👤 {user}" +
                                   (f"\n🏷️ Rolle: {role}" if key == 'teams' else "")):
            return

        def do_add():
            self.root.after(0, lambda: self.log(f"  ➕ Füge {user} zu {group} hinzu...", COLORS['warning']))

            if key == 'teams':
                if role == 'Owner':
                    cmd = f'Add-UnifiedGroupLinks -Identity "{group}" -LinkType Owners -Links "{user}"'
                else:
                    cmd = f'Add-UnifiedGroupLinks -Identity "{group}" -LinkType Members -Links "{user}"'
            else:
                cmd = f'Add-DistributionGroupMember -Identity "{group}" -Member "{user}"'

            ok, _, err = self.ps.execute(cmd)
            if ok:
                self.root.after(0, lambda: self.log(f"  ✅ {user} hinzugefügt!", COLORS['success']))
                self.root.after(0, lambda: self._action_done([], "hinzugefügt"))
            else:
                self.root.after(0, lambda: self._action_done([err], "hinzugefügt"))

        threading.Thread(target=do_add, daemon=True).start()

    def remove_group_member(self, key):
        if not self._validate_group(key):
            return
        group = self._get_email(getattr(self, f'{key}_group_combo').get())
        user = self._get_email(getattr(self, f'{key}_user_combo').get())

        if not messagebox.askyesno("⚠️ Warnung",
                                   f"Benutzer ENTFERNEN?\n\n📋 {group}\n👤 {user}\n\n⚠️ Nicht rückgängig!",
                                   icon="warning"):
            return

        def do_remove():
            self.root.after(0, lambda: self.log(f"  ➖ Entferne {user} aus {group}...", COLORS['warning']))

            if key == 'teams':
                # Bei Teams erst Member, dann Owner entfernen
                cmd = f'Remove-UnifiedGroupLinks -Identity "{group}" -LinkType Members -Links "{user}" -Confirm:$false'
                ok, _, err = self.ps.execute(cmd)
                # Auch Owner entfernen falls gesetzt
                self.ps.execute(f'Remove-UnifiedGroupLinks -Identity "{group}" -LinkType Owners -Links "{user}" -Confirm:$false')
            else:
                cmd = f'Remove-DistributionGroupMember -Identity "{group}" -Member "{user}" -Confirm:$false'
                ok, _, err = self.ps.execute(cmd)

            if ok:
                self.root.after(0, lambda: self.log(f"  ✅ {user} entfernt!", COLORS['success']))
                self.root.after(0, lambda: self._action_done([], "entfernt"))
            else:
                self.root.after(0, lambda: self._action_done([err], "entfernt"))

        threading.Thread(target=do_remove, daemon=True).start()

    def show_members(self, key):
        group_combo = getattr(self, f'{key}_group_combo')
        members_text = getattr(self, f'{key}_members_text')
        group = group_combo.get().strip()

        if not group or group.startswith("--"):
            messagebox.showwarning("Fehlt", "Bitte erst eine Gruppe auswählen!")
            return

        group_email = self._get_email(group)
        self.log(f"  📋 Lade Mitglieder von {group_email}...", COLORS['warning'])

        def do_load():
            if key == 'teams':
                cmd_members = f'Get-UnifiedGroupLinks -Identity "{group_email}" -LinkType Members | Select-Object Name, PrimarySmtpAddress | ConvertTo-Json -Compress'
                cmd_owners = f'Get-UnifiedGroupLinks -Identity "{group_email}" -LinkType Owners | Select-Object Name, PrimarySmtpAddress | ConvertTo-Json -Compress'
                ok_m, out_m, _ = self.ps.execute(cmd_members, timeout=60)
                ok_o, out_o, _ = self.ps.execute(cmd_owners, timeout=60)

                members = self._parse_json(out_m) if ok_m else []
                owners = self._parse_json(out_o) if ok_o else []

                lines = [f"=== {group_email} ===", ""]
                if owners:
                    lines.append(f"👑 Besitzer ({len(owners)}):")
                    for o in owners:
                        lines.append(f"  • {o.get('Name', '')} <{o.get('PrimarySmtpAddress', '')}>")
                    lines.append("")
                lines.append(f"👤 Mitglieder ({len(members)}):")
                for m in members:
                    lines.append(f"  • {m.get('Name', '')} <{m.get('PrimarySmtpAddress', '')}>")

                result = "\n".join(lines)
                total = len(members)
            else:
                cmd = f'Get-DistributionGroupMember -Identity "{group_email}" -ResultSize Unlimited | Select-Object Name, PrimarySmtpAddress | ConvertTo-Json -Compress'
                ok, out, _ = self.ps.execute(cmd, timeout=60)
                members = self._parse_json(out) if ok else []

                lines = [f"=== {group_email} ===", "", f"👤 Mitglieder ({len(members)}):"]
                for m in members:
                    lines.append(f"  • {m.get('Name', '')} <{m.get('PrimarySmtpAddress', '')}>")

                result = "\n".join(lines)
                total = len(members)

            self.root.after(0, lambda: self._show_members_result(key, result, total))

        threading.Thread(target=do_load, daemon=True).start()

    def _show_members_result(self, key, text, count):
        members_text = getattr(self, f'{key}_members_text')
        members_text.configure(state=tk.NORMAL)
        members_text.delete('1.0', tk.END)
        members_text.insert(tk.END, text)
        members_text.configure(state=tk.DISABLED)
        self.log(f"  ✅ {count} Mitglieder geladen", COLORS['success'])

    # ============================================================
    # COMMON
    # ============================================================

    def _action_done(self, errors, action_word):
        if errors:
            for e in errors:
                self.log(f"  ❌ {e}", COLORS['error'])
            messagebox.showerror("Fehler", f"Aktion teilweise fehlgeschlagen.\nSiehe Protokoll.")
        else:
            messagebox.showinfo("Erfolg", f"✅ Erfolgreich {action_word}!")

    def log(self, message, color=None):
        ts = datetime.now().strftime("%H:%M:%S")
        tag_map = {COLORS['success']: 'success', COLORS['error']: 'error',
                   COLORS['warning']: 'warning', COLORS['accent']: 'accent',
                   COLORS['text_dim']: 'dim'}
        tag = tag_map.get(color, 'info')
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] ", 'info')
        self.log_text.insert(tk.END, f"{message}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def cleanup(self):
        try:
            self.ps.execute("Disconnect-ExchangeOnline -Confirm:$false", timeout=10)
        except:
            pass
        self.ps.stop()


def main():
    root = tk.Tk()
    try:
        root.iconbitmap('exchange.ico')
    except:
        pass

    app = M365AdminTool(root)

    def on_closing():
        if messagebox.askokcancel("Beenden",
                                  "Programm beenden?\n\nExchange Online Session wird getrennt."):
            app.cleanup()
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()