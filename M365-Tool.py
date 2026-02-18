#!/usr/bin/env python3
"""
Microsoft 365 Admin Tool v4.0
Erstellt für Kaulich IT Systems GmbH

Features:
- Postfachberechtigungen (FullAccess, SendAs, AutoMapping)
- Microsoft 365 Gruppen / Teams-Gruppen (Mitglieder/Besitzer)
- Verteilerlisten (Distribution Groups)
- Sicherheitsgruppen (Mail-Enabled Security Groups)
- Benutzer-Offboarding (10 Schritte, Bericht-Export)
- Automatische Modulprüfung & Installation
- Dark Theme UI

Voraussetzungen:
- Windows mit PowerShell 5.1+
- ExchangeOnlineManagement Modul (wird automatisch installiert)
- Keine zusätzlichen Python-Pakete nötig (nur Standardbibliothek)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import threading
import json
import queue
import time
import os
from datetime import datetime

COLORS = {
    'bg': '#1e1e2e', 'bg_panel': '#252536', 'bg_input': '#2e2e42',
    'bg_header': '#16161e', 'accent': '#7aa2f7', 'accent_hover': '#89b4fa',
    'success': '#9ece6a', 'warning': '#e0af68', 'error': '#f7768e',
    'text': '#c0caf5', 'text_dim': '#565f89', 'text_muted': '#414868',
    'border': '#3b3b54', 'button_danger': '#f7768e', 'button_success': '#9ece6a',
    'purple': '#bb9af7', 'cyan': '#7dcfff', 'orange': '#ff9e64', 'green': '#73daca',
}

MODULES = {
    'postfach':    '📧 Postfachberechtigungen',
    'teams':       '👥 Microsoft 365 / Teams Gruppen',
    'verteiler':   '📨 Verteilerlisten',
    'security':    '🔒 Sicherheitsgruppen',
    'offboarding': '🚪 Benutzer-Offboarding',
}

OB_STEPS = [
    ('sign_in',          '🔒 Anmeldung blockieren'),
    ('reset_pw',         '🔑 Passwort zurücksetzen (zufällig)'),
    ('remove_groups',    '👥 Aus allen Gruppen entfernen'),
    ('remove_licenses',  '📋 Alle Lizenzen entziehen'),
    ('convert_shared',   '📧 In Shared Mailbox konvertieren'),
    ('set_ooo',          '✈️ Abwesenheitsnachricht setzen'),
    ('forwarding',       '📨 Mail-Weiterleitung einrichten'),
    ('hide_gal',         '👻 Aus Adressbuch ausblenden'),
    ('disable_sync',     '📱 ActiveSync/Mobile deaktivieren'),
    ('remove_delegates', '🔓 Postfach-Delegierungen entfernen'),
]

OB_STEP_NAMES = {k: v.split(' ', 1)[1] for k, v in OB_STEPS}


# ================================================================
# PowerShell Session
# ================================================================
class PowerShellSession:
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
            encoding='utf-8', errors='replace', bufsize=1)
        threading.Thread(target=self._read, daemon=True).start()
        self._cmd('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8')

    def _read(self):
        while self.process and self.process.poll() is None:
            try:
                line = self.process.stdout.readline()
                if line:
                    self.output_queue.put(line)
            except:
                break

    def _cmd(self, c):
        if self.process and self.process.poll() is None:
            self.process.stdin.write(c + "\n")
            self.process.stdin.flush()

    def execute(self, command, timeout=120):
        if not self.process or self.process.poll() is not None:
            return False, "", "PowerShell-Session nicht aktiv"
        while not self.output_queue.empty():
            try: self.output_queue.get_nowait()
            except: break

        ts = int(time.time() * 1000)
        end_m = f"###END_{ts}###"
        err_m = f"###ERROR_{ts}###"

        self._cmd(f"""
try {{
    {command}
}} catch {{
    Write-Output "{err_m}$($_.Exception.Message)"
}}
Write-Output "{end_m}"
""")
        out_lines, err_lines = [], []
        t0 = time.time()
        while True:
            if time.time() - t0 > timeout:
                return False, "", "Timeout"
            try:
                line = self.output_queue.get(timeout=0.5).rstrip()
                if end_m in line:
                    break
                elif err_m in line:
                    err_lines.append(line.split(err_m, 1)[1])
                else:
                    out_lines.append(line)
            except queue.Empty:
                continue
        return (not bool(err_lines)), "\n".join(out_lines), "\n".join(err_lines)

    def stop(self):
        if self.process:
            try:
                self._cmd("exit")
                self.process.terminate()
            except:
                pass
            self.process = None


# ================================================================
# Styled Button
# ================================================================
class StyledButton(tk.Canvas):
    def __init__(self, parent, text, command=None, bg=COLORS['accent'],
                 fg='#ffffff', width=160, height=34, font_size=10, **kw):
        super().__init__(parent, width=width, height=height,
                         bg=parent.cget('bg'), highlightthickness=0, **kw)
        self.command = command
        self.bg_color = bg
        self.fg_color = fg
        self.hover_color = self._adj(bg, 25)
        self.disabled_color = COLORS['text_muted']
        self.text = text
        self._w = width
        self._h = height
        self.font_size = font_size
        self.enabled = True
        self._draw(self.bg_color)
        self.bind('<Enter>', lambda e: self._draw(self.hover_color) if self.enabled else None)
        self.bind('<Leave>', lambda e: self._draw(self.bg_color if self.enabled else self.disabled_color))
        self.bind('<Button-1>', lambda e: self.command() if self.enabled and self.command else None)

    def _adj(self, c, a):
        try:
            r, g, b = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
            return f'#{min(255,r+a):02x}{min(255,g+a):02x}{min(255,b+a):02x}'
        except: return c

    def _draw(self, color):
        self.delete('all')
        r, w, h = 6, self._w, self._h
        for cx, cy, s, e in [(0,0,r*2,r*2,90,90),(w-r*2,0,w,r*2,0,90),
                               (0,h-r*2,r*2,h,180,90),(w-r*2,h-r*2,w,h,270,90)]:
            self.create_arc(cx, cy, s, e, start=e if isinstance(e, int) else s,
                            extent=90, fill=color, outline=color)
        # Simplified: just draw rects
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
            self.hover_color = self._adj(kw['bg'], 25)
            self._draw(self.bg_color)


# ================================================================
# Hauptanwendung
# ================================================================
class M365AdminTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Microsoft 365 Admin Tool v4.0 — Kaulich IT Systems GmbH")
        self.root.geometry("900x900")
        self.root.minsize(820, 700)
        self.root.configure(bg=COLORS['bg'])

        self.ps = PowerShellSession()
        self.ps.start()
        self.connected = False
        self.all_mailboxes = []
        self.all_groups = {'teams': [], 'verteiler': [], 'security': []}
        self.current_module = tk.StringVar(value='postfach')
        self.ob_report_lines = []

        self._build_ui()
        self.log("🚀 M365 Admin Tool v4.0 gestartet", COLORS['success'])
        self.root.after(500, self.check_module)

    # ============================================================
    #  UI AUFBAU
    # ============================================================
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=COLORS['bg_header'], height=50)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="⚡ Microsoft 365 Admin Tool", font=('Segoe UI', 16, 'bold'),
                 fg=COLORS['accent'], bg=COLORS['bg_header']).pack(side=tk.LEFT, padx=15)
        tk.Label(hdr, text="v4.0 — Kaulich IT Systems GmbH", font=('Segoe UI', 9),
                 fg=COLORS['text_dim'], bg=COLORS['bg_header']).pack(side=tk.RIGHT, padx=15)

        # Scrollbarer Bereich
        outer = tk.Frame(self.root, bg=COLORS['bg'])
        outer.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(outer, bg=COLORS['bg'], highlightthickness=0)
        sb = tk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        self.main = tk.Frame(canvas, bg=COLORS['bg'])
        self.main.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        self._cw = canvas.create_window((0, 0), window=self.main, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=15, pady=10)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(self._cw, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        m = self.main

        # --- Verbindung ---
        cf = self._sec(m, "🔐 Verbindung")
        ci = tk.Frame(cf, bg=COLORS['bg_panel']); ci.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(ci, text="Admin-UPN:", font=('Segoe UI', 10),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).grid(row=0, column=0, sticky=tk.W, padx=(0,10))
        self.admin_entry = tk.Entry(ci, width=35, font=('Segoe UI', 10), bg=COLORS['bg_input'],
                                    fg=COLORS['text'], insertbackground=COLORS['text'], relief=tk.FLAT,
                                    highlightthickness=1, highlightbackground=COLORS['border'],
                                    highlightcolor=COLORS['accent'])
        self.admin_entry.grid(row=0, column=1, padx=(0,10), ipady=5)
        self.connect_btn = StyledButton(ci, "🔌 Verbinden", command=self.connect, width=130)
        self.connect_btn.grid(row=0, column=2)
        self.module_label = tk.Label(cf, text="⏳ Prüfe Module...", font=('Segoe UI', 9),
                                     fg=COLORS['warning'], bg=COLORS['bg_panel'])
        self.module_label.pack(fill=tk.X, padx=10, pady=(0,5))
        tk.Label(cf, text="⚠️ Session wird beim Beenden automatisch getrennt",
                 font=('Segoe UI', 9), fg=COLORS['text_dim'],
                 bg=COLORS['bg_panel']).pack(fill=tk.X, padx=10, pady=(0,8))

        # --- Modul-Auswahl ---
        mf = self._sec(m, "🔧 Funktion")
        mi = tk.Frame(mf, bg=COLORS['bg_panel']); mi.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(mi, text="Modul:", font=('Segoe UI', 10, 'bold'),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0,10))
        self.module_combo = ttk.Combobox(mi, width=55, state="readonly", font=('Segoe UI', 10),
                                         values=list(MODULES.values()))
        self.module_combo.pack(side=tk.LEFT); self.module_combo.current(0)
        self.module_combo.bind('<<ComboboxSelected>>', self._on_mod)

        # --- Dynamischer Bereich ---
        self.dyn = tk.Frame(m, bg=COLORS['bg']); self.dyn.pack(fill=tk.X, pady=(5,0))
        self.mod_frames = {}
        self._build_postfach()
        self._build_group('teams', "Teams-Gruppe")
        self._build_group('verteiler', "Verteilerliste")
        self._build_group('security', "Sicherheitsgruppe")
        self._build_offboarding()

        # --- Protokoll ---
        lc = tk.Frame(m, bg=COLORS['bg']); lc.pack(fill=tk.X, pady=(8,0))
        tk.Label(lc, text="📋 Protokoll", font=('Segoe UI', 11, 'bold'),
                 fg=COLORS['accent'], bg=COLORS['bg']).pack(anchor=tk.W, pady=(0,4))
        ls = tk.Frame(lc, bg=COLORS['bg_panel'], highlightbackground=COLORS['border'], highlightthickness=1)
        ls.pack(fill=tk.X)
        li = tk.Frame(ls, bg=COLORS['bg_panel']); li.pack(fill=tk.X, padx=10, pady=10)
        self.log_text = tk.Text(li, height=8, font=('Consolas', 9), bg=COLORS['bg_input'],
                                fg=COLORS['text'], relief=tk.FLAT, wrap=tk.WORD,
                                highlightthickness=1, highlightbackground=COLORS['border'])
        lsb = tk.Scrollbar(li, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lsb.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True); lsb.pack(side=tk.RIGHT, fill=tk.Y)
        for tag, c in [('success',COLORS['success']),('error',COLORS['error']),
                        ('warning',COLORS['warning']),('info',COLORS['text']),
                        ('accent',COLORS['accent']),('dim',COLORS['text_dim'])]:
            self.log_text.tag_configure(tag, foreground=c)

        # Style
        st = ttk.Style(); st.theme_use('clam')
        st.configure('TCombobox', fieldbackground=COLORS['bg_input'], background=COLORS['bg_panel'],
                     foreground=COLORS['text'], arrowcolor=COLORS['text'],
                     bordercolor=COLORS['border'], lightcolor=COLORS['border'], darkcolor=COLORS['border'])
        st.configure("red.Horizontal.TProgressbar", troughcolor=COLORS['bg_input'],
                     background=COLORS['error'], bordercolor=COLORS['border'])

        self._show_mod('postfach')

    # --- UI Helpers ---
    def _sec(self, parent, title):
        c = tk.Frame(parent, bg=COLORS['bg']); c.pack(fill=tk.X, pady=(0,8))
        tk.Label(c, text=title, font=('Segoe UI', 11, 'bold'),
                 fg=COLORS['accent'], bg=COLORS['bg']).pack(anchor=tk.W, pady=(0,4))
        f = tk.Frame(c, bg=COLORS['bg_panel'], highlightbackground=COLORS['border'], highlightthickness=1)
        f.pack(fill=tk.X); return f

    def _sec_in(self, parent, title):
        c = tk.Frame(parent, bg=COLORS['bg']); c.pack(fill=tk.X, pady=(0,5))
        tk.Label(c, text=title, font=('Segoe UI', 11, 'bold'),
                 fg=COLORS['purple'], bg=COLORS['bg']).pack(anchor=tk.W, pady=(0,4))
        f = tk.Frame(c, bg=COLORS['bg_panel'], highlightbackground=COLORS['border'], highlightthickness=1)
        f.pack(fill=tk.X); return f

    def _combo_row(self, parent, label, hint=None):
        r = tk.Frame(parent, bg=COLORS['bg_panel']); r.pack(fill=tk.X, padx=10, pady=(5,0))
        tk.Label(r, text=label, font=('Segoe UI', 10), fg=COLORS['text'],
                 bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0,10))
        cb = ttk.Combobox(r, width=60, state="disabled", font=('Segoe UI', 9))
        cb.pack(side=tk.LEFT, fill=tk.X, expand=True); cb.set("-- Erst verbinden --")
        if hint:
            hr = tk.Frame(parent, bg=COLORS['bg_panel']); hr.pack(fill=tk.X, padx=10, pady=(1,0))
            tk.Label(hr, text=hint, font=('Segoe UI', 8), fg=COLORS['text_dim'],
                     bg=COLORS['bg_panel']).pack(anchor=tk.E)
        return cb

    def _search(self, parent, cb):
        r = tk.Frame(parent, bg=COLORS['bg_panel']); r.pack(fill=tk.X, padx=10, pady=(8,8))
        tk.Label(r, text="🔍", font=('Segoe UI', 10), fg=COLORS['text'],
                 bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0,5))
        e = tk.Entry(r, width=55, font=('Segoe UI', 10), bg=COLORS['bg_input'], fg=COLORS['text'],
                     insertbackground=COLORS['text'], relief=tk.FLAT, highlightthickness=1,
                     highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'])
        e.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4)
        e.bind('<KeyRelease>', cb); return e

    def _get_email(self, s):
        return s.split('<')[1].split('>')[0] if '<' in s and '>' in s else s

    # ============================================================
    #  MODUL: POSTFACH
    # ============================================================
    def _build_postfach(self):
        f = tk.Frame(self.dyn, bg=COLORS['bg']); self.mod_frames['postfach'] = f
        s = self._sec_in(f, "📧 Postfachberechtigungen verwalten")
        self.mb_target = self._combo_row(s, "Ziel-Postfach:", "(Postfach, auf das zugegriffen werden soll)")
        self.mb_user = self._combo_row(s, "Benutzer:      ", "(Benutzer, der Zugriff erhalten/verlieren soll)")
        self.mb_search = self._search(s, self._filter_mb)

        fr = tk.Frame(s, bg=COLORS['bg_panel']); fr.pack(fill=tk.X, padx=10, pady=(0,5))
        tk.Label(fr, text="Typ:", font=('Segoe UI', 9, 'bold'), fg=COLORS['text'],
                 bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0,8))
        self.mb_filter = tk.StringVar(value="all")
        for t, v in [("Alle","all"),("👤 Benutzer","user"),("👥 Shared","shared")]:
            tk.Radiobutton(fr, text=t, variable=self.mb_filter, value=v, font=('Segoe UI', 9),
                           fg=COLORS['text'], bg=COLORS['bg_panel'], selectcolor=COLORS['bg_input'],
                           activebackground=COLORS['bg_panel'], command=self._apply_mb_filter
                           ).pack(side=tk.LEFT, padx=(0,12))

        pr = tk.Frame(s, bg=COLORS['bg_panel']); pr.pack(fill=tk.X, padx=10, pady=(5,5))
        self.fullaccess_var = tk.BooleanVar(value=True)
        self.automapping_var = tk.BooleanVar(value=True)
        self.sendas_var = tk.BooleanVar(value=True)
        tk.Checkbutton(pr, text="📂 Vollzugriff", variable=self.fullaccess_var, font=('Segoe UI', 10),
                        fg=COLORS['text'], bg=COLORS['bg_panel'], selectcolor=COLORS['bg_input'],
                        activebackground=COLORS['bg_panel'],
                        command=lambda: self.automap_cb.configure(
                            state=tk.NORMAL if self.fullaccess_var.get() else tk.DISABLED)
                        ).pack(side=tk.LEFT)
        self.automap_cb = tk.Checkbutton(pr, text="🔗 AutoMapping", variable=self.automapping_var,
                                          font=('Segoe UI', 10), fg=COLORS['text'], bg=COLORS['bg_panel'],
                                          selectcolor=COLORS['bg_input'], activebackground=COLORS['bg_panel'])
        self.automap_cb.pack(side=tk.LEFT, padx=(20,0))
        tk.Checkbutton(pr, text="✉️ Senden als", variable=self.sendas_var, font=('Segoe UI', 10),
                        fg=COLORS['text'], bg=COLORS['bg_panel'], selectcolor=COLORS['bg_input'],
                        activebackground=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(20,0))

        br = tk.Frame(s, bg=COLORS['bg_panel']); br.pack(fill=tk.X, padx=10, pady=(5,10))
        StyledButton(br, "✅ Hinzufügen", command=self._add_mb, bg=COLORS['button_success'],
                     fg='#1e1e2e', width=180).pack(side=tk.LEFT, padx=(0,10))
        StyledButton(br, "❌ Entfernen", command=self._rem_mb, bg=COLORS['button_danger'],
                     width=180).pack(side=tk.LEFT, padx=(0,10))
        StyledButton(br, "🔌 Trennen", command=self.disconnect, bg=COLORS['bg_input'],
                     width=100).pack(side=tk.RIGHT)

    # ============================================================
    #  MODUL: GRUPPEN (Teams / Verteiler / Security)
    # ============================================================
    def _build_group(self, key, label):
        f = tk.Frame(self.dyn, bg=COLORS['bg']); self.mod_frames[key] = f
        icon = {'teams':'👥','verteiler':'📨','security':'🔒'}[key]
        s = self._sec_in(f, f"{icon} {label} verwalten")
        gc = self._combo_row(s, f"{label}:", "(Die Gruppe)")
        uc = self._combo_row(s, "Benutzer:     ", "(Der Benutzer)")
        se = self._search(s, lambda e, k=key: self._filter_grp(k))

        rv = tk.StringVar(value="Member")
        if key == 'teams':
            rr = tk.Frame(s, bg=COLORS['bg_panel']); rr.pack(fill=tk.X, padx=10, pady=(0,5))
            tk.Label(rr, text="Rolle:", font=('Segoe UI', 9, 'bold'), fg=COLORS['text'],
                     bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0,8))
            for t, v in [("👤 Mitglied","Member"),("👑 Besitzer","Owner")]:
                tk.Radiobutton(rr, text=t, variable=rv, value=v, font=('Segoe UI', 9),
                               fg=COLORS['text'], bg=COLORS['bg_panel'], selectcolor=COLORS['bg_input'],
                               activebackground=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0,12))

        mf = tk.Frame(s, bg=COLORS['bg_panel']); mf.pack(fill=tk.X, padx=10, pady=(0,5))
        StyledButton(mf, "📋 Mitglieder anzeigen", command=lambda k=key: self._show_members(k),
                     bg=COLORS['bg_input'], width=170, height=28, font_size=9).pack(anchor=tk.W, pady=(2,5))
        mt = tk.Text(mf, height=5, font=('Consolas', 9), bg=COLORS['bg_input'], fg=COLORS['text'],
                     relief=tk.FLAT, wrap=tk.WORD, state=tk.DISABLED,
                     highlightthickness=1, highlightbackground=COLORS['border'])
        mt.pack(fill=tk.X)

        br = tk.Frame(s, bg=COLORS['bg_panel']); br.pack(fill=tk.X, padx=10, pady=(5,10))
        StyledButton(br, "✅ Hinzufügen", command=lambda k=key: self._add_grp(k),
                     bg=COLORS['button_success'], fg='#1e1e2e', width=160).pack(side=tk.LEFT, padx=(0,10))
        StyledButton(br, "❌ Entfernen", command=lambda k=key: self._rem_grp(k),
                     bg=COLORS['button_danger'], width=160).pack(side=tk.LEFT, padx=(0,10))
        StyledButton(br, "🔌 Trennen", command=self.disconnect, bg=COLORS['bg_input'],
                     width=100).pack(side=tk.RIGHT)

        setattr(self, f'{key}_gc', gc); setattr(self, f'{key}_uc', uc)
        setattr(self, f'{key}_se', se); setattr(self, f'{key}_rv', rv)
        setattr(self, f'{key}_mt', mt)

    # ============================================================
    #  MODUL: OFFBOARDING
    # ============================================================
    def _build_offboarding(self):
        f = tk.Frame(self.dyn, bg=COLORS['bg']); self.mod_frames['offboarding'] = f
        s = self._sec_in(f, "🚪 Benutzer-Offboarding")

        self.ob_user = self._combo_row(s, "Benutzer:", "(Der Benutzer, der offgeboardet werden soll)")
        self.ob_search = self._search(s, self._filter_ob)

        wr = tk.Frame(s, bg=COLORS['bg_panel']); wr.pack(fill=tk.X, padx=10, pady=(0,8))
        tk.Label(wr, text="⚠️ ACHTUNG: Offboarding führt mehrere irreversible Aktionen durch!",
                 font=('Segoe UI', 9, 'bold'), fg=COLORS['error'], bg=COLORS['bg_panel']).pack(anchor=tk.W)

        # Checkboxen (2 Spalten)
        cbf = tk.Frame(s, bg=COLORS['bg_panel']); cbf.pack(fill=tk.X, padx=10, pady=(0,5))
        tk.Label(cbf, text="Offboarding-Schritte:", font=('Segoe UI', 10, 'bold'),
                 fg=COLORS['text'], bg=COLORS['bg_panel']).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0,5))
        self.ob_vars = {}
        for i, (key, label) in enumerate(OB_STEPS):
            v = tk.BooleanVar(value=True); self.ob_vars[key] = v
            tk.Checkbutton(cbf, text=label, variable=v, font=('Segoe UI', 9), fg=COLORS['text'],
                           bg=COLORS['bg_panel'], selectcolor=COLORS['bg_input'],
                           activebackground=COLORS['bg_panel']
                           ).grid(row=1+i//2, column=i%2, sticky=tk.W, padx=(0,20), pady=1)

        # Alle an/ab
        tr = tk.Frame(s, bg=COLORS['bg_panel']); tr.pack(fill=tk.X, padx=10, pady=(2,5))
        StyledButton(tr, "☑️ Alle an", command=lambda: [v.set(True) for v in self.ob_vars.values()],
                     bg=COLORS['bg_input'], width=90, height=26, font_size=8).pack(side=tk.LEFT, padx=(0,8))
        StyledButton(tr, "☐ Alle ab", command=lambda: [v.set(False) for v in self.ob_vars.values()],
                     bg=COLORS['bg_input'], width=90, height=26, font_size=8).pack(side=tk.LEFT)

        # OOO Text
        of = tk.Frame(s, bg=COLORS['bg_panel']); of.pack(fill=tk.X, padx=10, pady=(0,5))
        tk.Label(of, text="✈️ Abwesenheitsnachricht (optional):", font=('Segoe UI', 9),
                 fg=COLORS['text_dim'], bg=COLORS['bg_panel']).pack(anchor=tk.W)
        self.ob_ooo = tk.Text(of, height=2, font=('Segoe UI', 9), bg=COLORS['bg_input'], fg=COLORS['text'],
                               relief=tk.FLAT, wrap=tk.WORD, highlightthickness=1,
                               highlightbackground=COLORS['border'], highlightcolor=COLORS['accent'])
        self.ob_ooo.pack(fill=tk.X, pady=(3,0))
        self.ob_ooo.insert('1.0', 'Dieser Mitarbeiter ist nicht mehr im Unternehmen. '
                           'Bitte wenden Sie sich an helpdesk@kaulich-it.de')

        # Weiterleitung
        ff = tk.Frame(s, bg=COLORS['bg_panel']); ff.pack(fill=tk.X, padx=10, pady=(5,5))
        tk.Label(ff, text="📨 Weiterleitung an (optional):", font=('Segoe UI', 9),
                 fg=COLORS['text_dim'], bg=COLORS['bg_panel']).pack(side=tk.LEFT, padx=(0,8))
        self.ob_fwd = ttk.Combobox(ff, width=50, state="disabled", font=('Segoe UI', 9))
        self.ob_fwd.pack(side=tk.LEFT, fill=tk.X, expand=True); self.ob_fwd.set("")

        # Fortschritt
        pf = tk.Frame(s, bg=COLORS['bg_panel']); pf.pack(fill=tk.X, padx=10, pady=(5,5))
        self.ob_prog_lbl = tk.Label(pf, text="", font=('Segoe UI', 9), fg=COLORS['text_dim'],
                                     bg=COLORS['bg_panel']); self.ob_prog_lbl.pack(anchor=tk.W)
        self.ob_prog = ttk.Progressbar(pf, mode='determinate', style="red.Horizontal.TProgressbar")
        self.ob_prog.pack(fill=tk.X, pady=(3,0))

        # Ergebnis
        rf = tk.Frame(s, bg=COLORS['bg_panel']); rf.pack(fill=tk.X, padx=10, pady=(5,5))
        self.ob_result = tk.Text(rf, height=6, font=('Consolas', 9), bg=COLORS['bg_input'],
                                  fg=COLORS['text'], relief=tk.FLAT, wrap=tk.WORD, state=tk.DISABLED,
                                  highlightthickness=1, highlightbackground=COLORS['border'])
        rsb = tk.Scrollbar(rf, orient=tk.VERTICAL, command=self.ob_result.yview)
        self.ob_result.configure(yscrollcommand=rsb.set)
        self.ob_result.pack(side=tk.LEFT, fill=tk.X, expand=True); rsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons
        br = tk.Frame(s, bg=COLORS['bg_panel']); br.pack(fill=tk.X, padx=10, pady=(5,10))
        self.ob_run_btn = StyledButton(br, "🚪 Offboarding starten", command=self._run_ob,
                                        bg=COLORS['error'], width=200)
        self.ob_run_btn.pack(side=tk.LEFT, padx=(0,10))
        self.ob_export_btn = StyledButton(br, "💾 Bericht exportieren", command=self._export_ob,
                                           bg=COLORS['accent'], width=180)
        self.ob_export_btn.pack(side=tk.LEFT, padx=(0,10))
        self.ob_export_btn.configure(state=tk.DISABLED)
        StyledButton(br, "🔌 Trennen", command=self.disconnect, bg=COLORS['bg_input'],
                     width=100).pack(side=tk.RIGHT)

    # ============================================================
    #  MODUL-WECHSEL
    # ============================================================
    def _on_mod(self, e=None):
        for k, lbl in MODULES.items():
            if lbl == self.module_combo.get():
                self._show_mod(k); break

    def _show_mod(self, key):
        for f in self.mod_frames.values(): f.pack_forget()
        self.mod_frames[key].pack(fill=tk.X, in_=self.dyn)
        self.current_module.set(key)

    # ============================================================
    #  VERBINDUNG
    # ============================================================
    def check_module(self):
        self.connect_btn.configure(state=tk.DISABLED)
        self.log("🔍 Prüfe ExchangeOnlineManagement...", COLORS['warning'])
        def do():
            _, out, _ = self.ps.execute(
                '$m = Get-Module -ListAvailable -Name ExchangeOnlineManagement; '
                'if ($m) { Write-Output "INSTALLED:$($m.Version)" } else { Write-Output "NOT_INSTALLED" }', 30)
            if "INSTALLED:" in out:
                v = out.split("INSTALLED:")[1].strip().split("\n")[0]
                self.root.after(0, lambda: self._mod_ok(v))
            else:
                self.root.after(0, self._mod_miss)
        threading.Thread(target=do, daemon=True).start()

    def _mod_ok(self, v):
        self.module_label.configure(text=f"✅ ExchangeOnlineManagement v{v}", fg=COLORS['success'])
        self.connect_btn.configure(state=tk.NORMAL)
        self.log(f"✅ Modul v{v} gefunden — bereit", COLORS['success'])

    def _mod_miss(self):
        self.module_label.configure(text="❌ ExchangeOnlineManagement fehlt!", fg=COLORS['error'])
        self.log("❌ Modul fehlt!", COLORS['error'])
        if messagebox.askyesno("Modul fehlt", "ExchangeOnlineManagement nicht installiert.\nJetzt installieren?"):
            self._install_mod()
        else:
            self.log("💡 Manuell: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser", COLORS['accent'])

    def _install_mod(self):
        self.log("📦 Installiere... (1-2 Min.)", COLORS['warning'])
        self.module_label.configure(text="⏳ Installiere...", fg=COLORS['warning'])
        def do():
            _, out, err = self.ps.execute("""
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
$m = Get-Module -ListAvailable -Name ExchangeOnlineManagement
if ($m) { Write-Output "INSTALL_OK:$($m.Version)" } else { Write-Output "INSTALL_FAILED" }
""", 300)
            if "INSTALL_OK:" in out:
                v = out.split("INSTALL_OK:")[1].strip().split("\n")[0]
                self.root.after(0, lambda: self._mod_ok(v))
                self.root.after(0, lambda: messagebox.showinfo("Erfolg", f"Modul v{v} installiert!"))
            else:
                self.root.after(0, lambda: self.log(f"❌ Installation fehlgeschlagen: {err}", COLORS['error']))
        threading.Thread(target=do, daemon=True).start()

    def connect(self):
        admin = self.admin_entry.get().strip()
        if not admin: messagebox.showwarning("Fehlt", "Bitte Admin-UPN eingeben!"); return
        self.log("🔄 Verbinde zu Exchange Online...", COLORS['warning'])
        self.connect_btn.configure(state=tk.DISABLED)
        def do():
            ok, _, err = self.ps.execute(
                f'Connect-ExchangeOnline -UserPrincipalName "{admin}" -ShowBanner:$false', 180)
            if ok:
                v, vo, _ = self.ps.execute('Get-OrganizationConfig | Select-Object -ExpandProperty Name', 30)
                org = vo.strip().split("\n")[0] if v and vo.strip() else ""
                self.root.after(0, lambda: self._connected(org))
            else:
                self.root.after(0, lambda: self._conn_fail(err))
        threading.Thread(target=do, daemon=True).start()

    def _connected(self, org):
        self.connected = True
        self.connect_btn.configure(text="✅ Verbunden", bg=COLORS['success'])
        n = f" mit {org}" if org else ""
        self.log(f"✅ Verbunden{n}!", COLORS['success'])
        messagebox.showinfo("Verbunden", f"Erfolgreich verbunden{n}!\nDaten werden geladen...")
        self._load_data()

    def _conn_fail(self, err):
        self.connect_btn.configure(state=tk.NORMAL)
        self.log(f"❌ Verbindung fehlgeschlagen: {err}", COLORS['error'])
        messagebox.showerror("Fehler", f"Verbindung fehlgeschlagen:\n\n{err}")

    def disconnect(self):
        self.log("🔌 Trenne...", COLORS['warning'])
        self.ps.execute("Disconnect-ExchangeOnline -Confirm:$false", 30)
        self.connected = False; self.all_mailboxes = []
        for k in self.all_groups: self.all_groups[k] = []
        self.connect_btn.configure(text="🔌 Verbinden", bg=COLORS['accent'])
        self.connect_btn.configure(state=tk.NORMAL)
        for cb in [self.mb_target, self.mb_user, self.ob_user, self.ob_fwd]:
            cb.configure(values=[], state="disabled"); cb.set("-- Erst verbinden --")
        for k in ['teams','verteiler','security']:
            for a in ['gc','uc']:
                c = getattr(self, f'{k}_{a}'); c.configure(values=[], state="disabled"); c.set("-- Erst verbinden --")
        self.log("✅ Getrennt.", COLORS['success'])

    # ============================================================
    #  DATEN LADEN
    # ============================================================
    def _load_data(self):
        self.log("📥 Lade Postfächer & Gruppen...", COLORS['warning'])
        def do():
            r1, o1, _ = self.ps.execute('Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails | ConvertTo-Json -Compress', 180)
            r2, o2, _ = self.ps.execute('Get-UnifiedGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, GroupType | ConvertTo-Json -Compress', 180)
            r3, o3, _ = self.ps.execute('Get-DistributionGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, GroupType | ConvertTo-Json -Compress', 180)
            self.root.after(0, lambda: self._data_done(r1,o1,r2,o2,r3,o3))
        threading.Thread(target=do, daemon=True).start()

    def _pj(self, o):
        if not o or not o.strip(): return []
        try:
            s1, s2 = o.find('['), o.find('{')
            if s1 == -1 and s2 == -1: return []
            s = s1 if s1 != -1 and (s2 == -1 or s1 < s2) else s2
            d = json.loads(o[s:])
            return [d] if isinstance(d, dict) else d
        except: return []

    def _data_done(self, r1, o1, r2, o2, r3, o3):
        if r1:
            self.all_mailboxes = []
            for mb in self._pj(o1):
                e, n, t = mb.get('PrimarySmtpAddress',''), mb.get('DisplayName',''), mb.get('RecipientTypeDetails','')
                if e:
                    sh = 'Shared' in t
                    self.all_mailboxes.append({'display': f"{'👥' if sh else '👤'} {n} <{e}>",
                                               'email': e, 'name': n, 'type': 'shared' if sh else 'user'})
            self.all_mailboxes.sort(key=lambda x: x['display'])
            self._apply_mb_filter()
            self._update_ob_combos()
            self.log(f"  📧 {len(self.all_mailboxes)} Postfächer", COLORS['text_dim'])
        if r2:
            self.all_groups['teams'] = []
            for g in self._pj(o2):
                e, n = g.get('PrimarySmtpAddress',''), g.get('DisplayName','')
                if e: self.all_groups['teams'].append({'display': f"👥 {n} <{e}>", 'email': e, 'name': n})
            self.all_groups['teams'].sort(key=lambda x: x['display'])
            self._update_grp('teams')
            self.log(f"  👥 {len(self.all_groups['teams'])} Teams/M365 Gruppen", COLORS['text_dim'])
        if r3:
            self.all_groups['verteiler'] = []; self.all_groups['security'] = []
            for g in self._pj(o3):
                e, n, gt = g.get('PrimarySmtpAddress',''), g.get('DisplayName',''), str(g.get('GroupType',''))
                if e:
                    ent = {'display': f"{'🔒' if 'Security' in gt else '📨'} {n} <{e}>", 'email': e, 'name': n}
                    (self.all_groups['security'] if 'Security' in gt else self.all_groups['verteiler']).append(ent)
            for k in ['verteiler','security']:
                self.all_groups[k].sort(key=lambda x: x['display']); self._update_grp(k)
            self.log(f"  📨 {len(self.all_groups['verteiler'])} Verteilerlisten", COLORS['text_dim'])
            self.log(f"  🔒 {len(self.all_groups['security'])} Sicherheitsgruppen", COLORS['text_dim'])
        t = len(self.all_mailboxes) + sum(len(v) for v in self.all_groups.values())
        self.log(f"✅ {t} Objekte geladen!", COLORS['success'])

    def _update_grp(self, k):
        items = [g['display'] for g in self.all_groups[k]]
        getattr(self, f'{k}_gc').configure(values=items, state="normal"); getattr(self, f'{k}_gc').set("")
        ui = [mb['display'] for mb in self.all_mailboxes]
        getattr(self, f'{k}_uc').configure(values=ui, state="normal"); getattr(self, f'{k}_uc').set("")

    def _update_ob_combos(self):
        items = [mb['display'] for mb in self.all_mailboxes]
        self.ob_user.configure(values=items, state="normal"); self.ob_user.set("")
        self.ob_fwd.configure(values=items, state="normal"); self.ob_fwd.set("")

    # ============================================================
    #  POSTFACH-LOGIK
    # ============================================================
    def _apply_mb_filter(self):
        ft = self.mb_filter.get()
        fl = self.all_mailboxes if ft == "all" else [m for m in self.all_mailboxes if m['type'] == ft]
        i = [m['display'] for m in fl]
        self.mb_target.configure(values=i, state="normal"); self.mb_user.configure(values=i, state="normal")

    def _filter_mb(self, e=None):
        s = self.mb_search.get().lower(); ft = self.mb_filter.get()
        base = self.all_mailboxes if ft == "all" else [m for m in self.all_mailboxes if m['type'] == ft]
        fl = [m for m in base if s in m['display'].lower()] if s else base
        i = [m['display'] for m in fl]
        self.mb_target.configure(values=i); self.mb_user.configure(values=i)

    def _filter_grp(self, k):
        s = getattr(self, f'{k}_se').get().lower()
        fl = [g for g in self.all_groups[k] if s in g['display'].lower()] if s else self.all_groups[k]
        getattr(self, f'{k}_gc').configure(values=[g['display'] for g in fl])

    def _filter_ob(self, e=None):
        s = self.ob_search.get().lower()
        fl = [m for m in self.all_mailboxes if s in m['display'].lower()] if s else self.all_mailboxes
        self.ob_user.configure(values=[m['display'] for m in fl])

    def _val_mb(self):
        for cb, n in [(self.mb_target, "Ziel-Postfach"), (self.mb_user, "Benutzer")]:
            if not cb.get().strip() or cb.get().startswith("--"):
                messagebox.showwarning("Fehlt", f"Bitte {n} auswählen!"); return False
        if not self.fullaccess_var.get() and not self.sendas_var.get():
            messagebox.showwarning("Fehlt", "Mindestens eine Berechtigung auswählen!"); return False
        return True

    def _add_mb(self):
        if not self._val_mb(): return
        mb, us = self._get_email(self.mb_target.get()), self._get_email(self.mb_user.get())
        msg = f"Berechtigungen HINZUFÜGEN?\n\n📬 {mb}\n👤 {us}\n\n"
        msg += f"Vollzugriff: {'✅' if self.fullaccess_var.get() else '❌'}\n"
        if self.fullaccess_var.get(): msg += f"AutoMapping: {'✅' if self.automapping_var.get() else '❌'}\n"
        msg += f"Senden als: {'✅' if self.sendas_var.get() else '❌'}"
        if not messagebox.askyesno("Bestätigen", msg): return
        def do():
            errs = []
            if self.fullaccess_var.get():
                self.root.after(0, lambda: self.log("  📂 Vollzugriff...", COLORS['warning']))
                am = "$true" if self.automapping_var.get() else "$false"
                ok, _, e = self.ps.execute(f'Add-MailboxPermission -Identity "{mb}" -User "{us}" -AccessRights FullAccess -AutoMapping {am}')
                self.root.after(0, lambda: self.log("  ✅ Vollzugriff hinzugefügt" if ok else f"  ❌ {e}", COLORS['success'] if ok else COLORS['error']))
                if not ok: errs.append(e)
            if self.sendas_var.get():
                self.root.after(0, lambda: self.log("  ✉️ Senden als...", COLORS['warning']))
                ok, _, e = self.ps.execute(f'Add-RecipientPermission -Identity "{mb}" -Trustee "{us}" -AccessRights SendAs -Confirm:$false')
                self.root.after(0, lambda: self.log("  ✅ Senden als hinzugefügt" if ok else f"  ❌ {e}", COLORS['success'] if ok else COLORS['error']))
                if not ok: errs.append(e)
            self.root.after(0, lambda: self._done(errs, "hinzugefügt"))
        threading.Thread(target=do, daemon=True).start()

    def _rem_mb(self):
        if not self._val_mb(): return
        mb, us = self._get_email(self.mb_target.get()), self._get_email(self.mb_user.get())
        if not messagebox.askyesno("⚠️", f"Berechtigungen ENTFERNEN?\n\n📬 {mb}\n👤 {us}", icon="warning"): return
        def do():
            errs = []
            if self.fullaccess_var.get():
                ok, _, e = self.ps.execute(f'Remove-MailboxPermission -Identity "{mb}" -User "{us}" -AccessRights FullAccess -Confirm:$false')
                self.root.after(0, lambda: self.log(f"  {'✅ Vollzugriff entfernt' if ok else '❌ '+e}", COLORS['success'] if ok else COLORS['error']))
                if not ok: errs.append(e)
            if self.sendas_var.get():
                ok, _, e = self.ps.execute(f'Remove-RecipientPermission -Identity "{mb}" -Trustee "{us}" -AccessRights SendAs -Confirm:$false')
                self.root.after(0, lambda: self.log(f"  {'✅ Senden als entfernt' if ok else '❌ '+e}", COLORS['success'] if ok else COLORS['error']))
                if not ok: errs.append(e)
            self.root.after(0, lambda: self._done(errs, "entfernt"))
        threading.Thread(target=do, daemon=True).start()

    # ============================================================
    #  GRUPPEN-LOGIK
    # ============================================================
    def _val_grp(self, k):
        for a, n in [('gc','Gruppe'),('uc','Benutzer')]:
            v = getattr(self, f'{k}_{a}').get().strip()
            if not v or v.startswith("--"):
                messagebox.showwarning("Fehlt", f"Bitte {n} auswählen!"); return False
        return True

    def _add_grp(self, k):
        if not self._val_grp(k): return
        g, u = self._get_email(getattr(self,f'{k}_gc').get()), self._get_email(getattr(self,f'{k}_uc').get())
        r = getattr(self, f'{k}_rv').get()
        if not messagebox.askyesno("Bestätigen", f"Hinzufügen?\n\n📋 {g}\n👤 {u}" +
                                   (f"\n🏷️ {r}" if k=='teams' else "")): return
        def do():
            self.root.after(0, lambda: self.log(f"  ➕ Füge {u} hinzu...", COLORS['warning']))
            if k == 'teams':
                cmd = f'Add-UnifiedGroupLinks -Identity "{g}" -LinkType {"Owners" if r=="Owner" else "Members"} -Links "{u}"'
            else:
                cmd = f'Add-DistributionGroupMember -Identity "{g}" -Member "{u}"'
            ok, _, e = self.ps.execute(cmd)
            self.root.after(0, lambda: self.log(f"  {'✅' if ok else '❌'} {u} {'hinzugefügt' if ok else e}", COLORS['success'] if ok else COLORS['error']))
            self.root.after(0, lambda: self._done([] if ok else [e], "hinzugefügt"))
        threading.Thread(target=do, daemon=True).start()

    def _rem_grp(self, k):
        if not self._val_grp(k): return
        g, u = self._get_email(getattr(self,f'{k}_gc').get()), self._get_email(getattr(self,f'{k}_uc').get())
        if not messagebox.askyesno("⚠️", f"Entfernen?\n\n📋 {g}\n👤 {u}", icon="warning"): return
        def do():
            self.root.after(0, lambda: self.log(f"  ➖ Entferne {u}...", COLORS['warning']))
            if k == 'teams':
                ok, _, e = self.ps.execute(f'Remove-UnifiedGroupLinks -Identity "{g}" -LinkType Members -Links "{u}" -Confirm:$false')
                self.ps.execute(f'Remove-UnifiedGroupLinks -Identity "{g}" -LinkType Owners -Links "{u}" -Confirm:$false')
            else:
                ok, _, e = self.ps.execute(f'Remove-DistributionGroupMember -Identity "{g}" -Member "{u}" -Confirm:$false')
            self.root.after(0, lambda: self.log(f"  {'✅' if ok else '❌'} {u} {'entfernt' if ok else e}", COLORS['success'] if ok else COLORS['error']))
            self.root.after(0, lambda: self._done([] if ok else [e], "entfernt"))
        threading.Thread(target=do, daemon=True).start()

    def _show_members(self, k):
        g = getattr(self, f'{k}_gc').get().strip()
        if not g or g.startswith("--"): messagebox.showwarning("Fehlt", "Bitte Gruppe auswählen!"); return
        ge = self._get_email(g)
        self.log(f"  📋 Lade Mitglieder von {ge}...", COLORS['warning'])
        def do():
            if k == 'teams':
                _, om, _ = self.ps.execute(f'Get-UnifiedGroupLinks -Identity "{ge}" -LinkType Members | Select-Object Name, PrimarySmtpAddress | ConvertTo-Json -Compress', 60)
                _, oo, _ = self.ps.execute(f'Get-UnifiedGroupLinks -Identity "{ge}" -LinkType Owners | Select-Object Name, PrimarySmtpAddress | ConvertTo-Json -Compress', 60)
                ms, ow = self._pj(om), self._pj(oo)
                ln = [f"=== {ge} ===", ""]
                if ow: ln += [f"👑 Besitzer ({len(ow)}):"] + [f"  • {o.get('Name','')} <{o.get('PrimarySmtpAddress','')}>" for o in ow] + [""]
                ln += [f"👤 Mitglieder ({len(ms)}):"] + [f"  • {m.get('Name','')} <{m.get('PrimarySmtpAddress','')}>" for m in ms]
                c = len(ms)
            else:
                _, o, _ = self.ps.execute(f'Get-DistributionGroupMember -Identity "{ge}" -ResultSize Unlimited | Select-Object Name, PrimarySmtpAddress | ConvertTo-Json -Compress', 60)
                ms = self._pj(o)
                ln = [f"=== {ge} ===","",f"👤 Mitglieder ({len(ms)}):"] + [f"  • {m.get('Name','')} <{m.get('PrimarySmtpAddress','')}>" for m in ms]
                c = len(ms)
            self.root.after(0, lambda: self._members_done(k, "\n".join(ln), c))
        threading.Thread(target=do, daemon=True).start()

    def _members_done(self, k, txt, c):
        mt = getattr(self, f'{k}_mt')
        mt.configure(state=tk.NORMAL); mt.delete('1.0', tk.END); mt.insert(tk.END, txt); mt.configure(state=tk.DISABLED)
        self.log(f"  ✅ {c} Mitglieder geladen", COLORS['success'])

    # ============================================================
    #  OFFBOARDING-LOGIK
    # ============================================================
    def _run_ob(self):
        us = self.ob_user.get().strip()
        if not us or us.startswith("--"):
            messagebox.showwarning("Fehlt", "Bitte Benutzer auswählen!"); return
        active = {k: v.get() for k, v in self.ob_vars.items()}
        if not any(active.values()):
            messagebox.showwarning("Fehlt", "Mindestens einen Schritt auswählen!"); return

        ue = self._get_email(us)
        un = us.split('<')[0].strip().lstrip('👤👥 ')

        names = [OB_STEP_NAMES[k] for k, v in active.items() if v]
        msg = f"⚠️ OFFBOARDING für:\n\n👤 {un}\n📧 {ue}\n\nSchritte:\n"
        for n in names: msg += f"  • {n}\n"
        msg += "\n⚠️ Teilweise IRREVERSIBEL! Fortfahren?"
        if not messagebox.askyesno("⚠️ Offboarding bestätigen", msg, icon="warning"): return
        if not messagebox.askyesno("🔴 Letzte Warnung",
                                   f"WIRKLICH Offboarding für {ue}?\nNICHT rückgängig!", icon="warning"): return

        self.ob_run_btn.configure(state=tk.DISABLED)
        self.ob_export_btn.configure(state=tk.DISABLED)
        self.ob_result.configure(state=tk.NORMAL); self.ob_result.delete('1.0', tk.END); self.ob_result.configure(state=tk.DISABLED)

        self.ob_report_lines = [
            "=" * 60, "OFFBOARDING-BERICHT",
            f"Datum: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
            f"Benutzer: {un}", f"E-Mail: {ue}",
            f"Durchgeführt von: {self.admin_entry.get().strip()}",
            "=" * 60, ""
        ]

        def do():
            todo = [(k, v) for k, v in active.items() if v]
            total = len(todo)
            results = {}
            for i, (sk, _) in enumerate(todo):
                pct = int((i / total) * 100)
                sn = OB_STEP_NAMES[sk]
                self.root.after(0, lambda p=pct, s=sn: self._ob_prog_upd(p, s))
                self.root.after(0, lambda s=sn: self.log(f"  🔄 {s}...", COLORS['warning']))

                ok, detail = self._ob_exec(sk, ue)
                results[sk] = (ok, detail)
                st = "✅" if ok else "❌"
                self.ob_report_lines += [f"[{st}] {sn}", f"    {detail}", ""]
                self.root.after(0, lambda s=sn, st2=st, c=(COLORS['success'] if ok else COLORS['error']): self.log(f"  {st2} {s}", c))

            sc = sum(1 for ok, _ in results.values() if ok)
            fc = sum(1 for ok, _ in results.values() if not ok)
            self.ob_report_lines += ["=" * 60, f"ERGEBNIS: {sc} erfolgreich, {fc} fehlgeschlagen", "=" * 60]
            self.root.after(0, lambda: self._ob_done(results, ue, sc, fc))
        threading.Thread(target=do, daemon=True).start()

    def _ob_prog_upd(self, pct, step):
        self.ob_prog['value'] = pct
        self.ob_prog_lbl.configure(text=f"⏳ {step}... ({pct}%)")

    def _ob_exec(self, step, ue):
        """Einzelnen Offboarding-Schritt ausführen"""
        try:
            if step == 'sign_in':
                # Anmeldung blockieren via Set-User
                ok, _, e = self.ps.execute(f'Set-User -Identity "{ue}" -AccountDisabled $true', 60)
                return ok, "Anmeldung blockiert" if ok else f"Fehler: {e}"

            elif step == 'reset_pw':
                # Zufälliges 24-Zeichen Passwort
                ok, o, e = self.ps.execute(
                    f'$chars = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%&*"; '
                    f'$pw = -join (1..24 | ForEach-Object {{ $chars[(Get-Random -Maximum $chars.Length)] }}); '
                    f'Set-Mailbox -Identity "{ue}" -Password (ConvertTo-SecureString -String $pw -AsPlainText -Force) -ErrorAction Stop; '
                    f'Write-Output "PW_OK:$pw"', 60)
                if "PW_OK:" in o:
                    return True, f"Passwort zurückgesetzt (24 Zeichen zufällig)"
                return False, f"Fehler: {e}"

            elif step == 'remove_groups':
                removed, failed = 0, 0

                # Unified Groups (Teams/M365)
                ok, o, _ = self.ps.execute(
                    f'Get-UnifiedGroup -ResultSize Unlimited | Where-Object {{ '
                    f'(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members -ErrorAction SilentlyContinue | '
                    f'Where-Object {{$_.PrimarySmtpAddress -eq "{ue}"}}) -or '
                    f'(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Owners -ErrorAction SilentlyContinue | '
                    f'Where-Object {{$_.PrimarySmtpAddress -eq "{ue}"}}) }} | '
                    f'Select-Object -ExpandProperty PrimarySmtpAddress', 180)
                if ok and o.strip():
                    for grp in o.strip().split("\n"):
                        grp = grp.strip()
                        if not grp: continue
                        r1, _, _ = self.ps.execute(f'Remove-UnifiedGroupLinks -Identity "{grp}" -LinkType Members -Links "{ue}" -Confirm:$false -ErrorAction SilentlyContinue')
                        self.ps.execute(f'Remove-UnifiedGroupLinks -Identity "{grp}" -LinkType Owners -Links "{ue}" -Confirm:$false -ErrorAction SilentlyContinue')
                        if r1: removed += 1
                        else: failed += 1

                # Distribution Groups (Verteiler + Security)
                ok, o, _ = self.ps.execute(
                    f'Get-DistributionGroup -ResultSize Unlimited | Where-Object {{ '
                    f'(Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue | '
                    f'Where-Object {{$_.PrimarySmtpAddress -eq "{ue}"}}) }} | '
                    f'Select-Object -ExpandProperty PrimarySmtpAddress', 180)
                if ok and o.strip():
                    for grp in o.strip().split("\n"):
                        grp = grp.strip()
                        if not grp: continue
                        r1, _, _ = self.ps.execute(f'Remove-DistributionGroupMember -Identity "{grp}" -Member "{ue}" -Confirm:$false -ErrorAction SilentlyContinue')
                        if r1: removed += 1
                        else: failed += 1

                detail = f"Aus {removed} Gruppe(n) entfernt"
                if failed: detail += f", {failed} fehlgeschlagen"
                return failed == 0, detail

            elif step == 'remove_licenses':
                # Versuche Microsoft.Graph Modul
                ok, o, e = self.ps.execute(
                    f'$skus = (Get-MgUserLicenseDetail -UserId "{ue}" -ErrorAction SilentlyContinue).SkuId; '
                    f'if ($skus) {{ '
                    f'  foreach ($sku in $skus) {{ Set-MgUserLicense -UserId "{ue}" -RemoveLicenses @($sku) -AddLicenses @() -ErrorAction Stop }}; '
                    f'  Write-Output "LIC_REMOVED:$($skus.Count)" '
                    f'}} else {{ Write-Output "NO_GRAPH" }}', 90)
                if "LIC_REMOVED:" in o:
                    cnt = o.split("LIC_REMOVED:")[1].strip().split("\n")[0]
                    return True, f"{cnt} Lizenz(en) entfernt"
                elif "NO_GRAPH" in o:
                    return False, "Microsoft.Graph Modul nicht verfügbar — Lizenzen bitte manuell im Admin Center entziehen"
                return False, f"Fehler: {e}"

            elif step == 'convert_shared':
                ok, _, e = self.ps.execute(f'Set-Mailbox -Identity "{ue}" -Type Shared', 60)
                return ok, "In Shared Mailbox konvertiert" if ok else f"Fehler: {e}"

            elif step == 'set_ooo':
                msg = self.ob_ooo.get('1.0', tk.END).strip()
                if not msg: return True, "Übersprungen (keine Nachricht)"
                esc = msg.replace("'", "''").replace('"', '`"')
                ok, _, e = self.ps.execute(
                    f'Set-MailboxAutoReplyConfiguration -Identity "{ue}" '
                    f'-AutoReplyState Enabled -InternalMessage "{esc}" '
                    f'-ExternalMessage "{esc}" -ExternalAudience All', 60)
                return ok, "Abwesenheitsnachricht aktiviert" if ok else f"Fehler: {e}"

            elif step == 'forwarding':
                fwd = self.ob_fwd.get().strip()
                if not fwd: return True, "Übersprungen (kein Ziel angegeben)"
                fe = self._get_email(fwd) if '<' in fwd else fwd
                ok, _, e = self.ps.execute(
                    f'Set-Mailbox -Identity "{ue}" -ForwardingSmtpAddress "smtp:{fe}" -DeliverToMailboxAndForward $true', 60)
                return ok, f"Weiterleitung an {fe}" if ok else f"Fehler: {e}"

            elif step == 'hide_gal':
                ok, _, e = self.ps.execute(f'Set-Mailbox -Identity "{ue}" -HiddenFromAddressListsEnabled $true', 60)
                return ok, "Aus Adressbuch ausgeblendet" if ok else f"Fehler: {e}"

            elif step == 'disable_sync':
                ok, _, e = self.ps.execute(
                    f'Set-CASMailbox -Identity "{ue}" -ActiveSyncEnabled $false '
                    f'-OWAEnabled $false -PopEnabled $false -ImapEnabled $false '
                    f'-MAPIEnabled $false -EwsEnabled $false', 60)
                return ok, "Alle Protokolle deaktiviert (ActiveSync, OWA, POP, IMAP, MAPI, EWS)" if ok else f"Fehler: {e}"

            elif step == 'remove_delegates':
                ok, o, e = self.ps.execute(
                    f'$perms = Get-MailboxPermission -Identity "{ue}" | '
                    f'Where-Object {{$_.User -ne "NT AUTHORITY\\SELF" -and $_.IsInherited -eq $false}}; '
                    f'$count = 0; '
                    f'foreach ($p in $perms) {{ Remove-MailboxPermission -Identity "{ue}" -User $p.User -AccessRights $p.AccessRights -Confirm:$false -ErrorAction SilentlyContinue; $count++ }}; '
                    f'Write-Output "DEL_DONE:$count"', 90)
                if "DEL_DONE:" in o:
                    cnt = o.split("DEL_DONE:")[1].strip().split("\n")[0]
                    return True, f"{cnt} Delegierung(en) entfernt"
                return False, f"Fehler: {e}"

            return False, f"Unbekannter Schritt: {step}"
        except Exception as ex:
            return False, f"Exception: {ex}"

    def _ob_done(self, results, ue, sc, fc):
        self.ob_prog['value'] = 100
        self.ob_prog_lbl.configure(text="✅ Offboarding abgeschlossen")
        self.ob_run_btn.configure(state=tk.NORMAL)
        self.ob_export_btn.configure(state=tk.NORMAL)

        self.ob_result.configure(state=tk.NORMAL); self.ob_result.delete('1.0', tk.END)
        for line in self.ob_report_lines:
            self.ob_result.insert(tk.END, line + "\n")
        self.ob_result.configure(state=tk.DISABLED)

        self.log(f"🚪 Offboarding {ue}: {sc} ✅ / {fc} ❌",
                 COLORS['success'] if fc == 0 else COLORS['warning'])
        if fc > 0:
            messagebox.showwarning("Teilweise fehlgeschlagen",
                                   f"✅ {sc} erfolgreich\n❌ {fc} fehlgeschlagen\n\nSiehe Bericht.")
        else:
            messagebox.showinfo("Offboarding erfolgreich",
                               f"Alle {sc} Schritte erfolgreich!\nBericht kann exportiert werden.")

    def _export_ob(self):
        if not self.ob_report_lines:
            messagebox.showwarning("Kein Bericht", "Noch kein Bericht vorhanden."); return
        us = self.ob_user.get().strip()
        ue = self._get_email(us) if '<' in us else "benutzer"
        un = ue.split('@')[0] if '@' in ue else ue
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fp = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Textdatei","*.txt"),("Alle","*.*")],
            initialfile=f"Offboarding_{un}_{ts}.txt", title="Offboarding-Bericht speichern")
        if fp:
            try:
                with open(fp, 'w', encoding='utf-8') as f:
                    f.write("\n".join(self.ob_report_lines))
                self.log(f"💾 Bericht gespeichert: {fp}", COLORS['success'])
                messagebox.showinfo("Gespeichert", f"Bericht gespeichert:\n{fp}")
            except Exception as ex:
                self.log(f"❌ Fehler: {ex}", COLORS['error'])
                messagebox.showerror("Fehler", f"Speichern fehlgeschlagen:\n{ex}")

    # ============================================================
    #  ALLGEMEIN
    # ============================================================
    def _done(self, errs, word):
        if errs:
            for e in errs: self.log(f"  ❌ {e}", COLORS['error'])
            messagebox.showerror("Fehler", "Aktion teilweise fehlgeschlagen.\nSiehe Protokoll.")
        else:
            messagebox.showinfo("Erfolg", f"✅ Erfolgreich {word}!")

    def log(self, msg, color=None):
        ts = datetime.now().strftime("%H:%M:%S")
        tag = {COLORS['success']:'success', COLORS['error']:'error', COLORS['warning']:'warning',
               COLORS['accent']:'accent', COLORS['text_dim']:'dim'}.get(color, 'info')
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] ", 'info')
        self.log_text.insert(tk.END, f"{msg}\n", tag)
        self.log_text.see(tk.END); self.log_text.configure(state=tk.DISABLED)

    def cleanup(self):
        try: self.ps.execute("Disconnect-ExchangeOnline -Confirm:$false", 10)
        except: pass
        self.ps.stop()


def main():
    root = tk.Tk()
    try: root.iconbitmap('exchange.ico')
    except: pass
    app = M365AdminTool(root)
    def on_close():
        if messagebox.askokcancel("Beenden", "Beenden?\nExchange Online Session wird getrennt."):
            app.cleanup(); root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    main()