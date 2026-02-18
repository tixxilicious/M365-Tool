#!/usr/bin/env python3
"""
Microsoft 365 Admin Tool v5.0
Erstellt für Kaulich IT Systems GmbH

Module:
 1. Postfachberechtigungen        2. Teams/M365 Gruppen
 3. Verteilerlisten                4. Sicherheitsgruppen
 5. Benutzer-Offboarding           6. Lizenz-Übersicht
 7. Benutzer-Info                  8. Shared Mailbox erstellen
 9. Mail-Weiterleitungen          10. Berechtigungs-Audit
11. CSV-Export                    12. Bulk-Aktionen

Voraussetzungen: Windows, PowerShell 5.1+, ExchangeOnlineManagement
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess, threading, json, queue, time, os, csv
from datetime import datetime

# ── Farben ──────────────────────────────────────────────────
C = {
    'bg':'#1e1e2e','panel':'#252536','input':'#2e2e42','hdr':'#16161e',
    'accent':'#7aa2f7','ok':'#9ece6a','warn':'#e0af68','err':'#f7768e',
    'txt':'#c0caf5','dim':'#565f89','muted':'#414868','brd':'#3b3b54',
    'purple':'#bb9af7','cyan':'#7dcfff','orange':'#ff9e64','green':'#73daca',
}

MODULES = [
    ('postfach',    '📧 Postfachberechtigungen'),
    ('teams',       '👥 Teams/M365 Gruppen'),
    ('verteiler',   '📨 Verteilerlisten'),
    ('security',    '🔒 Sicherheitsgruppen'),
    ('offboarding', '🚪 Benutzer-Offboarding'),
    ('userinfo',    '👤 Benutzer-Info'),
    ('licenses',    '📊 Lizenz-Übersicht'),
    ('sharedmb',    '📧 Shared Mailbox erstellen'),
    ('forwarding',  '📬 Mail-Weiterleitungen'),
    ('audit',       '🔍 Berechtigungs-Audit'),
    ('csvexport',   '📋 CSV-Export'),
    ('bulk',        '🏷️ Bulk-Aktionen'),
]
MOD_KEYS = [k for k, _ in MODULES]
MOD_LABELS = [v for _, v in MODULES]

OB_STEPS = [
    ('sign_in','🔒 Anmeldung blockieren'),('reset_pw','🔑 Passwort zurücksetzen'),
    ('remove_groups','👥 Aus allen Gruppen entfernen'),('remove_licenses','📋 Lizenzen entziehen'),
    ('convert_shared','📧 In Shared Mailbox konvertieren'),('set_ooo','✈️ Abwesenheit setzen'),
    ('fwd','📨 Mail-Weiterleitung'),('hide_gal','👻 Aus Adressbuch ausblenden'),
    ('disable_sync','📱 Protokolle deaktivieren'),('remove_delegates','🔓 Delegierungen entfernen'),
]
OB_NAMES = {k: v.split(' ',1)[1] for k,v in OB_STEPS}


# ── PowerShell Session ──────────────────────────────────────
class PS:
    def __init__(self):
        self.proc = None; self.q = queue.Queue()
    def start(self):
        if self.proc: return
        self.proc = subprocess.Popen(["powershell","-NoLogo","-NoExit","-Command","-"],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace', bufsize=1)
        threading.Thread(target=self._rd, daemon=True).start()
        self._w('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8')
    def _rd(self):
        while self.proc and self.proc.poll() is None:
            try:
                l = self.proc.stdout.readline()
                if l: self.q.put(l)
            except: break
    def _w(self, c):
        if self.proc and self.proc.poll() is None:
            self.proc.stdin.write(c+"\n"); self.proc.stdin.flush()
    def run(self, cmd, timeout=120):
        if not self.proc or self.proc.poll() is not None:
            return False, "", "PS nicht aktiv"
        while not self.q.empty():
            try: self.q.get_nowait()
            except: break
        ts = int(time.time()*1000)
        em, erm = f"###END_{ts}###", f"###ERR_{ts}###"
        self._w(f'try {{ {cmd} }} catch {{ Write-Output "{erm}$($_.Exception.Message)" }}\nWrite-Output "{em}"')
        ol, el = [], []
        t0 = time.time()
        while True:
            if time.time()-t0 > timeout: return False, "", "Timeout"
            try:
                l = self.q.get(timeout=0.5).rstrip()
                if em in l: break
                elif erm in l: el.append(l.split(erm,1)[1])
                else: ol.append(l)
            except queue.Empty: continue
        return (not el), "\n".join(ol), "\n".join(el)
    def stop(self):
        if self.proc:
            try: self._w("exit"); self.proc.terminate()
            except: pass
            self.proc = None


# ── Styled Button (Python 3.14 safe) ───────────────────────
class Btn(tk.Canvas):
    def __init__(self, parent, text, command=None, bg=C['accent'],
                 fg='#fff', width=160, height=34, font_size=10, **kw):
        kw.pop('width',None); kw.pop('height',None)
        pbg = C['panel']
        try: pbg = parent.cget('bg')
        except: pass
        super().__init__(parent, highlightthickness=0, borderwidth=0, **kw)
        super().configure(width=width, height=height, bg=pbg)
        self._cmd = command; self._bg = bg; self._fg = fg
        self._hov = self._adj(bg, 25); self._dis = C['muted']
        self.txt = text; self._bw = width; self._bh = height
        self._fs = font_size; self._en = True
        self.after(10, lambda: self._draw(self._bg))
        self.bind('<Enter>', lambda e: self._draw(self._hov) if self._en else None)
        self.bind('<Leave>', lambda e: self._draw(self._bg if self._en else self._dis))
        self.bind('<Button-1>', lambda e: self._cmd() if self._en and self._cmd else None)

    def _adj(self, c, a):
        try:
            r,g,b = int(c[1:3],16),int(c[3:5],16),int(c[5:7],16)
            return f'#{min(255,r+a):02x}{min(255,g+a):02x}{min(255,b+a):02x}'
        except: return c

    def _draw(self, col):
        self.delete('all')
        r,w,h = 6, self._bw, self._bh
        for x0,y0,x1,y1,st in [(0,0,r*2,r*2,90),(w-r*2,0,w,r*2,0),
                                 (0,h-r*2,r*2,h,180),(w-r*2,h-r*2,w,h,270)]:
            self.create_arc(x0,y0,x1,y1,start=st,extent=90,fill=col,outline=col)
        self.create_rectangle(r,0,w-r,h,fill=col,outline=col)
        self.create_rectangle(0,r,w,h-r,fill=col,outline=col)
        tc = self._fg if self._en else C['dim']
        self.create_text(w/2,h/2,text=self.txt,fill=tc,font=('Segoe UI',self._fs,'bold'))

    def configure(self, **kw):
        if 'state' in kw:
            self._en = kw['state'] != tk.DISABLED; self._draw(self._bg if self._en else self._dis)
        if 'text' in kw: self.txt = kw['text']; self._draw(self._bg if self._en else self._dis)
        if 'bg' in kw:
            self._bg = kw['bg']; self._hov = self._adj(kw['bg'],25); self._draw(self._bg)


# ── Hauptanwendung ──────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("M365 Admin Tool v5.0 — Kaulich IT Systems GmbH")
        self.root.geometry("920x920"); self.root.minsize(850,700)
        self.root.configure(bg=C['bg'])
        self.ps = PS(); self.ps.start()
        self.connected = False
        self.mailboxes = []; self.groups = {'teams':[],'verteiler':[],'security':[]}
        self.ob_report = []
        self._build()
        self.log("🚀 M365 Admin Tool v5.0 gestartet", C['ok'])
        self.root.after(500, self._chk_mod)

    # ── JSON Helper ─────────────────────────────────────────
    def _pj(self, o):
        if not o or not o.strip(): return []
        try:
            s1, s2 = o.find('['), o.find('{')
            if s1==-1 and s2==-1: return []
            s = s1 if s1!=-1 and (s2==-1 or s1<s2) else s2
            d = json.loads(o[s:]); return [d] if isinstance(d,dict) else d
        except: return []

    def _ge(self, s):
        return s.split('<')[1].split('>')[0] if '<' in s and '>' in s else s

    # ── UI Aufbau ───────────────────────────────────────────
    def _build(self):
        hdr = tk.Frame(self.root, bg=C['hdr'], height=48)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="⚡ M365 Admin Tool", font=('Segoe UI',15,'bold'),
                 fg=C['accent'], bg=C['hdr']).pack(side=tk.LEFT, padx=12)
        tk.Label(hdr, text="v5.0 — Kaulich IT Systems", font=('Segoe UI',9),
                 fg=C['dim'], bg=C['hdr']).pack(side=tk.RIGHT, padx=12)

        outer = tk.Frame(self.root, bg=C['bg']); outer.pack(fill=tk.BOTH, expand=True)
        cv = tk.Canvas(outer, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(outer, orient=tk.VERTICAL, command=cv.yview)
        self.main = tk.Frame(cv, bg=C['bg'])
        self.main.bind('<Configure>', lambda e: cv.configure(scrollregion=cv.bbox("all")))
        self._cw = cv.create_window((0,0), window=self.main, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=12, pady=8)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.bind('<Configure>', lambda e: cv.itemconfig(self._cw, width=e.width))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))

        m = self.main

        # Verbindung
        cf = self._sec(m, "🔐 Verbindung")
        ci = tk.Frame(cf, bg=C['panel']); ci.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(ci,text="Admin-UPN:",font=('Segoe UI',10),fg=C['txt'],bg=C['panel']).grid(row=0,column=0,padx=(0,8))
        self.admin_e = tk.Entry(ci, width=35, font=('Segoe UI',10), bg=C['input'], fg=C['txt'],
                                insertbackground=C['txt'], relief=tk.FLAT, highlightthickness=1,
                                highlightbackground=C['brd'], highlightcolor=C['accent'])
        self.admin_e.grid(row=0, column=1, padx=(0,8), ipady=4)
        self.conn_btn = Btn(ci, "🔌 Verbinden", command=self._connect, width=130)
        self.conn_btn.grid(row=0, column=2)
        self.mod_lbl = tk.Label(cf, text="⏳ Prüfe...", font=('Segoe UI',9), fg=C['warn'], bg=C['panel'])
        self.mod_lbl.pack(fill=tk.X, padx=10, pady=(0,6))

        # Modul-Auswahl
        mf = self._sec(m, "🔧 Funktion")
        mi = tk.Frame(mf, bg=C['panel']); mi.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(mi,text="Modul:",font=('Segoe UI',10,'bold'),fg=C['txt'],bg=C['panel']).pack(side=tk.LEFT,padx=(0,8))
        self.mod_cb = ttk.Combobox(mi, width=55, state="readonly", font=('Segoe UI',10), values=MOD_LABELS)
        self.mod_cb.pack(side=tk.LEFT); self.mod_cb.current(0)
        self.mod_cb.bind('<<ComboboxSelected>>', self._on_mod)

        # Dynamischer Bereich
        self.dyn = tk.Frame(m, bg=C['bg']); self.dyn.pack(fill=tk.X, pady=(4,0))
        self.mf = {}

        self._b_postfach(); self._b_grp('teams','Teams-Gruppe')
        self._b_grp('verteiler','Verteilerliste'); self._b_grp('security','Sicherheitsgruppe')
        self._b_offboarding(); self._b_userinfo(); self._b_licenses()
        self._b_sharedmb(); self._b_forwarding(); self._b_audit()
        self._b_csvexport(); self._b_bulk()

        # Protokoll
        lc = tk.Frame(m, bg=C['bg']); lc.pack(fill=tk.X, pady=(8,0))
        tk.Label(lc,text="📋 Protokoll",font=('Segoe UI',11,'bold'),fg=C['accent'],bg=C['bg']).pack(anchor=tk.W,pady=(0,3))
        ls = tk.Frame(lc, bg=C['panel'], highlightbackground=C['brd'], highlightthickness=1); ls.pack(fill=tk.X)
        li = tk.Frame(ls, bg=C['panel']); li.pack(fill=tk.X, padx=10, pady=8)
        self.log_t = tk.Text(li, height=7, font=('Consolas',9), bg=C['input'], fg=C['txt'],
                             relief=tk.FLAT, wrap=tk.WORD, highlightthickness=1, highlightbackground=C['brd'])
        lsb = tk.Scrollbar(li, orient=tk.VERTICAL, command=self.log_t.yview)
        self.log_t.configure(yscrollcommand=lsb.set)
        self.log_t.pack(side=tk.LEFT, fill=tk.X, expand=True); lsb.pack(side=tk.RIGHT, fill=tk.Y)
        for t,c2 in [('ok',C['ok']),('err',C['err']),('warn',C['warn']),('info',C['txt']),('acc',C['accent']),('dim',C['dim'])]:
            self.log_t.tag_configure(t, foreground=c2)

        st = ttk.Style(); st.theme_use('clam')
        st.configure('TCombobox', fieldbackground=C['input'], background=C['panel'],
                     foreground=C['txt'], arrowcolor=C['txt'], bordercolor=C['brd'])
        self._show('postfach')

    # ── UI Helpers ──────────────────────────────────────────
    def _sec(self, p, title):
        c = tk.Frame(p, bg=C['bg']); c.pack(fill=tk.X, pady=(0,6))
        tk.Label(c,text=title,font=('Segoe UI',11,'bold'),fg=C['accent'],bg=C['bg']).pack(anchor=tk.W,pady=(0,3))
        f = tk.Frame(c, bg=C['panel'], highlightbackground=C['brd'], highlightthickness=1)
        f.pack(fill=tk.X); return f

    def _sec2(self, p, title):
        c = tk.Frame(p, bg=C['bg']); c.pack(fill=tk.X, pady=(0,4))
        tk.Label(c,text=title,font=('Segoe UI',11,'bold'),fg=C['purple'],bg=C['bg']).pack(anchor=tk.W,pady=(0,3))
        f = tk.Frame(c, bg=C['panel'], highlightbackground=C['brd'], highlightthickness=1)
        f.pack(fill=tk.X); return f

    def _cb(self, p, lbl, hint=None):
        r = tk.Frame(p, bg=C['panel']); r.pack(fill=tk.X, padx=10, pady=(4,0))
        tk.Label(r,text=lbl,font=('Segoe UI',10),fg=C['txt'],bg=C['panel']).pack(side=tk.LEFT,padx=(0,8))
        cb = ttk.Combobox(r, width=60, state="disabled", font=('Segoe UI',9))
        cb.pack(side=tk.LEFT, fill=tk.X, expand=True); cb.set("-- Erst verbinden --")
        if hint:
            hr = tk.Frame(p, bg=C['panel']); hr.pack(fill=tk.X, padx=10, pady=(1,0))
            tk.Label(hr,text=hint,font=('Segoe UI',8),fg=C['dim'],bg=C['panel']).pack(anchor=tk.E)
        return cb

    def _srch(self, p, fn):
        r = tk.Frame(p, bg=C['panel']); r.pack(fill=tk.X, padx=10, pady=(6,6))
        tk.Label(r,text="🔍",font=('Segoe UI',10),fg=C['txt'],bg=C['panel']).pack(side=tk.LEFT,padx=(0,4))
        e = tk.Entry(r, width=55, font=('Segoe UI',10), bg=C['input'], fg=C['txt'],
                     insertbackground=C['txt'], relief=tk.FLAT, highlightthickness=1,
                     highlightbackground=C['brd'], highlightcolor=C['accent'])
        e.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3); e.bind('<KeyRelease>', fn); return e

    def _txt(self, p, h=5):
        t = tk.Text(p, height=h, font=('Consolas',9), bg=C['input'], fg=C['txt'], relief=tk.FLAT,
                    wrap=tk.WORD, state=tk.DISABLED, highlightthickness=1, highlightbackground=C['brd'])
        t.pack(fill=tk.X); return t

    def _btns(self, p, pairs):
        r = tk.Frame(p, bg=C['panel']); r.pack(fill=tk.X, padx=10, pady=(4,8))
        widgets = []
        for txt, cmd, bg, side, w in pairs:
            b = Btn(r, txt, command=cmd, bg=bg, width=w, fg='#1e1e2e' if bg==C['ok'] else '#fff')
            b.pack(side=side, padx=(0,8) if side==tk.LEFT else (8,0))
            widgets.append(b)
        return widgets

    # ── Modul-Wechsel ───────────────────────────────────────
    def _on_mod(self, e=None):
        idx = self.mod_cb.current()
        if 0 <= idx < len(MOD_KEYS): self._show(MOD_KEYS[idx])
    def _show(self, k):
        for f in self.mf.values(): f.pack_forget()
        if k in self.mf: self.mf[k].pack(fill=tk.X, in_=self.dyn)

    # ── 1. Postfach ─────────────────────────────────────────
    def _b_postfach(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['postfach'] = f
        s = self._sec2(f, "📧 Postfachberechtigungen")
        self.mb_t = self._cb(s, "Ziel-Postfach:", "(Postfach)")
        self.mb_u = self._cb(s, "Benutzer:      ", "(Benutzer)")
        self.mb_s = self._srch(s, self._fmb)
        fr = tk.Frame(s, bg=C['panel']); fr.pack(fill=tk.X, padx=10, pady=(0,4))
        tk.Label(fr,text="Typ:",font=('Segoe UI',9,'bold'),fg=C['txt'],bg=C['panel']).pack(side=tk.LEFT,padx=(0,6))
        self.mb_ft = tk.StringVar(value="all")
        for t,v in [("Alle","all"),("👤 User","user"),("👥 Shared","shared")]:
            tk.Radiobutton(fr,text=t,variable=self.mb_ft,value=v,font=('Segoe UI',9),fg=C['txt'],
                           bg=C['panel'],selectcolor=C['input'],activebackground=C['panel'],
                           command=self._amb).pack(side=tk.LEFT,padx=(0,10))
        pr = tk.Frame(s, bg=C['panel']); pr.pack(fill=tk.X, padx=10, pady=(4,4))
        self.fa_v = tk.BooleanVar(value=True); self.am_v = tk.BooleanVar(value=True); self.sa_v = tk.BooleanVar(value=True)
        tk.Checkbutton(pr,text="📂 Vollzugriff",variable=self.fa_v,font=('Segoe UI',10),fg=C['txt'],
                        bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(side=tk.LEFT)
        self.am_cb = tk.Checkbutton(pr,text="🔗 AutoMap",variable=self.am_v,font=('Segoe UI',10),fg=C['txt'],
                                     bg=C['panel'],selectcolor=C['input'],activebackground=C['panel'])
        self.am_cb.pack(side=tk.LEFT,padx=(15,0))
        tk.Checkbutton(pr,text="✉️ Senden als",variable=self.sa_v,font=('Segoe UI',10),fg=C['txt'],
                        bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(side=tk.LEFT,padx=(15,0))
        self._btns(s, [("✅ Hinzufügen",self._add_mb,C['ok'],tk.LEFT,170),
                        ("❌ Entfernen",self._rem_mb,C['err'],tk.LEFT,170),
                        ("🔌 Trennen",self.disconnect,C['input'],tk.RIGHT,100)])

    # ── 2-4. Gruppen ────────────────────────────────────────
    def _b_grp(self, key, label):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf[key] = f
        ic = {'teams':'👥','verteiler':'📨','security':'🔒'}[key]
        s = self._sec2(f, f"{ic} {label}")
        gc = self._cb(s, f"{label}:"); uc = self._cb(s, "Benutzer:")
        se = self._srch(s, lambda e,k=key: self._fgrp(k))
        rv = tk.StringVar(value="Member")
        if key == 'teams':
            rr = tk.Frame(s, bg=C['panel']); rr.pack(fill=tk.X, padx=10, pady=(0,4))
            tk.Label(rr,text="Rolle:",font=('Segoe UI',9,'bold'),fg=C['txt'],bg=C['panel']).pack(side=tk.LEFT,padx=(0,6))
            for t,v in [("👤 Mitglied","Member"),("👑 Besitzer","Owner")]:
                tk.Radiobutton(rr,text=t,variable=rv,value=v,font=('Segoe UI',9),fg=C['txt'],
                               bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(side=tk.LEFT,padx=(0,10))
        mfr = tk.Frame(s, bg=C['panel']); mfr.pack(fill=tk.X, padx=10, pady=(0,4))
        Btn(mfr,"📋 Mitglieder",command=lambda k=key: self._smem(k),bg=C['input'],width=140,height=26,font_size=9).pack(anchor=tk.W,pady=(2,4))
        mt = self._txt(mfr, 4)
        self._btns(s, [("✅ Hinzufügen",lambda k=key:self._agrp(k),C['ok'],tk.LEFT,150),
                        ("❌ Entfernen",lambda k=key:self._rgrp(k),C['err'],tk.LEFT,150),
                        ("🔌 Trennen",self.disconnect,C['input'],tk.RIGHT,100)])
        for a,v in [('gc',gc),('uc',uc),('se',se),('rv',rv),('mt',mt)]:
            setattr(self, f'{key}_{a}', v)

    # ── 5. Offboarding ──────────────────────────────────────
    def _b_offboarding(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['offboarding'] = f
        s = self._sec2(f, "🚪 Benutzer-Offboarding")
        self.ob_u = self._cb(s, "Benutzer:"); self.ob_s = self._srch(s, self._fob)
        wr = tk.Frame(s, bg=C['panel']); wr.pack(fill=tk.X, padx=10, pady=(0,6))
        tk.Label(wr,text="⚠️ ACHTUNG: Teilweise irreversibel!",font=('Segoe UI',9,'bold'),
                 fg=C['err'],bg=C['panel']).pack(anchor=tk.W)
        cbf = tk.Frame(s, bg=C['panel']); cbf.pack(fill=tk.X, padx=10, pady=(0,4))
        self.ob_v = {}
        for i,(k,l) in enumerate(OB_STEPS):
            v = tk.BooleanVar(value=True); self.ob_v[k] = v
            tk.Checkbutton(cbf,text=l,variable=v,font=('Segoe UI',9),fg=C['txt'],bg=C['panel'],
                           selectcolor=C['input'],activebackground=C['panel']
                           ).grid(row=i//2,column=i%2,sticky=tk.W,padx=(0,15),pady=1)
        tr = tk.Frame(s, bg=C['panel']); tr.pack(fill=tk.X, padx=10, pady=(2,4))
        Btn(tr,"☑️ Alle an",command=lambda:[v.set(True) for v in self.ob_v.values()],
            bg=C['input'],width=85,height=24,font_size=8).pack(side=tk.LEFT,padx=(0,6))
        Btn(tr,"☐ Alle ab",command=lambda:[v.set(False) for v in self.ob_v.values()],
            bg=C['input'],width=85,height=24,font_size=8).pack(side=tk.LEFT)
        of = tk.Frame(s, bg=C['panel']); of.pack(fill=tk.X, padx=10, pady=(0,4))
        tk.Label(of,text="✈️ Abwesenheit:",font=('Segoe UI',9),fg=C['dim'],bg=C['panel']).pack(anchor=tk.W)
        self.ob_ooo = tk.Text(of,height=2,font=('Segoe UI',9),bg=C['input'],fg=C['txt'],relief=tk.FLAT,
                               wrap=tk.WORD,highlightthickness=1,highlightbackground=C['brd'])
        self.ob_ooo.pack(fill=tk.X,pady=(2,0))
        self.ob_ooo.insert('1.0','Dieser Mitarbeiter ist nicht mehr im Unternehmen. Bitte wenden Sie sich an helpdesk@kaulich-it.de')
        ff = tk.Frame(s, bg=C['panel']); ff.pack(fill=tk.X, padx=10, pady=(4,4))
        tk.Label(ff,text="📨 Weiterleitung an:",font=('Segoe UI',9),fg=C['dim'],bg=C['panel']).pack(side=tk.LEFT,padx=(0,6))
        self.ob_fwd = ttk.Combobox(ff,width=45,state="disabled",font=('Segoe UI',9))
        self.ob_fwd.pack(side=tk.LEFT,fill=tk.X,expand=True)
        pf = tk.Frame(s, bg=C['panel']); pf.pack(fill=tk.X, padx=10, pady=(4,4))
        self.ob_pl = tk.Label(pf,text="",font=('Segoe UI',9),fg=C['dim'],bg=C['panel']); self.ob_pl.pack(anchor=tk.W)
        self.ob_pb = ttk.Progressbar(pf,mode='determinate'); self.ob_pb.pack(fill=tk.X,pady=(2,0))
        rf = tk.Frame(s, bg=C['panel']); rf.pack(fill=tk.X, padx=10, pady=(4,4))
        self.ob_rt = self._txt(rf, 5)
        br = tk.Frame(s, bg=C['panel']); br.pack(fill=tk.X, padx=10, pady=(4,8))
        self.ob_rb = Btn(br,"🚪 Offboarding starten",command=self._run_ob,bg=C['err'],width=190)
        self.ob_rb.pack(side=tk.LEFT,padx=(0,8))
        self.ob_eb = Btn(br,"💾 Bericht",command=self._exp_ob,bg=C['accent'],width=130)
        self.ob_eb.pack(side=tk.LEFT,padx=(0,8)); self.ob_eb.configure(state=tk.DISABLED)

    # ── 6. Benutzer-Info ────────────────────────────────────
    def _b_userinfo(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['userinfo'] = f
        s = self._sec2(f, "👤 Benutzer-Info")
        self.ui_u = self._cb(s, "Benutzer:"); self.ui_s = self._srch(s, self._fui)
        Btn(tk.Frame(s, bg=C['panel']).also_pack(fill=tk.X, padx=10, pady=(0,4)) if False else s,
            "🔍 Info laden", command=self._load_ui, bg=C['accent'], width=140).pack(padx=10, anchor=tk.W, pady=(4,4))
        self.ui_t = self._txt(s, 12)

    # Fix: need a proper frame for button
    def _b_userinfo(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['userinfo'] = f
        s = self._sec2(f, "👤 Benutzer-Info")
        self.ui_u = self._cb(s, "Benutzer:"); self.ui_s = self._srch(s, self._fui)
        bf = tk.Frame(s, bg=C['panel']); bf.pack(fill=tk.X, padx=10, pady=(0,4))
        Btn(bf,"🔍 Info laden",command=self._load_ui,bg=C['accent'],width=140).pack(anchor=tk.W,pady=(2,4))
        rf = tk.Frame(s, bg=C['panel']); rf.pack(fill=tk.X, padx=10, pady=(0,8))
        self.ui_t = self._txt(rf, 14)

    # ── 7. Lizenz-Übersicht ─────────────────────────────────
    def _b_licenses(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['licenses'] = f
        s = self._sec2(f, "📊 Lizenz-Übersicht")
        bf = tk.Frame(s, bg=C['panel']); bf.pack(fill=tk.X, padx=10, pady=(8,4))
        Btn(bf,"📊 Lizenzen laden",command=self._load_lic,bg=C['accent'],width=160).pack(side=tk.LEFT,padx=(0,8))
        Btn(bf,"💾 CSV Export",command=self._exp_lic,bg=C['ok'],fg='#1e1e2e',width=130).pack(side=tk.LEFT)
        rf = tk.Frame(s, bg=C['panel']); rf.pack(fill=tk.X, padx=10, pady=(4,8))
        self.lic_t = self._txt(rf, 16)

    # ── 8. Shared Mailbox erstellen ─────────────────────────
    def _b_sharedmb(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['sharedmb'] = f
        s = self._sec2(f, "📧 Shared Mailbox erstellen")
        for lbl, attr in [("Name:", 'sm_name'), ("E-Mail:", 'sm_email'), ("Anzeigename:", 'sm_disp')]:
            r = tk.Frame(s, bg=C['panel']); r.pack(fill=tk.X, padx=10, pady=(4,0))
            tk.Label(r,text=lbl,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],width=12,anchor=tk.W).pack(side=tk.LEFT)
            e = tk.Entry(r,width=50,font=('Segoe UI',10),bg=C['input'],fg=C['txt'],insertbackground=C['txt'],
                         relief=tk.FLAT,highlightthickness=1,highlightbackground=C['brd'],highlightcolor=C['accent'])
            e.pack(side=tk.LEFT,fill=tk.X,expand=True,ipady=3); setattr(self, attr, e)
        # Berechtigungen direkt setzen
        pr = tk.Frame(s, bg=C['panel']); pr.pack(fill=tk.X, padx=10, pady=(6,0))
        tk.Label(pr,text="Berechtigungen (optional):",font=('Segoe UI',9,'bold'),fg=C['txt'],bg=C['panel']).pack(anchor=tk.W)
        self.sm_perm = self._cb(s, "Vollzugriff für:")
        bf = tk.Frame(s, bg=C['panel']); bf.pack(fill=tk.X, padx=10, pady=(6,8))
        Btn(bf,"📧 Erstellen",command=self._create_sm,bg=C['ok'],fg='#1e1e2e',width=150).pack(anchor=tk.W)

    # ── 9. Mail-Weiterleitungen ─────────────────────────────
    def _b_forwarding(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['forwarding'] = f
        s = self._sec2(f, "📬 Mail-Weiterleitungen verwalten")
        bf = tk.Frame(s, bg=C['panel']); bf.pack(fill=tk.X, padx=10, pady=(8,4))
        Btn(bf,"📬 Alle Weiterleitungen laden",command=self._load_fwd,bg=C['accent'],width=220).pack(side=tk.LEFT,padx=(0,8))
        rf = tk.Frame(s, bg=C['panel']); rf.pack(fill=tk.X, padx=10, pady=(0,4))
        self.fwd_t = self._txt(rf, 8)
        tk.Label(s,text="Weiterleitung setzen/entfernen:",font=('Segoe UI',10,'bold'),
                 fg=C['txt'],bg=C['panel']).pack(fill=tk.X,padx=10,pady=(4,0))
        self.fwd_src = self._cb(s, "Postfach:"); self.fwd_dst = self._cb(s, "Weiterleitung an:")
        self.fwd_keep = tk.BooleanVar(value=True)
        kr = tk.Frame(s, bg=C['panel']); kr.pack(fill=tk.X, padx=10, pady=(4,4))
        tk.Checkbutton(kr,text="📥 Kopie im Postfach behalten",variable=self.fwd_keep,font=('Segoe UI',9),
                        fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(anchor=tk.W)
        self._btns(s, [("✅ Setzen",self._set_fwd,C['ok'],tk.LEFT,140),
                        ("❌ Entfernen",self._rem_fwd,C['err'],tk.LEFT,140)])

    # ── 10. Berechtigungs-Audit ─────────────────────────────
    def _b_audit(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['audit'] = f
        s = self._sec2(f, "🔍 Berechtigungs-Audit")
        self.aud_u = self._cb(s, "Postfach:")
        self.aud_s = self._srch(s, self._faud)
        bf = tk.Frame(s, bg=C['panel']); bf.pack(fill=tk.X, padx=10, pady=(4,4))
        Btn(bf,"🔍 Audit starten",command=self._run_aud,bg=C['accent'],width=160).pack(side=tk.LEFT,padx=(0,8))
        Btn(bf,"💾 CSV Export",command=self._exp_aud,bg=C['ok'],fg='#1e1e2e',width=130).pack(side=tk.LEFT)
        rf = tk.Frame(s, bg=C['panel']); rf.pack(fill=tk.X, padx=10, pady=(4,8))
        self.aud_t = self._txt(rf, 12)

    # ── 11. CSV-Export ──────────────────────────────────────
    def _b_csvexport(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['csvexport'] = f
        s = self._sec2(f, "📋 CSV-Export")
        self.csv_opts = {}
        cbf = tk.Frame(s, bg=C['panel']); cbf.pack(fill=tk.X, padx=10, pady=(8,4))
        for i,(k,l) in enumerate([('users','👤 Alle Benutzer'),('shared','📧 Shared Mailboxen'),
                                    ('groups','👥 Alle Gruppen + Mitglieder'),('licenses','📊 Lizenzen'),
                                    ('forwarding','📬 Weiterleitungen')]):
            v = tk.BooleanVar(value=True); self.csv_opts[k] = v
            tk.Checkbutton(cbf,text=l,variable=v,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],
                           selectcolor=C['input'],activebackground=C['panel']).grid(row=i//2,column=i%2,sticky=tk.W,padx=(0,20),pady=2)
        bf = tk.Frame(s, bg=C['panel']); bf.pack(fill=tk.X, padx=10, pady=(4,8))
        Btn(bf,"📋 CSV Exportieren",command=self._run_csv,bg=C['ok'],fg='#1e1e2e',width=170).pack(anchor=tk.W)

    # ── 12. Bulk-Aktionen ───────────────────────────────────
    def _b_bulk(self):
        f = tk.Frame(self.dyn, bg=C['bg']); self.mf['bulk'] = f
        s = self._sec2(f, "🏷️ Bulk-Aktionen")
        tk.Label(s,text="Aktion:",font=('Segoe UI',10,'bold'),fg=C['txt'],bg=C['panel']).pack(fill=tk.X,padx=10,pady=(8,0))
        self.blk_act = ttk.Combobox(s,width=55,state="readonly",font=('Segoe UI',10),
                                     values=["➕ Zu Gruppe hinzufügen","➖ Aus Gruppe entfernen"])
        self.blk_act.pack(padx=10,anchor=tk.W,pady=(4,4)); self.blk_act.current(0)
        self.blk_grp = self._cb(s, "Ziel-Gruppe:")
        tk.Label(s,text="Benutzer (einer pro Zeile, E-Mail-Adressen):",font=('Segoe UI',9),
                 fg=C['dim'],bg=C['panel']).pack(fill=tk.X,padx=10,pady=(6,0))
        uf = tk.Frame(s, bg=C['panel']); uf.pack(fill=tk.X, padx=10, pady=(2,4))
        self.blk_users = tk.Text(uf,height=6,font=('Consolas',9),bg=C['input'],fg=C['txt'],
                                  relief=tk.FLAT,wrap=tk.WORD,highlightthickness=1,highlightbackground=C['brd'])
        self.blk_users.pack(fill=tk.X)
        bf = tk.Frame(s, bg=C['panel']); bf.pack(fill=tk.X, padx=10, pady=(4,4))
        Btn(bf,"📂 Aus CSV laden",command=self._blk_csv,bg=C['input'],width=140).pack(side=tk.LEFT,padx=(0,8))
        rf = tk.Frame(s, bg=C['panel']); rf.pack(fill=tk.X, padx=10, pady=(0,4))
        self.blk_t = self._txt(rf, 6)
        self._btns(s, [("▶️ Ausführen",self._run_blk,C['ok'],tk.LEFT,150)])

    # ── Verbindung ──────────────────────────────────────────
    def _chk_mod(self):
        self.conn_btn.configure(state=tk.DISABLED)
        self.log("🔍 Prüfe Modul...", C['warn'])
        def do():
            _, o, _ = self.ps.run('$m=Get-Module -ListAvailable -Name ExchangeOnlineManagement;if($m){Write-Output "OK:$($m.Version)"}else{Write-Output "MISS"}',30)
            if "OK:" in o:
                v = o.split("OK:")[1].strip().split("\n")[0]
                self.root.after(0, lambda: [self.mod_lbl.configure(text=f"✅ Modul v{v}",fg=C['ok']),
                                             self.conn_btn.configure(state=tk.NORMAL),
                                             self.log(f"✅ Modul v{v} — bereit",C['ok'])])
            else:
                self.root.after(0, lambda: [self.mod_lbl.configure(text="❌ Modul fehlt!",fg=C['err']),
                                             self.log("❌ Modul fehlt!",C['err'])])
                if messagebox.askyesno("Modul fehlt","ExchangeOnlineManagement installieren?"):
                    self.root.after(0, self._inst_mod)
        threading.Thread(target=do, daemon=True).start()

    def _inst_mod(self):
        self.log("📦 Installiere...",C['warn']); self.mod_lbl.configure(text="⏳ Installiere...",fg=C['warn'])
        def do():
            _,o,e = self.ps.run("""
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
$m=Get-Module -ListAvailable -Name ExchangeOnlineManagement;if($m){Write-Output "IOK:$($m.Version)"}else{Write-Output "IFAIL"}
""",300)
            if "IOK:" in o:
                v = o.split("IOK:")[1].strip().split("\n")[0]
                self.root.after(0, lambda: [self.mod_lbl.configure(text=f"✅ Modul v{v}",fg=C['ok']),
                                             self.conn_btn.configure(state=tk.NORMAL),
                                             self.log(f"✅ Installiert v{v}",C['ok'])])
            else:
                self.root.after(0, lambda: self.log(f"❌ Installation fehlgeschlagen: {e}",C['err']))
        threading.Thread(target=do, daemon=True).start()

    def _connect(self):
        a = self.admin_e.get().strip()
        if not a: messagebox.showwarning("Fehlt","Admin-UPN eingeben!"); return
        self.log("🔄 Verbinde...",C['warn']); self.conn_btn.configure(state=tk.DISABLED)
        def do():
            ok,_,e = self.ps.run(f'Connect-ExchangeOnline -UserPrincipalName "{a}" -ShowBanner:$false',180)
            if ok:
                v,vo,_ = self.ps.run('Get-OrganizationConfig | Select-Object -ExpandProperty Name',30)
                org = vo.strip().split("\n")[0] if v and vo.strip() else ""
                self.root.after(0, lambda: self._connected(org))
            else:
                self.root.after(0, lambda: [self.conn_btn.configure(state=tk.NORMAL),
                                             self.log(f"❌ {e}",C['err']),
                                             messagebox.showerror("Fehler",f"Verbindung fehlgeschlagen:\n{e}")])
        threading.Thread(target=do, daemon=True).start()

    def _connected(self, org):
        self.connected = True; self.conn_btn.configure(text="✅ Verbunden",bg=C['ok'])
        n = f" ({org})" if org else ""
        self.log(f"✅ Verbunden{n}!",C['ok'])
        messagebox.showinfo("Verbunden",f"Erfolgreich{n}!\nDaten werden geladen...")
        self._load()

    def disconnect(self):
        self.log("🔌 Trenne...",C['warn'])
        self.ps.run("Disconnect-ExchangeOnline -Confirm:$false",30)
        self.connected = False; self.mailboxes = []
        for k in self.groups: self.groups[k] = []
        self.conn_btn.configure(text="🔌 Verbinden",bg=C['accent']); self.conn_btn.configure(state=tk.NORMAL)
        for cb in [self.mb_t,self.mb_u,self.ob_u,self.ob_fwd,self.ui_u,self.sm_perm,
                   self.fwd_src,self.fwd_dst,self.aud_u,self.blk_grp]:
            try: cb.configure(values=[],state="disabled"); cb.set("-- Erst verbinden --")
            except: pass
        for k in ['teams','verteiler','security']:
            for a in ['gc','uc']:
                try: c=getattr(self,f'{k}_{a}'); c.configure(values=[],state="disabled"); c.set("-- Erst verbinden --")
                except: pass
        self.log("✅ Getrennt.",C['ok'])

    # ── Daten laden ─────────────────────────────────────────
    def _load(self):
        self.log("📥 Lade Daten...",C['warn'])
        def do():
            r1,o1,_ = self.ps.run('Get-Mailbox -ResultSize Unlimited|Select DisplayName,PrimarySmtpAddress,RecipientTypeDetails|ConvertTo-Json -Compress',180)
            r2,o2,_ = self.ps.run('Get-UnifiedGroup -ResultSize Unlimited|Select DisplayName,PrimarySmtpAddress,GroupType|ConvertTo-Json -Compress',180)
            r3,o3,_ = self.ps.run('Get-DistributionGroup -ResultSize Unlimited|Select DisplayName,PrimarySmtpAddress,GroupType|ConvertTo-Json -Compress',180)
            self.root.after(0, lambda: self._loaded(r1,o1,r2,o2,r3,o3))
        threading.Thread(target=do, daemon=True).start()

    def _loaded(self, r1,o1,r2,o2,r3,o3):
        if r1:
            self.mailboxes = []
            for mb in self._pj(o1):
                e,n,t = mb.get('PrimarySmtpAddress',''),mb.get('DisplayName',''),mb.get('RecipientTypeDetails','')
                if e:
                    sh = 'Shared' in t
                    self.mailboxes.append({'d':f"{'👥' if sh else '👤'} {n} <{e}>",'e':e,'n':n,'t':'shared' if sh else 'user'})
            self.mailboxes.sort(key=lambda x:x['d'])
            self._amb(); self._upd_ob(); self._upd_misc()
            self.log(f"  📧 {len(self.mailboxes)} Postfächer",C['dim'])
        if r2:
            self.groups['teams'] = [{'d':f"👥 {g.get('DisplayName','')} <{g.get('PrimarySmtpAddress','')}>",
                                      'e':g.get('PrimarySmtpAddress',''),'n':g.get('DisplayName','')}
                                     for g in self._pj(o2) if g.get('PrimarySmtpAddress')]
            self.groups['teams'].sort(key=lambda x:x['d']); self._ugrp('teams')
            self.log(f"  👥 {len(self.groups['teams'])} Teams/M365",C['dim'])
        if r3:
            self.groups['verteiler']=[]; self.groups['security']=[]
            for g in self._pj(o3):
                e,n,gt = g.get('PrimarySmtpAddress',''),g.get('DisplayName',''),str(g.get('GroupType',''))
                if e:
                    ent = {'d':f"{'🔒' if 'Security' in gt else '📨'} {n} <{e}>",'e':e,'n':n}
                    (self.groups['security'] if 'Security' in gt else self.groups['verteiler']).append(ent)
            for k in ['verteiler','security']:
                self.groups[k].sort(key=lambda x:x['d']); self._ugrp(k)
            self.log(f"  📨 {len(self.groups['verteiler'])} Verteiler, 🔒 {len(self.groups['security'])} Security",C['dim'])
        t = len(self.mailboxes)+sum(len(v) for v in self.groups.values())
        self.log(f"✅ {t} Objekte geladen!",C['ok'])

    def _ugrp(self, k):
        i = [g['d'] for g in self.groups[k]]
        getattr(self,f'{k}_gc').configure(values=i,state="normal"); getattr(self,f'{k}_gc').set("")
        ui = [m['d'] for m in self.mailboxes]
        getattr(self,f'{k}_uc').configure(values=ui,state="normal"); getattr(self,f'{k}_uc').set("")

    def _upd_ob(self):
        i = [m['d'] for m in self.mailboxes]
        self.ob_u.configure(values=i,state="normal"); self.ob_u.set("")
        self.ob_fwd.configure(values=i,state="normal"); self.ob_fwd.set("")

    def _upd_misc(self):
        i = [m['d'] for m in self.mailboxes]
        for cb in [self.ui_u,self.sm_perm,self.fwd_src,self.fwd_dst,self.aud_u]:
            try: cb.configure(values=i,state="normal"); cb.set("")
            except: pass
        # Bulk: alle Gruppen
        all_g = []
        for k in ['teams','verteiler','security']:
            all_g += [g['d'] for g in self.groups[k]]
        self.blk_grp.configure(values=all_g,state="normal"); self.blk_grp.set("")

    # ── Filter ──────────────────────────────────────────────
    def _amb(self):
        ft = self.mb_ft.get()
        fl = self.mailboxes if ft=="all" else [m for m in self.mailboxes if m['t']==ft]
        i = [m['d'] for m in fl]
        self.mb_t.configure(values=i,state="normal"); self.mb_u.configure(values=i,state="normal")
    def _fmb(self, e=None):
        s = self.mb_s.get().lower(); ft = self.mb_ft.get()
        b = self.mailboxes if ft=="all" else [m for m in self.mailboxes if m['t']==ft]
        fl = [m for m in b if s in m['d'].lower()] if s else b
        i = [m['d'] for m in fl]
        self.mb_t.configure(values=i); self.mb_u.configure(values=i)
    def _fgrp(self, k):
        s = getattr(self,f'{k}_se').get().lower()
        fl = [g for g in self.groups[k] if s in g['d'].lower()] if s else self.groups[k]
        getattr(self,f'{k}_gc').configure(values=[g['d'] for g in fl])
    def _fob(self, e=None):
        s = self.ob_s.get().lower()
        fl = [m for m in self.mailboxes if s in m['d'].lower()] if s else self.mailboxes
        self.ob_u.configure(values=[m['d'] for m in fl])
    def _fui(self, e=None):
        s = self.ui_s.get().lower()
        fl = [m for m in self.mailboxes if s in m['d'].lower()] if s else self.mailboxes
        self.ui_u.configure(values=[m['d'] for m in fl])
    def _faud(self, e=None):
        s = self.aud_s.get().lower()
        fl = [m for m in self.mailboxes if s in m['d'].lower()] if s else self.mailboxes
        self.aud_u.configure(values=[m['d'] for m in fl])

    # ── 1. Postfach-Logik ───────────────────────────────────
    def _add_mb(self):
        mb,us = self.mb_t.get().strip(), self.mb_u.get().strip()
        if not mb or mb.startswith("--") or not us or us.startswith("--"):
            messagebox.showwarning("Fehlt","Postfach und Benutzer auswählen!"); return
        mbe,use = self._ge(mb), self._ge(us)
        if not messagebox.askyesno("Bestätigen",f"Hinzufügen?\n📬 {mbe}\n👤 {use}"): return
        def do():
            errs = []
            if self.fa_v.get():
                am = "$true" if self.am_v.get() else "$false"
                ok,_,e = self.ps.run(f'Add-MailboxPermission -Identity "{mbe}" -User "{use}" -AccessRights FullAccess -AutoMapping {am}')
                self.root.after(0, lambda: self.log(f"  {'✅ Vollzugriff' if ok else '❌ '+e}",C['ok'] if ok else C['err']))
                if not ok: errs.append(e)
            if self.sa_v.get():
                ok,_,e = self.ps.run(f'Add-RecipientPermission -Identity "{mbe}" -Trustee "{use}" -AccessRights SendAs -Confirm:$false')
                self.root.after(0, lambda: self.log(f"  {'✅ Senden als' if ok else '❌ '+e}",C['ok'] if ok else C['err']))
                if not ok: errs.append(e)
            self.root.after(0, lambda: self._done(errs,"hinzugefügt"))
        threading.Thread(target=do, daemon=True).start()

    def _rem_mb(self):
        mb,us = self.mb_t.get().strip(), self.mb_u.get().strip()
        if not mb or mb.startswith("--") or not us or us.startswith("--"):
            messagebox.showwarning("Fehlt","Auswählen!"); return
        mbe,use = self._ge(mb), self._ge(us)
        if not messagebox.askyesno("⚠️",f"Entfernen?\n📬 {mbe}\n👤 {use}",icon="warning"): return
        def do():
            errs = []
            if self.fa_v.get():
                ok,_,e = self.ps.run(f'Remove-MailboxPermission -Identity "{mbe}" -User "{use}" -AccessRights FullAccess -Confirm:$false')
                if not ok: errs.append(e)
            if self.sa_v.get():
                ok,_,e = self.ps.run(f'Remove-RecipientPermission -Identity "{mbe}" -Trustee "{use}" -AccessRights SendAs -Confirm:$false')
                if not ok: errs.append(e)
            self.root.after(0, lambda: self._done(errs,"entfernt"))
        threading.Thread(target=do, daemon=True).start()

    # ── 2-4. Gruppen-Logik ──────────────────────────────────
    def _agrp(self, k):
        g,u = getattr(self,f'{k}_gc').get().strip(), getattr(self,f'{k}_uc').get().strip()
        if not g or g.startswith("--") or not u or u.startswith("--"):
            messagebox.showwarning("Fehlt","Auswählen!"); return
        ge,ue = self._ge(g),self._ge(u); rv = getattr(self,f'{k}_rv').get()
        if not messagebox.askyesno("Bestätigen",f"Hinzufügen?\n📋 {ge}\n👤 {ue}"): return
        def do():
            if k=='teams': cmd = f'Add-UnifiedGroupLinks -Identity "{ge}" -LinkType {"Owners" if rv=="Owner" else "Members"} -Links "{ue}"'
            else: cmd = f'Add-DistributionGroupMember -Identity "{ge}" -Member "{ue}"'
            ok,_,e = self.ps.run(cmd)
            self.root.after(0, lambda: self._done([] if ok else [e],"hinzugefügt"))
        threading.Thread(target=do, daemon=True).start()

    def _rgrp(self, k):
        g,u = getattr(self,f'{k}_gc').get().strip(), getattr(self,f'{k}_uc').get().strip()
        if not g or g.startswith("--") or not u or u.startswith("--"):
            messagebox.showwarning("Fehlt","Auswählen!"); return
        ge,ue = self._ge(g),self._ge(u)
        if not messagebox.askyesno("⚠️",f"Entfernen?\n📋 {ge}\n👤 {ue}",icon="warning"): return
        def do():
            if k=='teams':
                ok,_,e = self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Members -Links "{ue}" -Confirm:$false')
                self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Owners -Links "{ue}" -Confirm:$false')
            else: ok,_,e = self.ps.run(f'Remove-DistributionGroupMember -Identity "{ge}" -Member "{ue}" -Confirm:$false')
            self.root.after(0, lambda: self._done([] if ok else [e],"entfernt"))
        threading.Thread(target=do, daemon=True).start()

    def _smem(self, k):
        g = getattr(self,f'{k}_gc').get().strip()
        if not g or g.startswith("--"): messagebox.showwarning("Fehlt","Gruppe auswählen!"); return
        ge = self._ge(g); self.log(f"  📋 Lade Mitglieder {ge}...",C['warn'])
        def do():
            if k=='teams':
                _,om,_ = self.ps.run(f'Get-UnifiedGroupLinks -Identity "{ge}" -LinkType Members|Select Name,PrimarySmtpAddress|ConvertTo-Json -Compress',60)
                _,oo,_ = self.ps.run(f'Get-UnifiedGroupLinks -Identity "{ge}" -LinkType Owners|Select Name,PrimarySmtpAddress|ConvertTo-Json -Compress',60)
                ms,ow = self._pj(om),self._pj(oo)
                ln = [f"=== {ge} ===",""]
                if ow: ln += [f"👑 Besitzer ({len(ow)}):"] + [f"  • {o.get('Name','')} <{o.get('PrimarySmtpAddress','')}>" for o in ow] + [""]
                ln += [f"👤 Mitglieder ({len(ms)}):"] + [f"  • {m.get('Name','')} <{m.get('PrimarySmtpAddress','')}>" for m in ms]
            else:
                _,o,_ = self.ps.run(f'Get-DistributionGroupMember -Identity "{ge}" -ResultSize Unlimited|Select Name,PrimarySmtpAddress|ConvertTo-Json -Compress',60)
                ms = self._pj(o)
                ln = [f"=== {ge} ===","",f"👤 Mitglieder ({len(ms)}):"] + [f"  • {m.get('Name','')} <{m.get('PrimarySmtpAddress','')}>" for m in ms]
            self.root.after(0, lambda: self._settxt(getattr(self,f'{k}_mt'),"\n".join(ln)))
        threading.Thread(target=do, daemon=True).start()

    # ── 5. Offboarding-Logik ────────────────────────────────
    def _run_ob(self):
        us = self.ob_u.get().strip()
        if not us or us.startswith("--"): messagebox.showwarning("Fehlt","Benutzer auswählen!"); return
        active = {k:v.get() for k,v in self.ob_v.items()}
        if not any(active.values()): messagebox.showwarning("Fehlt","Mindestens einen Schritt!"); return
        ue = self._ge(us); un = us.split('<')[0].strip().lstrip('👤👥 ')
        names = [OB_NAMES[k] for k,v in active.items() if v]
        msg = f"⚠️ OFFBOARDING:\n\n👤 {un}\n📧 {ue}\n\n" + "\n".join(f"  • {n}" for n in names) + "\n\n⚠️ Irreversibel!"
        if not messagebox.askyesno("⚠️",msg,icon="warning"): return
        if not messagebox.askyesno("🔴","WIRKLICH fortfahren?",icon="warning"): return
        self.ob_rb.configure(state=tk.DISABLED); self.ob_eb.configure(state=tk.DISABLED)
        self._settxt(self.ob_rt,"")
        self.ob_report = ["="*60,"OFFBOARDING-BERICHT",f"Datum: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
                           f"Benutzer: {un}",f"E-Mail: {ue}",f"Admin: {self.admin_e.get().strip()}","="*60,""]
        def do():
            todo = [(k,v) for k,v in active.items() if v]; total = len(todo); res = {}
            for i,(sk,_) in enumerate(todo):
                pct = int(i/total*100); sn = OB_NAMES[sk]
                self.root.after(0, lambda p=pct,s=sn: [self.ob_pb.configure(value=p),self.ob_pl.configure(text=f"⏳ {s}... ({p}%)")])
                self.root.after(0, lambda s=sn: self.log(f"  🔄 {s}...",C['warn']))
                ok,det = self._ob_step(sk,ue); res[sk]=(ok,det)
                st = "✅" if ok else "❌"; self.ob_report += [f"[{st}] {sn}",f"    {det}",""]
                self.root.after(0, lambda s=sn,st2=st,c2=(C['ok'] if ok else C['err']): self.log(f"  {st2} {s}",c2))
            sc = sum(1 for ok,_ in res.values() if ok); fc = sum(1 for ok,_ in res.values() if not ok)
            self.ob_report += ["="*60,f"ERGEBNIS: {sc} OK, {fc} FEHLER","="*60]
            self.root.after(0, lambda: [self.ob_pb.configure(value=100),self.ob_pl.configure(text="✅ Fertig"),
                                         self.ob_rb.configure(state=tk.NORMAL),self.ob_eb.configure(state=tk.NORMAL),
                                         self._settxt(self.ob_rt,"\n".join(self.ob_report)),
                                         self.log(f"🚪 Offboarding {ue}: {sc}✅/{fc}❌",C['ok'] if fc==0 else C['warn']),
                                         messagebox.showinfo("Fertig",f"{sc} OK, {fc} Fehler") if fc==0 else
                                         messagebox.showwarning("Teilweise",f"{sc} OK, {fc} Fehler")])
        threading.Thread(target=do, daemon=True).start()

    def _ob_step(self, step, ue):
        try:
            if step=='sign_in':
                ok,_,e = self.ps.run(f'Set-User -Identity "{ue}" -AccountDisabled $true',60)
                return ok, "Anmeldung blockiert" if ok else f"Fehler: {e}"
            elif step=='reset_pw':
                ok,o,e = self.ps.run(f'$c="abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%&*";$pw=-join(1..24|ForEach-Object{{$c[(Get-Random -Max $c.Length)]}});Set-Mailbox -Identity "{ue}" -Password (ConvertTo-SecureString $pw -AsPlainText -Force) -ErrorAction Stop;Write-Output "PWOK"',60)
                return ("PWOK" in o), "Passwort zurückgesetzt" if "PWOK" in o else f"Fehler: {e}"
            elif step=='remove_groups':
                rm,fl = 0,0
                ok,o,_ = self.ps.run(f'Get-UnifiedGroup -ResultSize Unlimited|Where-Object{{(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})-or(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Owners -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand PrimarySmtpAddress',180)
                if ok and o.strip():
                    for g in o.strip().split("\n"):
                        g=g.strip();
                        if not g: continue
                        r,_,_ = self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{g}" -LinkType Members -Links "{ue}" -Confirm:$false -EA SilentlyContinue')
                        self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{g}" -LinkType Owners -Links "{ue}" -Confirm:$false -EA SilentlyContinue')
                        if r: rm+=1
                        else: fl+=1
                ok,o,_ = self.ps.run(f'Get-DistributionGroup -ResultSize Unlimited|Where-Object{{(Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand PrimarySmtpAddress',180)
                if ok and o.strip():
                    for g in o.strip().split("\n"):
                        g=g.strip();
                        if not g: continue
                        r,_,_ = self.ps.run(f'Remove-DistributionGroupMember -Identity "{g}" -Member "{ue}" -Confirm:$false -EA SilentlyContinue')
                        if r: rm+=1
                        else: fl+=1
                return fl==0, f"Aus {rm} Gruppe(n) entfernt" + (f", {fl} Fehler" if fl else "")
            elif step=='remove_licenses':
                ok,o,e = self.ps.run(f'$s=(Get-MgUserLicenseDetail -UserId "{ue}" -EA SilentlyContinue).SkuId;if($s){{foreach($k in $s){{Set-MgUserLicense -UserId "{ue}" -RemoveLicenses @($k) -AddLicenses @() -EA Stop}};Write-Output "LR:$($s.Count)"}}else{{Write-Output "NOGRAPH"}}',90)
                if "LR:" in o: return True, f"{o.split('LR:')[1].strip().split(chr(10))[0]} Lizenz(en) entfernt"
                return False, "Microsoft.Graph nicht verfügbar — manuell entfernen" if "NOGRAPH" in o else f"Fehler: {e}"
            elif step=='convert_shared':
                ok,_,e = self.ps.run(f'Set-Mailbox -Identity "{ue}" -Type Shared',60)
                return ok, "Shared Mailbox" if ok else f"Fehler: {e}"
            elif step=='set_ooo':
                msg = self.ob_ooo.get('1.0',tk.END).strip()
                if not msg: return True, "Übersprungen"
                esc = msg.replace("'","''").replace('"','`"')
                ok,_,e = self.ps.run(f'Set-MailboxAutoReplyConfiguration -Identity "{ue}" -AutoReplyState Enabled -InternalMessage "{esc}" -ExternalMessage "{esc}" -ExternalAudience All',60)
                return ok, "OOO aktiviert" if ok else f"Fehler: {e}"
            elif step=='fwd':
                fw = self.ob_fwd.get().strip()
                if not fw: return True, "Übersprungen"
                fe = self._ge(fw) if '<' in fw else fw
                ok,_,e = self.ps.run(f'Set-Mailbox -Identity "{ue}" -ForwardingSmtpAddress "smtp:{fe}" -DeliverToMailboxAndForward $true',60)
                return ok, f"→ {fe}" if ok else f"Fehler: {e}"
            elif step=='hide_gal':
                ok,_,e = self.ps.run(f'Set-Mailbox -Identity "{ue}" -HiddenFromAddressListsEnabled $true',60)
                return ok, "Ausgeblendet" if ok else f"Fehler: {e}"
            elif step=='disable_sync':
                ok,_,e = self.ps.run(f'Set-CASMailbox -Identity "{ue}" -ActiveSyncEnabled $false -OWAEnabled $false -PopEnabled $false -ImapEnabled $false -MAPIEnabled $false -EwsEnabled $false',60)
                return ok, "Alle Protokolle deaktiviert" if ok else f"Fehler: {e}"
            elif step=='remove_delegates':
                ok,o,e = self.ps.run(f'$p=Get-MailboxPermission -Identity "{ue}"|Where-Object{{$_.User -ne "NT AUTHORITY\\SELF" -and $_.IsInherited -eq $false}};$c=0;foreach($x in $p){{Remove-MailboxPermission -Identity "{ue}" -User $x.User -AccessRights $x.AccessRights -Confirm:$false -EA SilentlyContinue;$c++}};Write-Output "DD:$c"',90)
                if "DD:" in o: return True, f"{o.split('DD:')[1].strip().split(chr(10))[0]} entfernt"
                return False, f"Fehler: {e}"
            return False, "Unbekannt"
        except Exception as ex: return False, str(ex)

    def _exp_ob(self):
        if not self.ob_report: return
        us = self.ob_u.get().strip(); ue = self._ge(us) if '<' in us else "user"
        un = ue.split('@')[0] if '@' in ue else ue
        fp = filedialog.asksaveasfilename(defaultextension=".txt",
            initialfile=f"Offboarding_{un}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            filetypes=[("Text","*.txt")])
        if fp:
            with open(fp,'w',encoding='utf-8') as f: f.write("\n".join(self.ob_report))
            self.log(f"💾 {fp}",C['ok']); messagebox.showinfo("OK",f"Gespeichert:\n{fp}")

    # ── 6. Benutzer-Info ────────────────────────────────────
    def _load_ui(self):
        us = self.ui_u.get().strip()
        if not us or us.startswith("--"): messagebox.showwarning("Fehlt","Benutzer auswählen!"); return
        ue = self._ge(us); self.log(f"  🔍 Lade Info für {ue}...",C['warn'])
        def do():
            lines = [f"{'='*50}",f"  BENUTZER-INFO: {ue}",f"{'='*50}",""]
            # Postfach-Details
            ok,o,_ = self.ps.run(f'Get-Mailbox -Identity "{ue}"|Select DisplayName,PrimarySmtpAddress,RecipientTypeDetails,ForwardingSmtpAddress,HiddenFromAddressListsEnabled,WhenCreated|ConvertTo-Json',60)
            if ok:
                d = self._pj(o); mb = d[0] if d else {}
                lines += [f"📧 Name: {mb.get('DisplayName','')}",f"📧 Typ: {mb.get('RecipientTypeDetails','')}",
                          f"📨 Weiterleitung: {mb.get('ForwardingSmtpAddress','Keine')}",
                          f"👻 GAL versteckt: {mb.get('HiddenFromAddressListsEnabled','')}",
                          f"📅 Erstellt: {mb.get('WhenCreated','')}",""]
            # Postfachgröße
            ok,o,_ = self.ps.run(f'Get-MailboxStatistics -Identity "{ue}" -ErrorAction SilentlyContinue|Select TotalItemSize,ItemCount|ConvertTo-Json',30)
            if ok:
                d = self._pj(o); st = d[0] if d else {}
                lines += [f"📦 Größe: {st.get('TotalItemSize','')}",f"📬 Elemente: {st.get('ItemCount','')}",""]
            # Gruppen
            ok,o,_ = self.ps.run(f'Get-UnifiedGroup -ResultSize Unlimited|Where-Object{{(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand DisplayName',120)
            if ok and o.strip():
                gs = [g.strip() for g in o.strip().split("\n") if g.strip()]
                lines += [f"👥 Teams/M365 Gruppen ({len(gs)}):"] + [f"  • {g}" for g in gs] + [""]
            ok,o,_ = self.ps.run(f'Get-DistributionGroup -ResultSize Unlimited|Where-Object{{(Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand DisplayName',120)
            if ok and o.strip():
                gs = [g.strip() for g in o.strip().split("\n") if g.strip()]
                lines += [f"📨 Verteilerlisten/Security ({len(gs)}):"] + [f"  • {g}" for g in gs] + [""]
            # Berechtigungen auf andere Postfächer
            ok,o,_ = self.ps.run(f'Get-Mailbox -ResultSize Unlimited|Get-MailboxPermission|Where-Object{{$_.User -like "*{ue}*" -and $_.AccessRights -like "*FullAccess*"}}|Select -Expand Identity',60)
            if ok and o.strip():
                ps = [p.strip() for p in o.strip().split("\n") if p.strip()]
                lines += [f"🔑 Vollzugriff auf ({len(ps)}):"] + [f"  • {p}" for p in ps]
            self.root.after(0, lambda: [self._settxt(self.ui_t,"\n".join(lines)),
                                         self.log(f"  ✅ Info geladen",C['ok'])])
        threading.Thread(target=do, daemon=True).start()

    # ── 7. Lizenz-Übersicht ─────────────────────────────────
    def _load_lic(self):
        self.log("  📊 Lade Lizenzen...",C['warn'])
        def do():
            ok,o,e = self.ps.run('Get-MgSubscribedSku -ErrorAction SilentlyContinue|Select SkuPartNumber,ConsumedUnits,@{N="Total";E={$_.PrepaidUnits.Enabled}}|ConvertTo-Json',60)
            if ok and o.strip():
                data = self._pj(o)
                lines = [f"{'Lizenz':<40} {'Benutzt':>8} {'Gesamt':>8} {'Frei':>8}","─"*68]
                for d in data:
                    n,u,t = d.get('SkuPartNumber',''),d.get('ConsumedUnits',0),d.get('Total',0)
                    try: u,t = int(u),int(t)
                    except: u,t = 0,0
                    lines.append(f"{n:<40} {u:>8} {t:>8} {t-u:>8}")
                self.root.after(0, lambda: [self._settxt(self.lic_t,"\n".join(lines)),self.log("  ✅ Lizenzen geladen",C['ok'])])
                self._lic_data = data
            else:
                self.root.after(0, lambda: [self._settxt(self.lic_t,"❌ Microsoft.Graph Modul nicht verfügbar.\nBitte installieren: Install-Module Microsoft.Graph -Scope CurrentUser"),
                                             self.log(f"  ❌ Graph nicht verfügbar: {e}",C['err'])])
        threading.Thread(target=do, daemon=True).start()

    def _exp_lic(self):
        if not hasattr(self,'_lic_data') or not self._lic_data: messagebox.showwarning("Fehlt","Erst laden!"); return
        fp = filedialog.asksaveasfilename(defaultextension=".csv",initialfile=f"Lizenzen_{datetime.now().strftime('%Y%m%d')}.csv",filetypes=[("CSV","*.csv")])
        if fp:
            with open(fp,'w',newline='',encoding='utf-8') as f:
                w = csv.writer(f,delimiter=';'); w.writerow(['Lizenz','Benutzt','Gesamt','Frei'])
                for d in self._lic_data:
                    n,u,t = d.get('SkuPartNumber',''),d.get('ConsumedUnits',0),d.get('Total',0)
                    try: u,t = int(u),int(t)
                    except: u,t = 0,0
                    w.writerow([n,u,t,t-u])
            self.log(f"💾 {fp}",C['ok']); messagebox.showinfo("OK",f"Gespeichert:\n{fp}")

    # ── 8. Shared Mailbox erstellen ─────────────────────────
    def _create_sm(self):
        nm,em,dp = self.sm_name.get().strip(),self.sm_email.get().strip(),self.sm_disp.get().strip()
        if not nm or not em: messagebox.showwarning("Fehlt","Name und E-Mail eingeben!"); return
        if not dp: dp = nm
        if not messagebox.askyesno("Erstellen",f"Shared Mailbox erstellen?\n\n📧 {em}\n👤 {dp}"): return
        def do():
            self.root.after(0, lambda: self.log(f"  📧 Erstelle {em}...",C['warn']))
            ok,_,e = self.ps.run(f'New-Mailbox -Name "{nm}" -PrimarySmtpAddress "{em}" -DisplayName "{dp}" -Shared',60)
            if ok:
                self.root.after(0, lambda: self.log(f"  ✅ {em} erstellt!",C['ok']))
                pu = self.sm_perm.get().strip()
                if pu and not pu.startswith("--"):
                    pue = self._ge(pu)
                    self.ps.run(f'Add-MailboxPermission -Identity "{em}" -User "{pue}" -AccessRights FullAccess -AutoMapping $true')
                    self.ps.run(f'Add-RecipientPermission -Identity "{em}" -Trustee "{pue}" -AccessRights SendAs -Confirm:$false')
                    self.root.after(0, lambda: self.log(f"  ✅ Berechtigungen für {pue} gesetzt",C['ok']))
                self.root.after(0, lambda: messagebox.showinfo("OK",f"✅ {em} erstellt!"))
            else:
                self.root.after(0, lambda: [self.log(f"  ❌ {e}",C['err']),messagebox.showerror("Fehler",e)])
        threading.Thread(target=do, daemon=True).start()

    # ── 9. Weiterleitungen ──────────────────────────────────
    def _load_fwd(self):
        self.log("  📬 Lade Weiterleitungen...",C['warn'])
        def do():
            ok,o,_ = self.ps.run('Get-Mailbox -ResultSize Unlimited|Where-Object{$_.ForwardingSmtpAddress -ne $null}|Select DisplayName,PrimarySmtpAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward|ConvertTo-Json -Compress',120)
            data = self._pj(o) if ok else []
            if data:
                lines = [f"{'Postfach':<35} {'Weiterleitung an':<35} {'Kopie':>5}","─"*78]
                for d in data:
                    lines.append(f"{d.get('PrimarySmtpAddress',''):<35} {str(d.get('ForwardingSmtpAddress','')):<35} {'Ja' if d.get('DeliverToMailboxAndForward') else 'Nein':>5}")
                self.root.after(0, lambda: [self._settxt(self.fwd_t,"\n".join(lines)),self.log(f"  ✅ {len(data)} Weiterleitungen",C['ok'])])
            else:
                self.root.after(0, lambda: [self._settxt(self.fwd_t,"Keine Weiterleitungen konfiguriert."),self.log("  ✅ Keine Weiterleitungen",C['ok'])])
        threading.Thread(target=do, daemon=True).start()

    def _set_fwd(self):
        src,dst = self.fwd_src.get().strip(),self.fwd_dst.get().strip()
        if not src or src.startswith("--") or not dst or dst.startswith("--"):
            messagebox.showwarning("Fehlt","Postfach und Ziel auswählen!"); return
        se,de = self._ge(src),self._ge(dst); keep = "$true" if self.fwd_keep.get() else "$false"
        if not messagebox.askyesno("Setzen",f"Weiterleitung?\n📬 {se} → {de}"): return
        def do():
            ok,_,e = self.ps.run(f'Set-Mailbox -Identity "{se}" -ForwardingSmtpAddress "smtp:{de}" -DeliverToMailboxAndForward {keep}')
            self.root.after(0, lambda: self._done([] if ok else [e],"gesetzt"))
        threading.Thread(target=do, daemon=True).start()

    def _rem_fwd(self):
        src = self.fwd_src.get().strip()
        if not src or src.startswith("--"): messagebox.showwarning("Fehlt","Postfach auswählen!"); return
        se = self._ge(src)
        if not messagebox.askyesno("Entfernen",f"Weiterleitung entfernen für {se}?"): return
        def do():
            ok,_,e = self.ps.run(f'Set-Mailbox -Identity "{se}" -ForwardingSmtpAddress $null')
            self.root.after(0, lambda: self._done([] if ok else [e],"entfernt"))
        threading.Thread(target=do, daemon=True).start()

    # ── 10. Berechtigungs-Audit ─────────────────────────────
    def _run_aud(self):
        us = self.aud_u.get().strip()
        if not us or us.startswith("--"): messagebox.showwarning("Fehlt","Postfach auswählen!"); return
        ue = self._ge(us); self.log(f"  🔍 Audit für {ue}...",C['warn'])
        self._aud_data = []
        def do():
            lines = [f"{'='*50}",f"  BERECHTIGUNGS-AUDIT: {ue}",f"{'='*50}",""]
            # FullAccess
            ok,o,_ = self.ps.run(f'Get-MailboxPermission -Identity "{ue}"|Where-Object{{$_.User -ne "NT AUTHORITY\\SELF" -and $_.IsInherited -eq $false}}|Select User,AccessRights|ConvertTo-Json -Compress',60)
            perms = self._pj(o) if ok else []
            if perms:
                lines += ["📂 Vollzugriff:"]
                for p in perms:
                    lines.append(f"  • {p.get('User','')} → {p.get('AccessRights','')}")
                    self._aud_data.append({'Postfach':ue,'Typ':'FullAccess','Benutzer':str(p.get('User',''))})
                lines.append("")
            # SendAs
            ok,o,_ = self.ps.run(f'Get-RecipientPermission -Identity "{ue}"|Where-Object{{$_.Trustee -ne "NT AUTHORITY\\SELF"}}|Select Trustee,AccessRights|ConvertTo-Json -Compress',60)
            perms = self._pj(o) if ok else []
            if perms:
                lines += ["✉️ Senden als:"]
                for p in perms:
                    lines.append(f"  • {p.get('Trustee','')}")
                    self._aud_data.append({'Postfach':ue,'Typ':'SendAs','Benutzer':str(p.get('Trustee',''))})
                lines.append("")
            # SendOnBehalf
            ok,o,_ = self.ps.run(f'Get-Mailbox -Identity "{ue}"|Select -Expand GrantSendOnBehalfTo',30)
            if ok and o.strip():
                sob = [x.strip() for x in o.strip().split("\n") if x.strip()]
                lines += ["📤 Senden im Auftrag:"] + [f"  • {s}" for s in sob]
                for s in sob: self._aud_data.append({'Postfach':ue,'Typ':'SendOnBehalf','Benutzer':s})
            if not self._aud_data: lines.append("✅ Keine Berechtigungen gefunden.")
            self.root.after(0, lambda: [self._settxt(self.aud_t,"\n".join(lines)),self.log(f"  ✅ Audit fertig: {len(self._aud_data)} Einträge",C['ok'])])
        threading.Thread(target=do, daemon=True).start()

    def _exp_aud(self):
        if not hasattr(self,'_aud_data') or not self._aud_data: messagebox.showwarning("Fehlt","Erst Audit starten!"); return
        fp = filedialog.asksaveasfilename(defaultextension=".csv",initialfile=f"Audit_{datetime.now().strftime('%Y%m%d')}.csv",filetypes=[("CSV","*.csv")])
        if fp:
            with open(fp,'w',newline='',encoding='utf-8') as f:
                w = csv.DictWriter(f,fieldnames=['Postfach','Typ','Benutzer'],delimiter=';'); w.writeheader(); w.writerows(self._aud_data)
            self.log(f"💾 {fp}",C['ok']); messagebox.showinfo("OK",f"Gespeichert:\n{fp}")

    # ── 11. CSV-Export ──────────────────────────────────────
    def _run_csv(self):
        active = {k:v.get() for k,v in self.csv_opts.items()}
        if not any(active.values()): messagebox.showwarning("Fehlt","Mindestens eine Option!"); return
        fp = filedialog.askdirectory(title="Export-Ordner wählen")
        if not fp: return
        self.log("📋 CSV-Export gestartet...",C['warn'])
        def do():
            exports = []
            if active.get('users'):
                p = os.path.join(fp,f"Benutzer_{datetime.now().strftime('%Y%m%d')}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w = csv.writer(f,delimiter=';'); w.writerow(['Name','E-Mail','Typ'])
                    for m in self.mailboxes: w.writerow([m['n'],m['e'],m['t']])
                exports.append(p)
            if active.get('shared'):
                p = os.path.join(fp,f"SharedMailboxen_{datetime.now().strftime('%Y%m%d')}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w = csv.writer(f,delimiter=';'); w.writerow(['Name','E-Mail'])
                    for m in self.mailboxes:
                        if m['t']=='shared': w.writerow([m['n'],m['e']])
                exports.append(p)
            if active.get('groups'):
                p = os.path.join(fp,f"Gruppen_{datetime.now().strftime('%Y%m%d')}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w = csv.writer(f,delimiter=';'); w.writerow(['Typ','Name','E-Mail'])
                    for k in ['teams','verteiler','security']:
                        for g in self.groups[k]: w.writerow([k,g['n'],g['e']])
                exports.append(p)
            if active.get('forwarding'):
                ok,o,_ = self.ps.run('Get-Mailbox -ResultSize Unlimited|Where-Object{$_.ForwardingSmtpAddress -ne $null}|Select PrimarySmtpAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward|ConvertTo-Json -Compress',120)
                data = self._pj(o) if ok else []
                p = os.path.join(fp,f"Weiterleitungen_{datetime.now().strftime('%Y%m%d')}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w = csv.writer(f,delimiter=';'); w.writerow(['Postfach','Weiterleitung','Kopie'])
                    for d in data: w.writerow([d.get('PrimarySmtpAddress',''),d.get('ForwardingSmtpAddress',''),d.get('DeliverToMailboxAndForward','')])
                exports.append(p)
            self.root.after(0, lambda: [self.log(f"✅ {len(exports)} CSV(s) exportiert nach {fp}",C['ok']),
                                         messagebox.showinfo("Export",f"{len(exports)} Datei(en) exportiert:\n{fp}")])
        threading.Thread(target=do, daemon=True).start()

    # ── 12. Bulk-Aktionen ───────────────────────────────────
    def _blk_csv(self):
        fp = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("Text","*.txt")])
        if fp:
            with open(fp,'r',encoding='utf-8') as f:
                lines = [l.strip() for l in f if l.strip() and '@' in l]
            self.blk_users.delete('1.0',tk.END); self.blk_users.insert('1.0',"\n".join(lines))
            self.log(f"  📂 {len(lines)} Benutzer aus CSV geladen",C['ok'])

    def _run_blk(self):
        grp = self.blk_grp.get().strip()
        if not grp or grp.startswith("--"): messagebox.showwarning("Fehlt","Gruppe auswählen!"); return
        users = [l.strip() for l in self.blk_users.get('1.0',tk.END).strip().split("\n") if l.strip() and '@' in l.strip()]
        if not users: messagebox.showwarning("Fehlt","Benutzer eingeben!"); return
        ge = self._ge(grp); act = self.blk_act.get()
        adding = "Hinzufügen" in act
        if not messagebox.askyesno("Bulk",f"{'Hinzufügen' if adding else 'Entfernen'} von {len(users)} Benutzern?\n📋 {ge}"): return
        self.log(f"  🏷️ Bulk: {len(users)} Benutzer → {ge}",C['warn'])
        def do():
            ok_c, err_c = 0, 0
            is_unified = any(g['e']==ge for g in self.groups.get('teams',[]))
            for u in users:
                u = u.strip()
                if not u: continue
                if adding:
                    if is_unified: r,_,_ = self.ps.run(f'Add-UnifiedGroupLinks -Identity "{ge}" -LinkType Members -Links "{u}" -EA SilentlyContinue')
                    else: r,_,_ = self.ps.run(f'Add-DistributionGroupMember -Identity "{ge}" -Member "{u}" -EA SilentlyContinue')
                else:
                    if is_unified:
                        r,_,_ = self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Members -Links "{u}" -Confirm:$false -EA SilentlyContinue')
                        self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Owners -Links "{u}" -Confirm:$false -EA SilentlyContinue')
                    else: r,_,_ = self.ps.run(f'Remove-DistributionGroupMember -Identity "{ge}" -Member "{u}" -Confirm:$false -EA SilentlyContinue')
                if r: ok_c += 1
                else: err_c += 1
            lines = [f"✅ {ok_c} erfolgreich",f"❌ {err_c} fehlgeschlagen" if err_c else ""]
            self.root.after(0, lambda: [self._settxt(self.blk_t,"\n".join(l for l in lines if l)),
                                         self.log(f"  🏷️ Bulk fertig: {ok_c}✅ {err_c}❌",C['ok'] if err_c==0 else C['warn']),
                                         messagebox.showinfo("Bulk",f"{ok_c} OK, {err_c} Fehler")])
        threading.Thread(target=do, daemon=True).start()

    # ── Allgemein ───────────────────────────────────────────
    def _settxt(self, widget, txt):
        widget.configure(state=tk.NORMAL); widget.delete('1.0',tk.END)
        widget.insert(tk.END, txt); widget.configure(state=tk.DISABLED)

    def _done(self, errs, word):
        if errs:
            for e in errs: self.log(f"  ❌ {e}",C['err'])
            messagebox.showerror("Fehler","Siehe Protokoll.")
        else: messagebox.showinfo("OK",f"✅ Erfolgreich {word}!")

    def log(self, msg, color=None):
        ts = datetime.now().strftime("%H:%M:%S")
        tag = {C['ok']:'ok',C['err']:'err',C['warn']:'warn',C['accent']:'acc',C['dim']:'dim'}.get(color,'info')
        self.log_t.configure(state=tk.NORMAL)
        self.log_t.insert(tk.END,f"[{ts}] ",'info'); self.log_t.insert(tk.END,f"{msg}\n",tag)
        self.log_t.see(tk.END); self.log_t.configure(state=tk.DISABLED)

    def cleanup(self):
        try: self.ps.run("Disconnect-ExchangeOnline -Confirm:$false",10)
        except: pass
        self.ps.stop()


def main():
    root = tk.Tk()
    try: root.iconbitmap('exchange.ico')
    except: pass
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", lambda: [app.cleanup(), root.destroy()] if messagebox.askokcancel("Beenden","Beenden?") else None)
    root.mainloop()

if __name__ == "__main__":
    main()