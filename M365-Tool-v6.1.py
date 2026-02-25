#!/usr/bin/env python3
"""M365 Admin Tool v6.1 — Kaulich IT Systems GmbH — Sidebar-Navigation, 12 Module
   v6.1: Modul-Vorabprüfung, Scrollbar-Fix, PW-Reset via Graph"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess, threading, json, queue, time, os, csv
from datetime import datetime

C = {
    'bg':'#1a1b26','sidebar':'#16161e','panel':'#1f2028','input':'#282a36','hdr':'#12121a',
    'accent':'#7aa2f7','ok':'#9ece6a','warn':'#e0af68','err':'#f7768e',
    'txt':'#c0caf5','dim':'#565f89','muted':'#414868','brd':'#2e3044',
    'purple':'#bb9af7','cyan':'#7dcfff','orange':'#ff9e64','green':'#73daca',
    'sb_hover':'#1f2335','sb_active':'#292e42',
}
SIDEBAR = [
    ("POSTFACH", [('postfach','📧','Berechtigungen'),('sharedmb','📫','Shared erstellen'),
                   ('forwarding','📬','Weiterleitungen'),('audit','🔍','Audit')]),
    ("GRUPPEN", [('teams','👥','Teams / M365'),('verteiler','📨','Verteilerlisten'),
                  ('security','🔒','Sicherheit'),('bulk','🏷️','Bulk-Aktionen')]),
    ("BENUTZER", [('userinfo','👤','Benutzer-Info'),('licenses','📊','Lizenzen'),
                   ('offboarding','🚪','Offboarding')]),
    ("EXPORT", [('csvexport','📋','CSV-Export')]),
]
OB_STEPS = [
    ('sign_in','🔒 Anmeldung blockieren'),('reset_pw','🔑 Passwort zurücksetzen'),
    ('remove_groups','👥 Gruppen entfernen'),('remove_licenses','📋 Lizenzen entziehen'),
    ('convert_shared','📧 → Shared Mailbox'),('set_ooo','✈️ Abwesenheit'),
    ('fwd','📨 Weiterleitung'),('hide_gal','👻 GAL ausblenden'),
    ('disable_sync','📱 Protokolle aus'),('remove_delegates','🔓 Delegierungen'),
]
OB_N = {k:v.split(' ',1)[1] for k,v in OB_STEPS}

# ── Module die geprüft werden ───────────────────────────
REQUIRED_MODULES = [
    {
        'name': 'ExchangeOnlineManagement',
        'display': 'Exchange Online Management',
        'icon': '📧',
        'required': True,
        'description': 'Postfächer, Gruppen, Berechtigungen',
    },
    {
        'name': 'Microsoft.Graph',
        'display': 'Microsoft Graph',
        'icon': '📊',
        'required': True,
        'description': 'Lizenzen, Passwort-Reset, Benutzer-Mgmt',
    },
]

class PS:
    def __init__(self): self.proc=None; self.q=queue.Queue()
    def start(self):
        if self.proc: return
        self.proc=subprocess.Popen(["powershell","-NoLogo","-NoExit","-Command","-"],
            stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
            text=True,encoding='utf-8',errors='replace',bufsize=1)
        threading.Thread(target=self._rd,daemon=True).start()
        self._w('[Console]::OutputEncoding=[System.Text.Encoding]::UTF8')
    def _rd(self):
        while self.proc and self.proc.poll() is None:
            try:
                l=self.proc.stdout.readline()
                if l: self.q.put(l)
            except: break
    def _w(self,c):
        if self.proc and self.proc.poll() is None:
            self.proc.stdin.write(c+"\n"); self.proc.stdin.flush()
    def run(self,cmd,timeout=120):
        if not self.proc or self.proc.poll() is not None: return False,"","PS nicht aktiv"
        while not self.q.empty():
            try: self.q.get_nowait()
            except: break
        ts=int(time.time()*1000); em,erm=f"###E{ts}###",f"###X{ts}###"
        self._w(f'try{{{cmd}}}catch{{Write-Output "{erm}$($_.Exception.Message)"}}\nWrite-Output "{em}"')
        ol,el=[],[]
        t0=time.time()
        while True:
            if time.time()-t0>timeout: return False,"","Timeout"
            try:
                l=self.q.get(timeout=0.5).rstrip()
                if em in l: break
                elif erm in l: el.append(l.split(erm,1)[1])
                else: ol.append(l)
            except queue.Empty: continue
        return (not el),"\n".join(ol),"\n".join(el)
    def stop(self):
        if self.proc:
            try: self._w("exit"); self.proc.terminate()
            except: pass
            self.proc=None

class Btn(tk.Canvas):
    def __init__(self,parent,text,command=None,bg=C['accent'],fg='#fff',width=160,height=32,font_size=10,**kw):
        kw.pop('width',None);kw.pop('height',None)
        pbg=C['panel']
        try: pbg=parent.cget('bg')
        except: pass
        super().__init__(parent,highlightthickness=0,borderwidth=0,**kw)
        super().configure(width=width,height=height,bg=pbg)
        self._cmd=command;self._bg=bg;self._fg=fg
        self._hov=self._adj(bg,20);self._dis=C['muted']
        self.txt=text;self._bw=width;self._bh=height;self._fs=font_size;self._en=True
        self.after(10,lambda:self._draw(self._bg))
        self.bind('<Enter>',lambda e:self._draw(self._hov) if self._en else None)
        self.bind('<Leave>',lambda e:self._draw(self._bg if self._en else self._dis))
        self.bind('<Button-1>',lambda e:self._cmd() if self._en and self._cmd else None)
    def _adj(self,c,a):
        try: r,g,b=int(c[1:3],16),int(c[3:5],16),int(c[5:7],16); return f'#{min(255,r+a):02x}{min(255,g+a):02x}{min(255,b+a):02x}'
        except: return c
    def _draw(self,col):
        self.delete('all');r,w,h=5,self._bw,self._bh
        for x0,y0,x1,y1,st in [(0,0,r*2,r*2,90),(w-r*2,0,w,r*2,0),(0,h-r*2,r*2,h,180),(w-r*2,h-r*2,w,h,270)]:
            self.create_arc(x0,y0,x1,y1,start=st,extent=90,fill=col,outline=col)
        self.create_rectangle(r,0,w-r,h,fill=col,outline=col)
        self.create_rectangle(0,r,w,h-r,fill=col,outline=col)
        self.create_text(w/2,h/2,text=self.txt,fill=self._fg if self._en else C['dim'],font=('Segoe UI',self._fs,'bold'))
    def configure(self,**kw):
        if 'state' in kw: self._en=kw['state']!=tk.DISABLED;self._draw(self._bg if self._en else self._dis)
        if 'text' in kw: self.txt=kw['text'];self._draw(self._bg if self._en else self._dis)
        if 'bg' in kw: self._bg=kw['bg'];self._hov=self._adj(kw['bg'],20);self._draw(self._bg)

class App:
    def __init__(self, root):
        self.root=root
        self.root.title("M365 Admin Tool v6.1 — Kaulich IT Systems GmbH")
        self.root.geometry("1100x750"); self.root.minsize(1000,650)
        self.root.configure(bg=C['bg'])
        self.ps=PS(); self.ps.start()
        self.connected=False; self.mailboxes=[]; self.groups={'teams':[],'verteiler':[],'security':[]}
        self.ob_report=[]; self._aud_data=[]; self._lic_data=[]
        self.sidebar_btns={}; self.pages={}
        self.mod_status = {}  # Modul-Status: name -> {installed, version}
        self._build()
        self.log("🚀 M365 Admin Tool v6.1",C['ok'])
        self.root.after(500,self._chk_all_modules)

    def _pj(self,o):
        if not o or not o.strip(): return []
        try:
            s1,s2=o.find('['),o.find('{')
            if s1==-1 and s2==-1: return []
            s=s1 if s1!=-1 and (s2==-1 or s1<s2) else s2
            d=json.loads(o[s:]); return [d] if isinstance(d,dict) else d
        except: return []
    def _ge(self,s): return s.split('<')[1].split('>')[0] if '<' in s and '>' in s else s

    # ── LAYOUT ───────────────────────────────────────────
    def _build(self):
        hdr=tk.Frame(self.root,bg=C['hdr'],height=44); hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr,text="⚡ M365 Admin Tool",font=('Segoe UI',14,'bold'),fg=C['accent'],bg=C['hdr']).pack(side=tk.LEFT,padx=12)
        cf=tk.Frame(hdr,bg=C['hdr']); cf.pack(side=tk.RIGHT,padx=12)
        self.conn_lbl=tk.Label(cf,text="⚫ Nicht verbunden",font=('Segoe UI',9),fg=C['dim'],bg=C['hdr'])
        self.conn_lbl.pack(side=tk.RIGHT,padx=(8,0))
        self.conn_btn=Btn(cf,"🔌 Verbinden",command=self._show_conn,bg=C['accent'],width=120,height=28,font_size=9)
        self.conn_btn.pack(side=tk.RIGHT)
        self.mod_lbl=tk.Label(cf,text="",font=('Segoe UI',8),fg=C['dim'],bg=C['hdr'])
        self.mod_lbl.pack(side=tk.RIGHT,padx=(0,8))

        body=tk.Frame(self.root,bg=C['bg']); body.pack(fill=tk.BOTH,expand=True)
        self.sidebar_frame=tk.Frame(body,bg=C['sidebar'],width=180)
        self.sidebar_frame.pack(side=tk.LEFT,fill=tk.Y); self.sidebar_frame.pack_propagate(False)
        for gn,items in SIDEBAR:
            tk.Label(self.sidebar_frame,text=gn,font=('Segoe UI',8,'bold'),fg=C['muted'],bg=C['sidebar']).pack(fill=tk.X,padx=12,pady=(10,2))
            for key,icon,label in items:
                btn=tk.Frame(self.sidebar_frame,bg=C['sidebar'],cursor='hand2'); btn.pack(fill=tk.X,padx=6,pady=1)
                lbl=tk.Label(btn,text=f" {icon}  {label}",font=('Segoe UI',10),fg=C['dim'],bg=C['sidebar'],anchor=tk.W,cursor='hand2')
                lbl.pack(fill=tk.X,padx=6,pady=4)
                btn.bind('<Button-1>',lambda e,k=key:self._nav(k)); lbl.bind('<Button-1>',lambda e,k=key:self._nav(k))
                btn.bind('<Enter>',lambda e,b=btn,l=lbl:[b.configure(bg=C['sb_hover']),l.configure(bg=C['sb_hover'])])
                btn.bind('<Leave>',lambda e,b=btn,l=lbl,k=key:self._sb_leave(b,l,k))
                self.sidebar_btns[key]=(btn,lbl)
        tk.Label(self.sidebar_frame,text="Kaulich IT\nv6.1",font=('Segoe UI',8),fg=C['muted'],bg=C['sidebar'],justify=tk.CENTER).pack(side=tk.BOTTOM,pady=8)

        right=tk.Frame(body,bg=C['bg']); right.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        self._canvas=tk.Canvas(right,bg=C['bg'],highlightthickness=0)
        sb=tk.Scrollbar(right,orient=tk.VERTICAL,command=self._canvas.yview)
        self.content=tk.Frame(self._canvas,bg=C['bg'])
        self.content.bind('<Configure>',lambda e:self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._cw=self._canvas.create_window((0,0),window=self.content,anchor="nw")
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.pack(side=tk.LEFT,fill=tk.BOTH,expand=True,padx=(8,0),pady=8)
        sb.pack(side=tk.RIGHT,fill=tk.Y)
        self._canvas.bind('<Configure>',lambda e:self._canvas.itemconfig(self._cw,width=e.width))

        # ── Scrollbar-Fix: nur scrollen wenn Maus über dem Canvas ──
        self._canvas.bind('<Enter>', self._bind_mousewheel)
        self._canvas.bind('<Leave>', self._unbind_mousewheel)

        self._b_conn()
        self._b_preflight()  # NEU: Vorab-Prüfung Seite
        self._b_postfach(); self._b_grp('teams','Teams / M365 Gruppe'); self._b_grp('verteiler','Verteilerliste')
        self._b_grp('security','Sicherheitsgruppe'); self._b_offboarding(); self._b_userinfo()
        self._b_licenses(); self._b_sharedmb(); self._b_forwarding(); self._b_audit()
        self._b_csvexport(); self._b_bulk()

        self.log_frame=tk.Frame(self.content,bg=C['bg']); self.log_frame.pack(fill=tk.X,pady=(10,0))
        tk.Label(self.log_frame,text="📋 Protokoll",font=('Segoe UI',10,'bold'),fg=C['accent'],bg=C['bg']).pack(anchor=tk.W,pady=(0,3))
        lf=tk.Frame(self.log_frame,bg=C['panel'],highlightbackground=C['brd'],highlightthickness=1); lf.pack(fill=tk.X)
        self.log_t=tk.Text(lf,height=6,font=('Consolas',9),bg=C['input'],fg=C['txt'],relief=tk.FLAT,wrap=tk.WORD,highlightthickness=0)
        lsb=tk.Scrollbar(lf,orient=tk.VERTICAL,command=self.log_t.yview); self.log_t.configure(yscrollcommand=lsb.set)
        self.log_t.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=6,pady=6); lsb.pack(side=tk.RIGHT,fill=tk.Y,padx=(0,2),pady=6)
        for t,c2 in [('ok',C['ok']),('err',C['err']),('warn',C['warn']),('info',C['txt']),('acc',C['accent']),('dim',C['dim'])]:
            self.log_t.tag_configure(t,foreground=c2)
        st=ttk.Style(); st.theme_use('clam')
        st.configure('TCombobox',fieldbackground=C['input'],background=C['panel'],foreground=C['txt'],arrowcolor=C['txt'],bordercolor=C['brd'])
        self._nav('_preflight')

    # ── Scrollbar-Fix Methoden ───────────────────────────
    def _bind_mousewheel(self, event):
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self._canvas.bind_all("<Button-4>", self._on_mousewheel_linux)
        self._canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _unbind_mousewheel(self, event):
        self._canvas.unbind_all("<MouseWheel>")
        self._canvas.unbind_all("<Button-4>")
        self._canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        # Nicht scrollen wenn Content kleiner als Canvas
        bbox = self._canvas.bbox("all")
        if bbox:
            content_height = bbox[3] - bbox[1]
            canvas_height = self._canvas.winfo_height()
            if content_height <= canvas_height:
                return
        self._canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_mousewheel_linux(self, event):
        bbox = self._canvas.bbox("all")
        if bbox:
            content_height = bbox[3] - bbox[1]
            canvas_height = self._canvas.winfo_height()
            if content_height <= canvas_height:
                return
        if event.num == 4:
            self._canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            self._canvas.yview_scroll(3, "units")

    def _nav(self,key):
        for k,(b,l) in self.sidebar_btns.items():
            if k==key: b.configure(bg=C['sb_active']);l.configure(bg=C['sb_active'],fg=C['accent'])
            else: b.configure(bg=C['sidebar']);l.configure(bg=C['sidebar'],fg=C['dim'])
        for f in self.pages.values(): f.pack_forget()
        if key in self.pages:
            self.pages[key].pack(fill=tk.X,before=self.log_frame)
        # Nach Navigation: Canvas zum Anfang scrollen
        self._canvas.yview_moveto(0)
    def _sb_leave(self,btn,lbl,key):
        if key in self.pages and self.pages[key].winfo_ismapped():
            btn.configure(bg=C['sb_active']);lbl.configure(bg=C['sb_active']); return
        btn.configure(bg=C['sidebar']);lbl.configure(bg=C['sidebar'])
    def _show_conn(self): self._nav('_conn')

    # ── UI Helpers ───────────────────────────────────────
    def _page(self,key,title,icon=""):
        f=tk.Frame(self.content,bg=C['bg']); self.pages[key]=f
        tk.Label(f,text=f"{icon}  {title}" if icon else title,font=('Segoe UI',14,'bold'),fg=C['txt'],bg=C['bg']).pack(anchor=tk.W,pady=(0,8))
        return f
    def _card(self,p):
        f=tk.Frame(p,bg=C['panel'],highlightbackground=C['brd'],highlightthickness=1); f.pack(fill=tk.X,pady=(0,8)); return f
    def _combo(self,p,label):
        r=tk.Frame(p,bg=C['panel']); r.pack(fill=tk.X,padx=12,pady=(6,0))
        tk.Label(r,text=label,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],width=15,anchor=tk.W).pack(side=tk.LEFT)
        cb=ttk.Combobox(r,width=55,state="disabled",font=('Segoe UI',9))
        cb.pack(side=tk.LEFT,fill=tk.X,expand=True); cb.set("— Erst verbinden —"); return cb
    def _entry(self,p,label):
        r=tk.Frame(p,bg=C['panel']); r.pack(fill=tk.X,padx=12,pady=(6,0))
        tk.Label(r,text=label,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],width=15,anchor=tk.W).pack(side=tk.LEFT)
        e=tk.Entry(r,font=('Segoe UI',10),bg=C['input'],fg=C['txt'],insertbackground=C['txt'],relief=tk.FLAT,
                   highlightthickness=1,highlightbackground=C['brd'],highlightcolor=C['accent'])
        e.pack(side=tk.LEFT,fill=tk.X,expand=True,ipady=3); return e
    def _search(self,p,fn):
        r=tk.Frame(p,bg=C['panel']); r.pack(fill=tk.X,padx=12,pady=(6,4))
        tk.Label(r,text="🔍",font=('Segoe UI',10),fg=C['dim'],bg=C['panel']).pack(side=tk.LEFT,padx=(0,4))
        e=tk.Entry(r,font=('Segoe UI',9),bg=C['input'],fg=C['txt'],insertbackground=C['txt'],relief=tk.FLAT,
                   highlightthickness=1,highlightbackground=C['brd'],highlightcolor=C['accent'])
        e.pack(side=tk.LEFT,fill=tk.X,expand=True,ipady=2); e.bind('<KeyRelease>',fn); return e
    def _txtbox(self,p,h=5):
        f=tk.Frame(p,bg=C['panel']); f.pack(fill=tk.X,padx=12,pady=(4,8))
        t=tk.Text(f,height=h,font=('Consolas',9),bg=C['input'],fg=C['txt'],relief=tk.FLAT,wrap=tk.WORD,
                  state=tk.DISABLED,highlightthickness=1,highlightbackground=C['brd'])
        sb=tk.Scrollbar(f,orient=tk.VERTICAL,command=t.yview); t.configure(yscrollcommand=sb.set)
        t.pack(side=tk.LEFT,fill=tk.X,expand=True); sb.pack(side=tk.RIGHT,fill=tk.Y); return t
    def _btnrow(self,p):
        r=tk.Frame(p,bg=C['panel']); r.pack(fill=tk.X,padx=12,pady=(4,10)); return r
    def _settxt(self,w,txt):
        w.configure(state=tk.NORMAL);w.delete('1.0',tk.END);w.insert(tk.END,txt);w.configure(state=tk.DISABLED)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  VORAB-PRÜFUNG (Preflight)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    def _b_preflight(self):
        p = self._page('_preflight', 'System-Voraussetzungen', '🔧')
        cd = self._card(p)
        tk.Label(cd, text="Folgende PowerShell-Module werden benötigt:",
                 font=('Segoe UI', 10), fg=C['txt'], bg=C['panel']).pack(fill=tk.X, padx=12, pady=(8, 4))

        self._pf_frame = tk.Frame(cd, bg=C['panel'])
        self._pf_frame.pack(fill=tk.X, padx=12, pady=(0, 4))
        self._pf_labels = {}

        for i, mod in enumerate(REQUIRED_MODULES):
            row = tk.Frame(self._pf_frame, bg=C['panel'])
            row.pack(fill=tk.X, pady=2)

            icon_lbl = tk.Label(row, text="⏳", font=('Segoe UI', 12), fg=C['warn'], bg=C['panel'], width=3)
            icon_lbl.pack(side=tk.LEFT)

            info = tk.Frame(row, bg=C['panel'])
            info.pack(side=tk.LEFT, fill=tk.X, expand=True)

            name_lbl = tk.Label(info, text=f"{mod['icon']}  {mod['display']}",
                                font=('Segoe UI', 10, 'bold'), fg=C['txt'], bg=C['panel'], anchor=tk.W)
            name_lbl.pack(anchor=tk.W)

            desc_lbl = tk.Label(info, text=mod['description'],
                                font=('Segoe UI', 8), fg=C['dim'], bg=C['panel'], anchor=tk.W)
            desc_lbl.pack(anchor=tk.W)

            status_lbl = tk.Label(row, text="Prüfe...", font=('Segoe UI', 9), fg=C['warn'], bg=C['panel'])
            status_lbl.pack(side=tk.RIGHT, padx=(0, 8))

            self._pf_labels[mod['name']] = (icon_lbl, status_lbl)

        # Zusätzliche Prüfungen
        sep = tk.Frame(cd, bg=C['brd'], height=1)
        sep.pack(fill=tk.X, padx=12, pady=(8, 4))

        tk.Label(cd, text="Zusätzliche Prüfungen:",
                 font=('Segoe UI', 10), fg=C['txt'], bg=C['panel']).pack(fill=tk.X, padx=12, pady=(4, 4))

        extra_frame = tk.Frame(cd, bg=C['panel'])
        extra_frame.pack(fill=tk.X, padx=12, pady=(0, 4))

        for check_name, check_label in [('powershell', '⚡ PowerShell-Version'),
                                          ('execution_policy', '🔐 Execution Policy'),
                                          ('tls', '🔒 TLS 1.2')]:
            row = tk.Frame(extra_frame, bg=C['panel'])
            row.pack(fill=tk.X, pady=2)
            icon_lbl = tk.Label(row, text="⏳", font=('Segoe UI', 12), fg=C['warn'], bg=C['panel'], width=3)
            icon_lbl.pack(side=tk.LEFT)
            tk.Label(row, text=check_label, font=('Segoe UI', 10), fg=C['txt'], bg=C['panel'], anchor=tk.W).pack(side=tk.LEFT)
            status_lbl = tk.Label(row, text="Prüfe...", font=('Segoe UI', 9), fg=C['warn'], bg=C['panel'])
            status_lbl.pack(side=tk.RIGHT, padx=(0, 8))
            self._pf_labels[check_name] = (icon_lbl, status_lbl)

        # Ergebnis-Bereich
        self._pf_result = tk.Label(cd, text="", font=('Segoe UI', 11, 'bold'), fg=C['warn'], bg=C['panel'])
        self._pf_result.pack(fill=tk.X, padx=12, pady=(8, 4))

        br = self._btnrow(cd)
        self._pf_install_btn = Btn(br, "📦 Fehlende Module installieren",
                                    command=self._install_missing, bg=C['warn'], fg='#1a1b26', width=250)
        self._pf_install_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._pf_install_btn.configure(state=tk.DISABLED)

        Btn(br, "🔄 Erneut prüfen", command=self._chk_all_modules, bg=C['input'], width=150).pack(side=tk.LEFT, padx=(0, 8))
        self._pf_continue_btn = Btn(br, "▶️ Weiter zur Verbindung",
                                     command=lambda: self._nav('_conn'), bg=C['ok'], fg='#1a1b26', width=200)
        self._pf_continue_btn.pack(side=tk.LEFT)
        self._pf_continue_btn.configure(state=tk.DISABLED)

    def _chk_all_modules(self):
        """Alle Module und Voraussetzungen prüfen"""
        self.log("🔧 Starte Vorab-Prüfung...", C['warn'])
        self._pf_result.configure(text="⏳ Prüfe Voraussetzungen...", fg=C['warn'])
        self._pf_install_btn.configure(state=tk.DISABLED)
        self._pf_continue_btn.configure(state=tk.DISABLED)

        # Reset alle Labels
        for name, (icon_lbl, status_lbl) in self._pf_labels.items():
            icon_lbl.configure(text="⏳", fg=C['warn'])
            status_lbl.configure(text="Prüfe...", fg=C['warn'])

        def do():
            all_ok = True
            missing = []

            # 1. PowerShell-Version
            ok, o, _ = self.ps.run('$PSVersionTable.PSVersion.ToString()', 15)
            v = o.strip().split("\n")[0] if ok and o.strip() else "?"
            ps_ok = ok and v != "?"
            self.root.after(0, lambda: self._pf_set('powershell', ps_ok, f"v{v}" if ps_ok else "Fehler"))
            if not ps_ok: all_ok = False

            # 2. Execution Policy
            ok, o, _ = self.ps.run('Get-ExecutionPolicy -Scope CurrentUser', 15)
            policy = o.strip().split("\n")[0] if ok else "?"
            ep_ok = policy.lower() in ['remotesigned', 'unrestricted', 'bypass', 'allsigned']
            self.root.after(0, lambda: self._pf_set('execution_policy', ep_ok,
                                                     f"{policy}" if ep_ok else f"{policy} — RemoteSigned empfohlen"))
            if not ep_ok:
                # Versuche automatisch zu setzen
                self.ps.run('Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -EA SilentlyContinue', 15)
                ok2, o2, _ = self.ps.run('Get-ExecutionPolicy -Scope CurrentUser', 15)
                policy2 = o2.strip().split("\n")[0] if ok2 else "?"
                ep_ok2 = policy2.lower() in ['remotesigned', 'unrestricted', 'bypass', 'allsigned']
                self.root.after(0, lambda: self._pf_set('execution_policy', ep_ok2,
                                                         f"{policy2} (auto-gesetzt)" if ep_ok2 else f"{policy2} — bitte manuell setzen"))
                if not ep_ok2: all_ok = False

            # 3. TLS 1.2
            ok, o, _ = self.ps.run('[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Write-Output "TLS12OK"', 15)
            tls_ok = "TLS12OK" in o
            self.root.after(0, lambda: self._pf_set('tls', tls_ok, "TLS 1.2 aktiv" if tls_ok else "Fehler"))
            if not tls_ok: all_ok = False

            # 4. PowerShell-Module prüfen
            for mod in REQUIRED_MODULES:
                name = mod['name']
                self.root.after(0, lambda n=name: self.log(f"  🔍 Prüfe {n}...", C['dim']))
                ok, o, _ = self.ps.run(
                    f'$m=Get-Module -ListAvailable -Name {name};if($m){{Write-Output "MOD_OK:$($m.Version|Select -First 1)"}}else{{Write-Output "MOD_MISS"}}', 30)

                if "MOD_OK:" in o:
                    ver = o.split("MOD_OK:")[1].strip().split("\n")[0]
                    self.mod_status[name] = {'installed': True, 'version': ver}
                    self.root.after(0, lambda n=name, v=ver: [
                        self._pf_set(n, True, f"v{v}"),
                        self.log(f"  ✅ {n} v{v}", C['ok'])
                    ])
                else:
                    self.mod_status[name] = {'installed': False, 'version': None}
                    missing.append(mod)
                    self.root.after(0, lambda n=name: [
                        self._pf_set(n, False, "Nicht installiert"),
                        self.log(f"  ❌ {n} fehlt!", C['err'])
                    ])
                    all_ok = False

            # Update Header-Label
            exo = self.mod_status.get('ExchangeOnlineManagement', {})
            graph = self.mod_status.get('Microsoft.Graph', {})
            hdr_parts = []
            if exo.get('installed'): hdr_parts.append(f"📧v{exo['version']}")
            else: hdr_parts.append("📧❌")
            if graph.get('installed'): hdr_parts.append(f"📊v{graph['version']}")
            else: hdr_parts.append("📊❌")
            hdr_text = " ".join(hdr_parts)

            self.root.after(0, lambda: self.mod_lbl.configure(
                text=hdr_text, fg=C['ok'] if all_ok else C['err']))

            # Ergebnis
            if all_ok:
                self.root.after(0, lambda: [
                    self._pf_result.configure(text="✅ Alle Voraussetzungen erfüllt!", fg=C['ok']),
                    self._pf_continue_btn.configure(state=tk.NORMAL),
                    self.log("✅ Vorab-Prüfung bestanden!", C['ok']),
                ])
            else:
                self.root.after(0, lambda: [
                    self._pf_result.configure(text=f"⚠️ {len(missing)} Modul(e) fehlen", fg=C['err']),
                    self._pf_install_btn.configure(state=tk.NORMAL),
                    self._pf_continue_btn.configure(state=tk.NORMAL),  # trotzdem erlauben, aber warnen
                    self.log(f"⚠️ {len(missing)} fehlende Module!", C['err']),
                ])

        threading.Thread(target=do, daemon=True).start()

    def _pf_set(self, name, ok, text):
        """Preflight-Label aktualisieren"""
        if name in self._pf_labels:
            icon_lbl, status_lbl = self._pf_labels[name]
            icon_lbl.configure(text="✅" if ok else "❌", fg=C['ok'] if ok else C['err'])
            status_lbl.configure(text=text, fg=C['ok'] if ok else C['err'])

    def _install_missing(self):
        """Fehlende Module installieren"""
        missing = [mod for mod in REQUIRED_MODULES if not self.mod_status.get(mod['name'], {}).get('installed')]
        if not missing:
            messagebox.showinfo("OK", "Alle Module bereits installiert!")
            return
        names = ", ".join(m['display'] for m in missing)
        if not messagebox.askyesno("Installieren", f"Folgende Module installieren?\n\n{names}\n\nDies kann einige Minuten dauern."):
            return

        self._pf_install_btn.configure(state=tk.DISABLED)
        self._pf_result.configure(text="📦 Installation läuft...", fg=C['warn'])
        self.log("📦 Installiere fehlende Module...", C['warn'])

        def do():
            # NuGet Provider sicherstellen
            self.ps.run('Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -EA SilentlyContinue|Out-Null', 60)

            for mod in missing:
                name = mod['name']
                self.root.after(0, lambda n=name: [
                    self._pf_set(n, None, "Installiere..."),
                    self.log(f"  📦 Installiere {n}...", C['warn'])
                ])
                # Labels auf "Installiere" setzen
                self.root.after(0, lambda n=name: self._pf_labels[n][0].configure(text="⏳", fg=C['warn']))
                self.root.after(0, lambda n=name: self._pf_labels[n][1].configure(text="Installiere...", fg=C['warn']))

                ok, o, e = self.ps.run(
                    f'Install-Module {name} -Scope CurrentUser -Force -AllowClobber -EA Stop;'
                    f'$m=Get-Module -ListAvailable -Name {name};'
                    f'if($m){{Write-Output "INST_OK:$($m.Version|Select -First 1)"}}else{{Write-Output "INST_FAIL"}}',
                    300)

                if "INST_OK:" in o:
                    ver = o.split("INST_OK:")[1].strip().split("\n")[0]
                    self.mod_status[name] = {'installed': True, 'version': ver}
                    self.root.after(0, lambda n=name, v=ver: [
                        self._pf_set(n, True, f"v{v} (neu installiert)"),
                        self.log(f"  ✅ {n} v{v} installiert", C['ok'])
                    ])
                else:
                    self.root.after(0, lambda n=name, err=e: [
                        self._pf_set(n, False, f"Installation fehlgeschlagen"),
                        self.log(f"  ❌ {n}: {err}", C['err'])
                    ])

            # Erneut prüfen
            self.root.after(500, self._chk_all_modules)

        threading.Thread(target=do, daemon=True).start()

    # ── MODULE ───────────────────────────────────────────
    def _b_conn(self):
        p=self._page('_conn','Verbindung','🔐'); cd=self._card(p)
        self.admin_e=self._entry(cd,"Admin-UPN:")
        br=self._btnrow(cd)
        Btn(br,"🔌 Verbinden",command=self._connect,bg=C['accent'],width=150).pack(side=tk.LEFT,padx=(0,8))
        Btn(br,"🔌 Trennen",command=self.disconnect,bg=C['input'],width=120).pack(side=tk.LEFT)

    def _b_postfach(self):
        p=self._page('postfach','Postfachberechtigungen','📧'); cd=self._card(p)
        self.mb_t=self._combo(cd,"Ziel-Postfach:"); self.mb_u=self._combo(cd,"Benutzer:")
        self.mb_s=self._search(cd,self._fmb)
        fr=tk.Frame(cd,bg=C['panel']); fr.pack(fill=tk.X,padx=12,pady=(4,0))
        self.mb_ft=tk.StringVar(value="all")
        for t,v in [("Alle","all"),("👤 User","user"),("👥 Shared","shared")]:
            tk.Radiobutton(fr,text=t,variable=self.mb_ft,value=v,font=('Segoe UI',9),fg=C['txt'],
                           bg=C['panel'],selectcolor=C['input'],activebackground=C['panel'],command=self._amb).pack(side=tk.LEFT,padx=(0,10))
        pr=tk.Frame(cd,bg=C['panel']); pr.pack(fill=tk.X,padx=12,pady=(6,0))
        self.fa_v=tk.BooleanVar(value=True);self.am_v=tk.BooleanVar(value=True);self.sa_v=tk.BooleanVar(value=True)
        tk.Checkbutton(pr,text="📂 Vollzugriff",variable=self.fa_v,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(side=tk.LEFT)
        tk.Checkbutton(pr,text="🔗 AutoMap",variable=self.am_v,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(side=tk.LEFT,padx=(12,0))
        tk.Checkbutton(pr,text="✉️ Senden als",variable=self.sa_v,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(side=tk.LEFT,padx=(12,0))
        br=self._btnrow(cd)
        Btn(br,"✅ Hinzufügen",command=self._add_mb,bg=C['ok'],fg='#1a1b26',width=155).pack(side=tk.LEFT,padx=(0,8))
        Btn(br,"❌ Entfernen",command=self._rem_mb,bg=C['err'],width=155).pack(side=tk.LEFT)

    def _b_grp(self,key,label):
        ic={'teams':'👥','verteiler':'📨','security':'🔒'}[key]
        p=self._page(key,label,ic); cd=self._card(p)
        gc=self._combo(cd,"Gruppe:"); uc=self._combo(cd,"Benutzer:")
        se=self._search(cd,lambda e,k=key:self._fgrp(k))
        rv=tk.StringVar(value="Member")
        if key=='teams':
            rr=tk.Frame(cd,bg=C['panel']); rr.pack(fill=tk.X,padx=12,pady=(4,0))
            for t,v in [("👤 Mitglied","Member"),("👑 Besitzer","Owner")]:
                tk.Radiobutton(rr,text=t,variable=rv,value=v,font=('Segoe UI',9),fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(side=tk.LEFT,padx=(0,10))
        mt=self._txtbox(cd,4); br=self._btnrow(cd)
        Btn(br,"📋 Mitglieder",command=lambda k=key:self._smem(k),bg=C['input'],width=125).pack(side=tk.LEFT,padx=(0,8))
        Btn(br,"✅ Hinzufügen",command=lambda k=key:self._agrp(k),bg=C['ok'],fg='#1a1b26',width=125).pack(side=tk.LEFT,padx=(0,8))
        Btn(br,"❌ Entfernen",command=lambda k=key:self._rgrp(k),bg=C['err'],width=125).pack(side=tk.LEFT)
        for a,v2 in [('gc',gc),('uc',uc),('se',se),('rv',rv),('mt',mt)]: setattr(self,f'{key}_{a}',v2)

    def _b_offboarding(self):
        p=self._page('offboarding','Benutzer-Offboarding','🚪'); cd=self._card(p)
        self.ob_u=self._combo(cd,"Benutzer:"); self.ob_s=self._search(cd,self._fob)
        wr=tk.Frame(cd,bg=C['panel']); wr.pack(fill=tk.X,padx=12,pady=(0,4))
        tk.Label(wr,text="⚠️ Teilweise irreversible Aktionen!",font=('Segoe UI',9,'bold'),fg=C['err'],bg=C['panel']).pack(anchor=tk.W)
        cbf=tk.Frame(cd,bg=C['panel']); cbf.pack(fill=tk.X,padx=12,pady=(0,0))
        self.ob_v={}
        for i,(k,l) in enumerate(OB_STEPS):
            v=tk.BooleanVar(value=True); self.ob_v[k]=v
            tk.Checkbutton(cbf,text=l,variable=v,font=('Segoe UI',9),fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).grid(row=i//2,column=i%2,sticky=tk.W,padx=(0,10),pady=1)
        tr=tk.Frame(cd,bg=C['panel']); tr.pack(fill=tk.X,padx=12,pady=(4,0))
        Btn(tr,"Alle an",command=lambda:[v.set(True) for v in self.ob_v.values()],bg=C['input'],width=65,height=24,font_size=8).pack(side=tk.LEFT,padx=(0,4))
        Btn(tr,"Alle ab",command=lambda:[v.set(False) for v in self.ob_v.values()],bg=C['input'],width=65,height=24,font_size=8).pack(side=tk.LEFT)
        tk.Label(cd,text="✈️ Abwesenheitsnachricht:",font=('Segoe UI',9),fg=C['dim'],bg=C['panel']).pack(fill=tk.X,padx=12,pady=(6,0))
        of=tk.Frame(cd,bg=C['panel']); of.pack(fill=tk.X,padx=12,pady=(2,0))
        self.ob_ooo=tk.Text(of,height=2,font=('Segoe UI',9),bg=C['input'],fg=C['txt'],relief=tk.FLAT,wrap=tk.WORD,highlightthickness=1,highlightbackground=C['brd'])
        self.ob_ooo.pack(fill=tk.X); self.ob_ooo.insert('1.0','Dieser Mitarbeiter ist nicht mehr im Unternehmen. Bitte wenden Sie sich an helpdesk@kaulich-it.de')
        self.ob_fwd=self._combo(cd,"Weiterleitung an:")
        pf=tk.Frame(cd,bg=C['panel']); pf.pack(fill=tk.X,padx=12,pady=(6,0))
        self.ob_pl=tk.Label(pf,text="",font=('Segoe UI',9),fg=C['dim'],bg=C['panel']); self.ob_pl.pack(anchor=tk.W)
        self.ob_pb=ttk.Progressbar(pf,mode='determinate'); self.ob_pb.pack(fill=tk.X,pady=(2,0))
        self.ob_rt=self._txtbox(cd,5); br=self._btnrow(cd)
        self.ob_rb=Btn(br,"🚪 Offboarding starten",command=self._run_ob,bg=C['err'],width=180); self.ob_rb.pack(side=tk.LEFT,padx=(0,8))
        self.ob_eb=Btn(br,"💾 Bericht",command=self._exp_ob,bg=C['accent'],width=120); self.ob_eb.pack(side=tk.LEFT)
        self.ob_eb.configure(state=tk.DISABLED)

    def _b_userinfo(self):
        p=self._page('userinfo','Benutzer-Info','👤'); cd=self._card(p)
        self.ui_u=self._combo(cd,"Benutzer:"); self.ui_s=self._search(cd,self._fui)
        br=self._btnrow(cd); Btn(br,"🔍 Info laden",command=self._load_ui,bg=C['accent'],width=140).pack(side=tk.LEFT)
        self.ui_t=self._txtbox(cd,14)

    def _b_licenses(self):
        p=self._page('licenses','Lizenz-Übersicht','📊'); cd=self._card(p)
        # Warnhinweis wenn Graph fehlt
        self._lic_warn = tk.Label(cd, text="", font=('Segoe UI', 9), fg=C['err'], bg=C['panel'])
        self._lic_warn.pack(fill=tk.X, padx=12, pady=(4, 0))
        br=self._btnrow(cd)
        Btn(br,"📊 Laden",command=self._load_lic,bg=C['accent'],width=140).pack(side=tk.LEFT,padx=(0,8))
        Btn(br,"💾 CSV",command=self._exp_lic,bg=C['ok'],fg='#1a1b26',width=110).pack(side=tk.LEFT)
        self.lic_t=self._txtbox(cd,14)

    def _b_sharedmb(self):
        p=self._page('sharedmb','Shared Mailbox erstellen','📫'); cd=self._card(p)
        self.sm_name=self._entry(cd,"Name:"); self.sm_email=self._entry(cd,"E-Mail:")
        self.sm_disp=self._entry(cd,"Anzeigename:")
        tk.Label(cd,text="Berechtigungen (optional):",font=('Segoe UI',9,'bold'),fg=C['txt'],bg=C['panel']).pack(fill=tk.X,padx=12,pady=(8,0))
        self.sm_perm=self._combo(cd,"Vollzugriff für:")
        br=self._btnrow(cd); Btn(br,"📧 Erstellen",command=self._create_sm,bg=C['ok'],fg='#1a1b26',width=150).pack(side=tk.LEFT)

    def _b_forwarding(self):
        p=self._page('forwarding','Mail-Weiterleitungen','📬'); cd=self._card(p)
        br1=self._btnrow(cd); Btn(br1,"📬 Alle laden",command=self._load_fwd,bg=C['accent'],width=160).pack(side=tk.LEFT)
        self.fwd_t=self._txtbox(cd,6)
        cd2=self._card(p)
        tk.Label(cd2,text="Weiterleitung setzen / entfernen",font=('Segoe UI',10,'bold'),fg=C['txt'],bg=C['panel']).pack(fill=tk.X,padx=12,pady=(8,0))
        self.fwd_src=self._combo(cd2,"Postfach:"); self.fwd_dst=self._combo(cd2,"Weiterleitung an:")
        self.fwd_keep=tk.BooleanVar(value=True)
        kr=tk.Frame(cd2,bg=C['panel']); kr.pack(fill=tk.X,padx=12,pady=(4,0))
        tk.Checkbutton(kr,text="📥 Kopie im Postfach behalten",variable=self.fwd_keep,font=('Segoe UI',9),fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).pack(anchor=tk.W)
        br2=self._btnrow(cd2)
        Btn(br2,"✅ Setzen",command=self._set_fwd,bg=C['ok'],fg='#1a1b26',width=130).pack(side=tk.LEFT,padx=(0,8))
        Btn(br2,"❌ Entfernen",command=self._rem_fwd,bg=C['err'],width=130).pack(side=tk.LEFT)

    def _b_audit(self):
        p=self._page('audit','Berechtigungs-Audit','🔍'); cd=self._card(p)
        self.aud_u=self._combo(cd,"Postfach:"); self.aud_s=self._search(cd,self._faud)
        br=self._btnrow(cd)
        Btn(br,"🔍 Audit",command=self._run_aud,bg=C['accent'],width=130).pack(side=tk.LEFT,padx=(0,8))
        Btn(br,"💾 CSV",command=self._exp_aud,bg=C['ok'],fg='#1a1b26',width=110).pack(side=tk.LEFT)
        self.aud_t=self._txtbox(cd,12)

    def _b_csvexport(self):
        p=self._page('csvexport','CSV-Export','📋'); cd=self._card(p)
        self.csv_opts={}
        cbf=tk.Frame(cd,bg=C['panel']); cbf.pack(fill=tk.X,padx=12,pady=(8,4))
        for i,(k,l) in enumerate([('users','👤 Alle Benutzer'),('shared','📧 Shared Mailboxen'),
                                    ('groups','👥 Alle Gruppen'),('licenses','📊 Lizenzen'),('forwarding','📬 Weiterleitungen')]):
            v=tk.BooleanVar(value=True); self.csv_opts[k]=v
            tk.Checkbutton(cbf,text=l,variable=v,font=('Segoe UI',10),fg=C['txt'],bg=C['panel'],selectcolor=C['input'],activebackground=C['panel']).grid(row=i//2,column=i%2,sticky=tk.W,padx=(0,20),pady=2)
        br=self._btnrow(cd); Btn(br,"📋 Exportieren",command=self._run_csv,bg=C['ok'],fg='#1a1b26',width=160).pack(side=tk.LEFT)

    def _b_bulk(self):
        p=self._page('bulk','Bulk-Aktionen','🏷️'); cd=self._card(p)
        tk.Label(cd,text="Aktion:",font=('Segoe UI',10,'bold'),fg=C['txt'],bg=C['panel']).pack(fill=tk.X,padx=12,pady=(8,0))
        self.blk_act=ttk.Combobox(cd,width=50,state="readonly",font=('Segoe UI',10),
                                   values=["➕ Zu Gruppe hinzufügen","➖ Aus Gruppe entfernen"])
        self.blk_act.pack(padx=12,anchor=tk.W,pady=(4,0)); self.blk_act.current(0)
        self.blk_grp=self._combo(cd,"Ziel-Gruppe:")
        tk.Label(cd,text="Benutzer (E-Mail pro Zeile):",font=('Segoe UI',9),fg=C['dim'],bg=C['panel']).pack(fill=tk.X,padx=12,pady=(6,0))
        uf=tk.Frame(cd,bg=C['panel']); uf.pack(fill=tk.X,padx=12,pady=(2,0))
        self.blk_users=tk.Text(uf,height=5,font=('Consolas',9),bg=C['input'],fg=C['txt'],relief=tk.FLAT,wrap=tk.WORD,highlightthickness=1,highlightbackground=C['brd'])
        self.blk_users.pack(fill=tk.X)
        br=self._btnrow(cd)
        Btn(br,"📂 Aus CSV laden",command=self._blk_csv,bg=C['input'],width=130).pack(side=tk.LEFT,padx=(0,8))
        Btn(br,"▶️ Ausführen",command=self._run_blk,bg=C['ok'],fg='#1a1b26',width=130).pack(side=tk.LEFT)
        self.blk_t=self._txtbox(cd,5)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  VERBINDUNG
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    def _connect(self):
        a=self.admin_e.get().strip()
        if not a: messagebox.showwarning("Fehlt","Admin-UPN!"); return

        # Warnung wenn Module fehlen
        missing_mods = [m['display'] for m in REQUIRED_MODULES
                        if not self.mod_status.get(m['name'], {}).get('installed')]
        if missing_mods:
            msg = "Folgende Module fehlen:\n• " + "\n• ".join(missing_mods) + "\n\nEinige Funktionen werden nicht verfügbar sein.\nTrotzdem verbinden?"
            if not messagebox.askyesno("⚠️ Module fehlen", msg, icon="warning"):
                return

        self.log("🔄 Verbinde...",C['warn']); self.conn_btn.configure(state=tk.DISABLED)
        def do():
            ok,_,e=self.ps.run(f'Connect-ExchangeOnline -UserPrincipalName "{a}" -ShowBanner:$false',180)
            if ok:
                v,vo,_=self.ps.run('Get-OrganizationConfig|Select -Expand Name',30)
                org=vo.strip().split("\n")[0] if v and vo.strip() else ""
                # Auch Graph verbinden wenn Modul vorhanden
                if self.mod_status.get('Microsoft.Graph', {}).get('installed'):
                    self.root.after(0, lambda: self.log("🔄 Verbinde Microsoft Graph...", C['warn']))
                    gok, _, ge = self.ps.run(f'Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All" -NoWelcome -EA SilentlyContinue', 120)
                    if gok:
                        self.root.after(0, lambda: self.log("  ✅ Graph verbunden", C['ok']))
                    else:
                        self.root.after(0, lambda: self.log(f"  ⚠️ Graph: {ge}", C['warn']))
                self.root.after(0,lambda:self._connected(org))
            else: self.root.after(0,lambda:[self.conn_btn.configure(state=tk.NORMAL),self.log(f"❌ {e}",C['err']),messagebox.showerror("Fehler",e)])
        threading.Thread(target=do,daemon=True).start()

    def _connected(self,org):
        self.connected=True; n=f" ({org})" if org else ""
        self.conn_lbl.configure(text=f"🟢 {org}" if org else "🟢 Verbunden",fg=C['ok'])
        self.conn_btn.configure(text="✅",bg=C['ok'])
        self.log(f"✅ Verbunden{n}!",C['ok']); self._load()

    def disconnect(self):
        self.log("🔌 Trenne...",C['warn']); self.ps.run("Disconnect-ExchangeOnline -Confirm:$false",30)
        self.ps.run("Disconnect-MgGraph -EA SilentlyContinue", 10)
        self.connected=False; self.mailboxes=[]; self.groups={'teams':[],'verteiler':[],'security':[]}
        self.conn_lbl.configure(text="⚫ Nicht verbunden",fg=C['dim'])
        self.conn_btn.configure(text="🔌 Verbinden",bg=C['accent']); self.conn_btn.configure(state=tk.NORMAL)
        for cb in [self.mb_t,self.mb_u,self.ob_u,self.ob_fwd,self.ui_u,self.sm_perm,self.fwd_src,self.fwd_dst,self.aud_u,self.blk_grp]:
            try: cb.configure(values=[],state="disabled"); cb.set("— Erst verbinden —")
            except: pass
        for k in ['teams','verteiler','security']:
            for a in ['gc','uc']:
                try: getattr(self,f'{k}_{a}').configure(values=[],state="disabled"); getattr(self,f'{k}_{a}').set("— Erst verbinden —")
                except: pass
        self.log("✅ Getrennt",C['ok'])

    def _load(self):
        self.log("📥 Lade Daten...",C['warn'])
        def do():
            r1,o1,_=self.ps.run('Get-Mailbox -ResultSize Unlimited|Select DisplayName,PrimarySmtpAddress,RecipientTypeDetails|ConvertTo-Json -Compress',180)
            r2,o2,_=self.ps.run('Get-UnifiedGroup -ResultSize Unlimited|Select DisplayName,PrimarySmtpAddress,GroupType|ConvertTo-Json -Compress',180)
            r3,o3,_=self.ps.run('Get-DistributionGroup -ResultSize Unlimited|Select DisplayName,PrimarySmtpAddress,GroupType|ConvertTo-Json -Compress',180)
            self.root.after(0,lambda:self._loaded(r1,o1,r2,o2,r3,o3))
        threading.Thread(target=do,daemon=True).start()

    def _loaded(self,r1,o1,r2,o2,r3,o3):
        if r1:
            self.mailboxes=[]
            for mb in self._pj(o1):
                e,n,t=mb.get('PrimarySmtpAddress',''),mb.get('DisplayName',''),mb.get('RecipientTypeDetails','')
                if e:
                    sh='Shared' in t
                    self.mailboxes.append({'d':f"{'👥' if sh else '👤'} {n} <{e}>",'e':e,'n':n,'t':'shared' if sh else 'user'})
            self.mailboxes.sort(key=lambda x:x['d']); self._amb(); self._upd_all()
            self.log(f"  📧 {len(self.mailboxes)} Postfächer",C['dim'])
        if r2:
            self.groups['teams']=[{'d':f"👥 {g.get('DisplayName','')} <{g.get('PrimarySmtpAddress','')}>",
                'e':g.get('PrimarySmtpAddress',''),'n':g.get('DisplayName','')} for g in self._pj(o2) if g.get('PrimarySmtpAddress')]
            self.groups['teams'].sort(key=lambda x:x['d']); self._ugrp('teams')
            self.log(f"  👥 {len(self.groups['teams'])} Teams",C['dim'])
        if r3:
            self.groups['verteiler']=[]; self.groups['security']=[]
            for g in self._pj(o3):
                e,n,gt=g.get('PrimarySmtpAddress',''),g.get('DisplayName',''),str(g.get('GroupType',''))
                if e:
                    ent={'d':f"{'🔒' if 'Security' in gt else '📨'} {n} <{e}>",'e':e,'n':n}
                    (self.groups['security'] if 'Security' in gt else self.groups['verteiler']).append(ent)
            for k in ['verteiler','security']: self.groups[k].sort(key=lambda x:x['d']); self._ugrp(k)
            self.log(f"  📨 {len(self.groups['verteiler'])} Verteiler, 🔒 {len(self.groups['security'])} Security",C['dim'])
        t=len(self.mailboxes)+sum(len(v) for v in self.groups.values())
        self.log(f"✅ {t} Objekte geladen!",C['ok'])

        # Lizenz-Warnung aktualisieren
        if not self.mod_status.get('Microsoft.Graph', {}).get('installed'):
            self._lic_warn.configure(text="⚠️ Microsoft.Graph-Modul fehlt — Lizenz-Abfragen nicht möglich")
        else:
            self._lic_warn.configure(text="")

    def _ugrp(self,k):
        i=[g['d'] for g in self.groups[k]]
        getattr(self,f'{k}_gc').configure(values=i,state="normal"); getattr(self,f'{k}_gc').set("")
        ui=[m['d'] for m in self.mailboxes]
        getattr(self,f'{k}_uc').configure(values=ui,state="normal"); getattr(self,f'{k}_uc').set("")
    def _upd_all(self):
        i=[m['d'] for m in self.mailboxes]
        for cb in [self.ob_u,self.ob_fwd,self.ui_u,self.sm_perm,self.fwd_src,self.fwd_dst,self.aud_u]:
            try: cb.configure(values=i,state="normal"); cb.set("")
            except: pass
        ag=[]
        for k in ['teams','verteiler','security']: ag+=[g['d'] for g in self.groups[k]]
        self.blk_grp.configure(values=ag,state="normal"); self.blk_grp.set("")

    # ── Filter ───────────────────────────────────────────
    def _amb(self):
        ft=self.mb_ft.get(); fl=self.mailboxes if ft=="all" else [m for m in self.mailboxes if m['t']==ft]
        i=[m['d'] for m in fl]; self.mb_t.configure(values=i,state="normal"); self.mb_u.configure(values=i,state="normal")
    def _fmb(self,e=None):
        s=self.mb_s.get().lower(); ft=self.mb_ft.get()
        b=self.mailboxes if ft=="all" else [m for m in self.mailboxes if m['t']==ft]
        fl=[m for m in b if s in m['d'].lower()] if s else b
        self.mb_t.configure(values=[m['d'] for m in fl]); self.mb_u.configure(values=[m['d'] for m in fl])
    def _fgrp(self,k):
        s=getattr(self,f'{k}_se').get().lower()
        fl=[g for g in self.groups[k] if s in g['d'].lower()] if s else self.groups[k]
        getattr(self,f'{k}_gc').configure(values=[g['d'] for g in fl])
    def _fob(self,e=None):
        s=self.ob_s.get().lower()
        fl=[m for m in self.mailboxes if s in m['d'].lower()] if s else self.mailboxes
        self.ob_u.configure(values=[m['d'] for m in fl])
    def _fui(self,e=None):
        s=self.ui_s.get().lower()
        fl=[m for m in self.mailboxes if s in m['d'].lower()] if s else self.mailboxes
        self.ui_u.configure(values=[m['d'] for m in fl])
    def _faud(self,e=None):
        s=self.aud_s.get().lower()
        fl=[m for m in self.mailboxes if s in m['d'].lower()] if s else self.mailboxes
        self.aud_u.configure(values=[m['d'] for m in fl])

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  LOGIK
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    def _add_mb(self):
        mb,us=self.mb_t.get().strip(),self.mb_u.get().strip()
        if not mb or mb.startswith("—") or not us or us.startswith("—"): messagebox.showwarning("Fehlt","Auswählen!"); return
        mbe,use=self._ge(mb),self._ge(us)
        if not messagebox.askyesno("Bestätigen",f"Hinzufügen?\n📬 {mbe}\n👤 {use}"): return
        def do():
            errs=[]
            if self.fa_v.get():
                am="$true" if self.am_v.get() else "$false"
                ok,_,e=self.ps.run(f'Add-MailboxPermission -Identity "{mbe}" -User "{use}" -AccessRights FullAccess -AutoMapping {am}')
                self.root.after(0,lambda:self.log(f"  {'✅ Vollzugriff' if ok else '❌ '+e}",C['ok'] if ok else C['err']))
                if not ok: errs.append(e)
            if self.sa_v.get():
                ok,_,e=self.ps.run(f'Add-RecipientPermission -Identity "{mbe}" -Trustee "{use}" -AccessRights SendAs -Confirm:$false')
                self.root.after(0,lambda:self.log(f"  {'✅ Senden als' if ok else '❌ '+e}",C['ok'] if ok else C['err']))
                if not ok: errs.append(e)
            self.root.after(0,lambda:self._done(errs,"hinzugefügt"))
        threading.Thread(target=do,daemon=True).start()

    def _rem_mb(self):
        mb,us=self.mb_t.get().strip(),self.mb_u.get().strip()
        if not mb or mb.startswith("—") or not us or us.startswith("—"): messagebox.showwarning("Fehlt","Auswählen!"); return
        mbe,use=self._ge(mb),self._ge(us)
        if not messagebox.askyesno("⚠️",f"Entfernen?\n📬 {mbe}\n👤 {use}",icon="warning"): return
        def do():
            errs=[]
            if self.fa_v.get():
                ok,_,e=self.ps.run(f'Remove-MailboxPermission -Identity "{mbe}" -User "{use}" -AccessRights FullAccess -Confirm:$false')
                if not ok: errs.append(e)
            if self.sa_v.get():
                ok,_,e=self.ps.run(f'Remove-RecipientPermission -Identity "{mbe}" -Trustee "{use}" -AccessRights SendAs -Confirm:$false')
                if not ok: errs.append(e)
            self.root.after(0,lambda:self._done(errs,"entfernt"))
        threading.Thread(target=do,daemon=True).start()

    def _agrp(self,k):
        g,u=getattr(self,f'{k}_gc').get().strip(),getattr(self,f'{k}_uc').get().strip()
        if not g or g.startswith("—") or not u or u.startswith("—"): messagebox.showwarning("Fehlt","Auswählen!"); return
        ge,ue=self._ge(g),self._ge(u); rv=getattr(self,f'{k}_rv').get()
        if not messagebox.askyesno("OK",f"Hinzufügen?\n📋 {ge}\n👤 {ue}"): return
        def do():
            if k=='teams': cmd=f'Add-UnifiedGroupLinks -Identity "{ge}" -LinkType {"Owners" if rv=="Owner" else "Members"} -Links "{ue}"'
            else: cmd=f'Add-DistributionGroupMember -Identity "{ge}" -Member "{ue}"'
            ok,_,e=self.ps.run(cmd); self.root.after(0,lambda:self._done([] if ok else [e],"hinzugefügt"))
        threading.Thread(target=do,daemon=True).start()

    def _rgrp(self,k):
        g,u=getattr(self,f'{k}_gc').get().strip(),getattr(self,f'{k}_uc').get().strip()
        if not g or g.startswith("—") or not u or u.startswith("—"): messagebox.showwarning("Fehlt","Auswählen!"); return
        ge,ue=self._ge(g),self._ge(u)
        if not messagebox.askyesno("⚠️",f"Entfernen?\n📋 {ge}\n👤 {ue}",icon="warning"): return
        def do():
            if k=='teams':
                ok,_,e=self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Members -Links "{ue}" -Confirm:$false')
                self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Owners -Links "{ue}" -Confirm:$false')
            else: ok,_,e=self.ps.run(f'Remove-DistributionGroupMember -Identity "{ge}" -Member "{ue}" -Confirm:$false')
            self.root.after(0,lambda:self._done([] if ok else [e],"entfernt"))
        threading.Thread(target=do,daemon=True).start()

    def _smem(self,k):
        g=getattr(self,f'{k}_gc').get().strip()
        if not g or g.startswith("—"): messagebox.showwarning("Fehlt","Gruppe!"); return
        ge=self._ge(g); self.log(f"  📋 Lade {ge}...",C['warn'])
        def do():
            if k=='teams':
                _,om,_=self.ps.run(f'Get-UnifiedGroupLinks -Identity "{ge}" -LinkType Members|Select Name,PrimarySmtpAddress|ConvertTo-Json -Compress',60)
                _,oo,_=self.ps.run(f'Get-UnifiedGroupLinks -Identity "{ge}" -LinkType Owners|Select Name,PrimarySmtpAddress|ConvertTo-Json -Compress',60)
                ms,ow=self._pj(om),self._pj(oo)
                ln=[f"=== {ge} ===",""]
                if ow: ln+=[f"👑 Besitzer ({len(ow)}):"] + [f"  • {o.get('Name','')} <{o.get('PrimarySmtpAddress','')}>" for o in ow]+[""]
                ln+=[f"👤 Mitglieder ({len(ms)}):"] + [f"  • {m.get('Name','')} <{m.get('PrimarySmtpAddress','')}>" for m in ms]
            else:
                _,o,_=self.ps.run(f'Get-DistributionGroupMember -Identity "{ge}" -ResultSize Unlimited|Select Name,PrimarySmtpAddress|ConvertTo-Json -Compress',60)
                ms=self._pj(o); ln=[f"=== {ge} ===","",f"👤 Mitglieder ({len(ms)}):"] + [f"  • {m.get('Name','')} <{m.get('PrimarySmtpAddress','')}>" for m in ms]
            self.root.after(0,lambda:self._settxt(getattr(self,f'{k}_mt'),"\n".join(ln)))
        threading.Thread(target=do,daemon=True).start()

    # ── Offboarding ──────────────────────────────────────
    def _run_ob(self):
        us=self.ob_u.get().strip()
        if not us or us.startswith("—"): messagebox.showwarning("Fehlt","Benutzer!"); return
        active={k:v.get() for k,v in self.ob_v.items()}
        if not any(active.values()): messagebox.showwarning("Fehlt","Mindestens 1 Schritt!"); return
        ue=self._ge(us); un=us.split('<')[0].strip().lstrip('👤👥 ')
        names=[OB_N[k] for k,v in active.items() if v]

        # Warnung bei fehlendem Graph-Modul
        graph_needed = active.get('reset_pw') or active.get('remove_licenses') or active.get('sign_in')
        graph_ok = self.mod_status.get('Microsoft.Graph', {}).get('installed', False)
        graph_warn = ""
        if graph_needed and not graph_ok:
            graph_warn = "\n\n⚠️ Microsoft.Graph fehlt! PW-Reset, Anmelde-Block und Lizenzen werden über Graph gesteuert und könnten fehlschlagen."

        msg=f"⚠️ OFFBOARDING:\n👤 {un}\n📧 {ue}\n\n"+"\n".join(f"  • {n}" for n in names)+f"\n\n⚠️ Irreversibel!{graph_warn}"
        if not messagebox.askyesno("⚠️",msg,icon="warning"): return
        if not messagebox.askyesno("🔴","WIRKLICH?",icon="warning"): return
        self.ob_rb.configure(state=tk.DISABLED);self.ob_eb.configure(state=tk.DISABLED)
        self._settxt(self.ob_rt,"")
        self.ob_report=["="*55,"OFFBOARDING-BERICHT",f"Datum: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
                         f"Benutzer: {un}",f"E-Mail: {ue}",f"Admin: {self.admin_e.get().strip()}","="*55,""]
        def do():
            todo=[(k,v) for k,v in active.items() if v]; total=len(todo); res={}
            for i,(sk,_) in enumerate(todo):
                pct=int(i/total*100); sn=OB_N[sk]
                self.root.after(0,lambda p=pct,s=sn:[self.ob_pb.configure(value=p),self.ob_pl.configure(text=f"⏳ {s}... ({p}%)")])
                self.root.after(0,lambda s=sn:self.log(f"  🔄 {s}...",C['warn']))
                ok,det=self._ob_step(sk,ue); res[sk]=(ok,det)
                st="✅" if ok else "❌"; self.ob_report+=[f"[{st}] {sn}",f"    {det}",""]
                self.root.after(0,lambda s=sn,st2=st,c2=(C['ok'] if ok else C['err']):self.log(f"  {st2} {s}",c2))
            sc=sum(1 for ok,_ in res.values() if ok); fc=sum(1 for ok,_ in res.values() if not ok)
            self.ob_report+=["="*55,f"ERGEBNIS: {sc} OK, {fc} FEHLER","="*55]
            self.root.after(0,lambda:[self.ob_pb.configure(value=100),self.ob_pl.configure(text="✅ Fertig"),
                self.ob_rb.configure(state=tk.NORMAL),self.ob_eb.configure(state=tk.NORMAL),
                self._settxt(self.ob_rt,"\n".join(self.ob_report)),
                self.log(f"🚪 {ue}: {sc}✅/{fc}❌",C['ok'] if fc==0 else C['warn'])])
        threading.Thread(target=do,daemon=True).start()

    def _ob_step(self,step,ue):
        try:
            if step=='sign_in':
                # Graph-basiert: Account deaktivieren
                if self.mod_status.get('Microsoft.Graph', {}).get('installed'):
                    ok,o,e=self.ps.run(
                        f'Update-MgUser -UserId "{ue}" -AccountEnabled:$false -EA Stop; Write-Output "BLOCK_OK"', 60)
                    if "BLOCK_OK" in o:
                        # Bestehende Sessions widerrufen
                        self.ps.run(f'Revoke-MgUserSignInSession -UserId "{ue}" -EA SilentlyContinue', 30)
                        return True, "Blockiert + Sessions widerrufen"
                    return False, f"Fehler: {e}"
                else:
                    # Fallback EXO
                    ok,_,e=self.ps.run(f'Set-User -Identity "{ue}" -AccountDisabled $true',60)
                    return ok,"Blockiert (EXO)" if ok else f"Fehler: {e}"

            elif step=='reset_pw':
                # NEU: Passwort-Reset über Microsoft Graph statt Set-Mailbox
                if self.mod_status.get('Microsoft.Graph', {}).get('installed'):
                    ok,o,e=self.ps.run(
                        f'$chars="abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%&*";'
                        f'$pw=-join(1..24|ForEach-Object{{$chars[(Get-Random -Max $chars.Length)]}});'
                        f'$params=@{{PasswordProfile=@{{Password=$pw;ForceChangePasswordNextSignIn=$true}}}};'
                        f'Update-MgUser -UserId "{ue}" @params -EA Stop;'
                        f'Write-Output "PWOK"', 60)
                    if "PWOK" in o:
                        return True, "PW zurückgesetzt (Graph, Änderung beim nächsten Login erzwungen)"
                    return False, f"Graph-Fehler: {e}"
                else:
                    return False, "Microsoft.Graph-Modul fehlt — PW-Reset nicht möglich. Bitte manuell im Admin Center."

            elif step=='remove_groups':
                rm,fl=0,0
                ok,o,_=self.ps.run(f'Get-UnifiedGroup -ResultSize Unlimited|Where-Object{{(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})-or(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Owners -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand PrimarySmtpAddress',180)
                if ok and o.strip():
                    for g in o.strip().split("\n"):
                        g=g.strip()
                        if not g: continue
                        r,_,_=self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{g}" -LinkType Members -Links "{ue}" -Confirm:$false -EA SilentlyContinue')
                        self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{g}" -LinkType Owners -Links "{ue}" -Confirm:$false -EA SilentlyContinue')
                        if r: rm+=1
                        else: fl+=1
                ok,o,_=self.ps.run(f'Get-DistributionGroup -ResultSize Unlimited|Where-Object{{(Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand PrimarySmtpAddress',180)
                if ok and o.strip():
                    for g in o.strip().split("\n"):
                        g=g.strip()
                        if not g: continue
                        r,_,_=self.ps.run(f'Remove-DistributionGroupMember -Identity "{g}" -Member "{ue}" -Confirm:$false -EA SilentlyContinue')
                        if r: rm+=1
                        else: fl+=1
                return fl==0,f"{rm} Gruppen entfernt"+("" if fl==0 else f", {fl} Fehler")

            elif step=='remove_licenses':
                if not self.mod_status.get('Microsoft.Graph', {}).get('installed'):
                    return False, "Microsoft.Graph-Modul fehlt — Lizenzen manuell entziehen"
                ok,o,e=self.ps.run(f'$s=(Get-MgUserLicenseDetail -UserId "{ue}" -EA SilentlyContinue).SkuId;if($s){{foreach($k in $s){{Set-MgUserLicense -UserId "{ue}" -RemoveLicenses @($k) -AddLicenses @() -EA Stop}};Write-Output "LR:$($s.Count)"}}else{{Write-Output "NG"}}',90)
                if "LR:" in o: return True,f"{o.split('LR:')[1].strip().split(chr(10))[0]} Lizenz(en) entfernt"
                if "NG" in o: return True, "Keine Lizenzen zugewiesen"
                return False,f"Fehler: {e}"

            elif step=='convert_shared':
                ok,_,e=self.ps.run(f'Set-Mailbox -Identity "{ue}" -Type Shared',60)
                return ok,"→ Shared" if ok else f"Fehler: {e}"
            elif step=='set_ooo':
                msg=self.ob_ooo.get('1.0',tk.END).strip()
                if not msg: return True,"Übersprungen"
                esc=msg.replace("'","''").replace('"','`"')
                ok,_,e=self.ps.run(f'Set-MailboxAutoReplyConfiguration -Identity "{ue}" -AutoReplyState Enabled -InternalMessage "{esc}" -ExternalMessage "{esc}" -ExternalAudience All',60)
                return ok,"OOO an" if ok else f"Fehler: {e}"
            elif step=='fwd':
                fw=self.ob_fwd.get().strip()
                if not fw or fw.startswith("—"): return True,"Übersprungen"
                fe=self._ge(fw) if '<' in fw else fw
                ok,_,e=self.ps.run(f'Set-Mailbox -Identity "{ue}" -ForwardingSmtpAddress "smtp:{fe}" -DeliverToMailboxAndForward $true',60)
                return ok,f"→ {fe}" if ok else f"Fehler: {e}"
            elif step=='hide_gal':
                ok,_,e=self.ps.run(f'Set-Mailbox -Identity "{ue}" -HiddenFromAddressListsEnabled $true',60)
                return ok,"GAL versteckt" if ok else f"Fehler: {e}"
            elif step=='disable_sync':
                ok,_,e=self.ps.run(f'Set-CASMailbox -Identity "{ue}" -ActiveSyncEnabled $false -OWAEnabled $false -PopEnabled $false -ImapEnabled $false -MAPIEnabled $false -EwsEnabled $false',60)
                return ok,"Protokolle aus" if ok else f"Fehler: {e}"
            elif step=='remove_delegates':
                ok,o,e=self.ps.run(f'$p=Get-MailboxPermission -Identity "{ue}"|Where-Object{{$_.User -ne "NT AUTHORITY\\SELF" -and $_.IsInherited -eq $false}};$c=0;foreach($x in $p){{Remove-MailboxPermission -Identity "{ue}" -User $x.User -AccessRights $x.AccessRights -Confirm:$false -EA SilentlyContinue;$c++}};Write-Output "DD:$c"',90)
                if "DD:" in o: return True,f"{o.split('DD:')[1].strip().split(chr(10))[0]} entfernt"
                return False,f"Fehler: {e}"
            return False,"?"
        except Exception as ex: return False,str(ex)

    def _exp_ob(self):
        if not self.ob_report: return
        us=self.ob_u.get().strip(); ue=self._ge(us) if '<' in us else "user"
        un=ue.split('@')[0] if '@' in ue else ue
        fp=filedialog.asksaveasfilename(defaultextension=".txt",initialfile=f"Offboarding_{un}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",filetypes=[("Text","*.txt")])
        if fp:
            with open(fp,'w',encoding='utf-8') as f: f.write("\n".join(self.ob_report))
            self.log(f"💾 {fp}",C['ok'])

    # ── Benutzer-Info ────────────────────────────────────
    def _load_ui(self):
        us=self.ui_u.get().strip()
        if not us or us.startswith("—"): messagebox.showwarning("Fehlt","Benutzer!"); return
        ue=self._ge(us); self.log(f"  🔍 Info {ue}...",C['warn'])
        def do():
            ln=[f"{'='*50}",f"  {ue}",f"{'='*50}",""]
            ok,o,_=self.ps.run(f'Get-Mailbox -Identity "{ue}"|Select DisplayName,PrimarySmtpAddress,RecipientTypeDetails,ForwardingSmtpAddress,HiddenFromAddressListsEnabled,WhenCreated|ConvertTo-Json',60)
            if ok:
                d=self._pj(o); mb=d[0] if d else {}
                ln+=[f"📧 Name: {mb.get('DisplayName','')}",f"📧 Typ: {mb.get('RecipientTypeDetails','')}",
                     f"📨 Weiterleitung: {mb.get('ForwardingSmtpAddress','Keine')}",
                     f"👻 GAL: {'Versteckt' if mb.get('HiddenFromAddressListsEnabled') else 'Sichtbar'}",
                     f"📅 Erstellt: {mb.get('WhenCreated','')}",""]
            ok,o,_=self.ps.run(f'Get-MailboxStatistics -Identity "{ue}" -EA SilentlyContinue|Select TotalItemSize,ItemCount|ConvertTo-Json',30)
            if ok:
                d=self._pj(o); st=d[0] if d else {}
                ln+=[f"📦 Größe: {st.get('TotalItemSize','')}",f"📬 Elemente: {st.get('ItemCount','')}",""]

            # Graph-basierte Lizenz-Info
            if self.mod_status.get('Microsoft.Graph', {}).get('installed'):
                ok,o,_=self.ps.run(f'$lics=Get-MgUserLicenseDetail -UserId "{ue}" -EA SilentlyContinue;if($lics){{$lics|Select -Expand SkuPartNumber|ForEach-Object{{Write-Output "LIC:$_"}}}}else{{Write-Output "LIC:Keine"}}',30)
                if ok and o.strip():
                    lics=[l.split("LIC:")[1] for l in o.strip().split("\n") if "LIC:" in l]
                    ln+=[f"📊 Lizenzen ({len(lics)}):"] + [f"  • {l}" for l in lics] + [""]

            ok,o,_=self.ps.run(f'Get-UnifiedGroup -ResultSize Unlimited|Where-Object{{(Get-UnifiedGroupLinks -Identity $_.Identity -LinkType Members -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand DisplayName',120)
            if ok and o.strip():
                gs=[g.strip() for g in o.strip().split("\n") if g.strip()]
                ln+=[f"👥 Teams ({len(gs)}):"] + [f"  • {g}" for g in gs]+[""]
            ok,o,_=self.ps.run(f'Get-DistributionGroup -ResultSize Unlimited|Where-Object{{(Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited -EA SilentlyContinue|Where-Object{{$_.PrimarySmtpAddress -eq "{ue}"}})}}|Select -Expand DisplayName',120)
            if ok and o.strip():
                gs=[g.strip() for g in o.strip().split("\n") if g.strip()]
                ln+=[f"📨 Verteiler/Security ({len(gs)}):"] + [f"  • {g}" for g in gs]+[""]
            ok,o,_=self.ps.run(f'Get-Mailbox -ResultSize Unlimited|Get-MailboxPermission|Where-Object{{$_.User -like "*{ue}*" -and $_.AccessRights -like "*FullAccess*"}}|Select -Expand Identity',60)
            if ok and o.strip():
                ps2=[p.strip() for p in o.strip().split("\n") if p.strip()]
                ln+=[f"🔑 Vollzugriff auf ({len(ps2)}):"] + [f"  • {p}" for p in ps2]
            self.root.after(0,lambda:[self._settxt(self.ui_t,"\n".join(ln)),self.log("  ✅ Info geladen",C['ok'])])
        threading.Thread(target=do,daemon=True).start()

    # ── Lizenzen ─────────────────────────────────────────
    def _load_lic(self):
        if not self.mod_status.get('Microsoft.Graph', {}).get('installed'):
            messagebox.showerror("Modul fehlt",
                "Microsoft.Graph ist nicht installiert.\n\n"
                "Bitte zurück zur Vorab-Prüfung und installieren,\n"
                "oder manuell: Install-Module Microsoft.Graph -Scope CurrentUser")
            return
        self.log("  📊 Lizenzen...",C['warn'])
        def do():
            ok,o,e=self.ps.run('Get-MgSubscribedSku -EA SilentlyContinue|Select SkuPartNumber,ConsumedUnits,@{N="Total";E={$_.PrepaidUnits.Enabled}}|ConvertTo-Json',60)
            if ok and o.strip():
                self._lic_data=self._pj(o)
                ln=[f"{'Lizenz':<40} {'Benutzt':>8} {'Gesamt':>8} {'Frei':>8}","─"*68]
                for d in self._lic_data:
                    n,u,t=d.get('SkuPartNumber',''),d.get('ConsumedUnits',0),d.get('Total',0)
                    try: u,t=int(u),int(t)
                    except: u,t=0,0
                    ln.append(f"{n:<40} {u:>8} {t:>8} {t-u:>8}")
                self.root.after(0,lambda:[self._settxt(self.lic_t,"\n".join(ln)),self.log("  ✅ Lizenzen",C['ok'])])
            else:
                self.root.after(0,lambda:[
                    self._settxt(self.lic_t,f"❌ Fehler beim Laden der Lizenzen\n\n{e}\n\nGraph-Verbindung aktiv?"),
                    self.log(f"  ❌ Lizenzen: {e}", C['err'])
                ])
        threading.Thread(target=do,daemon=True).start()

    def _exp_lic(self):
        if not self._lic_data: messagebox.showwarning("Fehlt","Erst laden!"); return
        fp=filedialog.asksaveasfilename(defaultextension=".csv",initialfile=f"Lizenzen_{datetime.now().strftime('%Y%m%d')}.csv",filetypes=[("CSV","*.csv")])
        if fp:
            with open(fp,'w',newline='',encoding='utf-8') as f:
                w=csv.writer(f,delimiter=';'); w.writerow(['Lizenz','Benutzt','Gesamt','Frei'])
                for d in self._lic_data:
                    n,u,t=d.get('SkuPartNumber',''),d.get('ConsumedUnits',0),d.get('Total',0)
                    try: u,t=int(u),int(t)
                    except: u,t=0,0
                    w.writerow([n,u,t,t-u])
            self.log(f"💾 {fp}",C['ok'])

    # ── Shared Mailbox ───────────────────────────────────
    def _create_sm(self):
        nm,em=self.sm_name.get().strip(),self.sm_email.get().strip()
        if not nm or not em: messagebox.showwarning("Fehlt","Name + E-Mail!"); return
        dp=self.sm_disp.get().strip() or nm
        if not messagebox.askyesno("Erstellen",f"📧 {em}\n👤 {dp}"): return
        def do():
            self.root.after(0,lambda:self.log(f"  📧 Erstelle {em}...",C['warn']))
            ok,_,e=self.ps.run(f'New-Mailbox -Name "{nm}" -PrimarySmtpAddress "{em}" -DisplayName "{dp}" -Shared',60)
            if ok:
                self.root.after(0,lambda:self.log(f"  ✅ {em}",C['ok']))
                pu=self.sm_perm.get().strip()
                if pu and not pu.startswith("—"):
                    pue=self._ge(pu)
                    self.ps.run(f'Add-MailboxPermission -Identity "{em}" -User "{pue}" -AccessRights FullAccess -AutoMapping $true')
                    self.ps.run(f'Add-RecipientPermission -Identity "{em}" -Trustee "{pue}" -AccessRights SendAs -Confirm:$false')
                    self.root.after(0,lambda:self.log(f"  ✅ Rechte für {pue}",C['ok']))
                self.root.after(0,lambda:messagebox.showinfo("OK",f"✅ {em} erstellt!"))
            else: self.root.after(0,lambda:[self.log(f"  ❌ {e}",C['err']),messagebox.showerror("Fehler",e)])
        threading.Thread(target=do,daemon=True).start()

    # ── Weiterleitungen ──────────────────────────────────
    def _load_fwd(self):
        self.log("  📬 Weiterleitungen...",C['warn'])
        def do():
            ok,o,_=self.ps.run('Get-Mailbox -ResultSize Unlimited|Where-Object{$_.ForwardingSmtpAddress -ne $null}|Select PrimarySmtpAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward|ConvertTo-Json -Compress',120)
            data=self._pj(o) if ok else []
            if data:
                ln=[f"{'Postfach':<35} {'→ Weiterleitung':<35} {'Kopie':>5}","─"*78]
                for d in data: ln.append(f"{d.get('PrimarySmtpAddress',''):<35} {str(d.get('ForwardingSmtpAddress','')):<35} {'Ja' if d.get('DeliverToMailboxAndForward') else 'Nein':>5}")
                self.root.after(0,lambda:[self._settxt(self.fwd_t,"\n".join(ln)),self.log(f"  ✅ {len(data)} Weiterleitungen",C['ok'])])
            else: self.root.after(0,lambda:self._settxt(self.fwd_t,"Keine Weiterleitungen."))
        threading.Thread(target=do,daemon=True).start()
    def _set_fwd(self):
        src,dst=self.fwd_src.get().strip(),self.fwd_dst.get().strip()
        if not src or src.startswith("—") or not dst or dst.startswith("—"): messagebox.showwarning("Fehlt","Auswählen!"); return
        se,de=self._ge(src),self._ge(dst); keep="$true" if self.fwd_keep.get() else "$false"
        if not messagebox.askyesno("Setzen",f"📬 {se} → {de}"): return
        def do():
            ok,_,e=self.ps.run(f'Set-Mailbox -Identity "{se}" -ForwardingSmtpAddress "smtp:{de}" -DeliverToMailboxAndForward {keep}')
            self.root.after(0,lambda:self._done([] if ok else [e],"gesetzt"))
        threading.Thread(target=do,daemon=True).start()
    def _rem_fwd(self):
        src=self.fwd_src.get().strip()
        if not src or src.startswith("—"): messagebox.showwarning("Fehlt","Postfach!"); return
        se=self._ge(src)
        if not messagebox.askyesno("Entfernen",f"Weiterleitung für {se} entfernen?"): return
        def do():
            ok,_,e=self.ps.run(f'Set-Mailbox -Identity "{se}" -ForwardingSmtpAddress $null')
            self.root.after(0,lambda:self._done([] if ok else [e],"entfernt"))
        threading.Thread(target=do,daemon=True).start()

    # ── Audit ────────────────────────────────────────────
    def _run_aud(self):
        us=self.aud_u.get().strip()
        if not us or us.startswith("—"): messagebox.showwarning("Fehlt","Postfach!"); return
        ue=self._ge(us); self.log(f"  🔍 Audit {ue}...",C['warn']); self._aud_data=[]
        def do():
            ln=[f"{'='*50}",f"  AUDIT: {ue}",f"{'='*50}",""]
            ok,o,_=self.ps.run(f'Get-MailboxPermission -Identity "{ue}"|Where-Object{{$_.User -ne "NT AUTHORITY\\SELF" -and $_.IsInherited -eq $false}}|Select User,AccessRights|ConvertTo-Json -Compress',60)
            perms=self._pj(o) if ok else []
            if perms:
                ln+=["📂 Vollzugriff:"]
                for p in perms: ln.append(f"  • {p.get('User','')}"); self._aud_data.append({'Postfach':ue,'Typ':'FullAccess','Benutzer':str(p.get('User',''))})
                ln.append("")
            ok,o,_=self.ps.run(f'Get-RecipientPermission -Identity "{ue}"|Where-Object{{$_.Trustee -ne "NT AUTHORITY\\SELF"}}|Select Trustee|ConvertTo-Json -Compress',60)
            perms=self._pj(o) if ok else []
            if perms:
                ln+=["✉️ Senden als:"]
                for p in perms: ln.append(f"  • {p.get('Trustee','')}"); self._aud_data.append({'Postfach':ue,'Typ':'SendAs','Benutzer':str(p.get('Trustee',''))})
                ln.append("")
            ok,o,_=self.ps.run(f'Get-Mailbox -Identity "{ue}"|Select -Expand GrantSendOnBehalfTo',30)
            if ok and o.strip():
                sob=[x.strip() for x in o.strip().split("\n") if x.strip()]
                ln+=["📤 Senden im Auftrag:"] + [f"  • {s}" for s in sob]
                for s in sob: self._aud_data.append({'Postfach':ue,'Typ':'SendOnBehalf','Benutzer':s})
            if not self._aud_data: ln.append("✅ Keine Berechtigungen.")
            self.root.after(0,lambda:[self._settxt(self.aud_t,"\n".join(ln)),self.log(f"  ✅ {len(self._aud_data)} Einträge",C['ok'])])
        threading.Thread(target=do,daemon=True).start()
    def _exp_aud(self):
        if not self._aud_data: messagebox.showwarning("Fehlt","Erst Audit!"); return
        fp=filedialog.asksaveasfilename(defaultextension=".csv",initialfile=f"Audit_{datetime.now().strftime('%Y%m%d')}.csv",filetypes=[("CSV","*.csv")])
        if fp:
            with open(fp,'w',newline='',encoding='utf-8') as f:
                w=csv.DictWriter(f,fieldnames=['Postfach','Typ','Benutzer'],delimiter=';'); w.writeheader(); w.writerows(self._aud_data)
            self.log(f"💾 {fp}",C['ok'])

    # ── CSV-Export ───────────────────────────────────────
    def _run_csv(self):
        active={k:v.get() for k,v in self.csv_opts.items()}
        if not any(active.values()): messagebox.showwarning("Fehlt","Option!"); return
        fp=filedialog.askdirectory(title="Export-Ordner")
        if not fp: return
        self.log("📋 CSV-Export...",C['warn'])
        def do():
            n=0; ts=datetime.now().strftime('%Y%m%d')
            if active.get('users'):
                p=os.path.join(fp,f"Benutzer_{ts}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w=csv.writer(f,delimiter=';'); w.writerow(['Name','E-Mail','Typ'])
                    for m in self.mailboxes: w.writerow([m['n'],m['e'],m['t']])
                n+=1
            if active.get('shared'):
                p=os.path.join(fp,f"Shared_{ts}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w=csv.writer(f,delimiter=';'); w.writerow(['Name','E-Mail'])
                    for m in self.mailboxes:
                        if m['t']=='shared': w.writerow([m['n'],m['e']])
                n+=1
            if active.get('groups'):
                p=os.path.join(fp,f"Gruppen_{ts}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w=csv.writer(f,delimiter=';'); w.writerow(['Typ','Name','E-Mail'])
                    for k in ['teams','verteiler','security']:
                        for g in self.groups[k]: w.writerow([k,g['n'],g['e']])
                n+=1
            if active.get('forwarding'):
                ok,o,_=self.ps.run('Get-Mailbox -ResultSize Unlimited|Where-Object{$_.ForwardingSmtpAddress -ne $null}|Select PrimarySmtpAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward|ConvertTo-Json -Compress',120)
                data=self._pj(o) if ok else []
                p=os.path.join(fp,f"Weiterleitungen_{ts}.csv")
                with open(p,'w',newline='',encoding='utf-8') as f:
                    w=csv.writer(f,delimiter=';'); w.writerow(['Postfach','Weiterleitung','Kopie'])
                    for d in data: w.writerow([d.get('PrimarySmtpAddress',''),d.get('ForwardingSmtpAddress',''),d.get('DeliverToMailboxAndForward','')])
                n+=1
            self.root.after(0,lambda:[self.log(f"✅ {n} CSV(s) → {fp}",C['ok']),messagebox.showinfo("Export",f"{n} Datei(en) in {fp}")])
        threading.Thread(target=do,daemon=True).start()

    # ── Bulk ─────────────────────────────────────────────
    def _blk_csv(self):
        fp=filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("Text","*.txt")])
        if fp:
            with open(fp,'r',encoding='utf-8') as f: lines=[l.strip() for l in f if l.strip() and '@' in l]
            self.blk_users.delete('1.0',tk.END); self.blk_users.insert('1.0',"\n".join(lines))
            self.log(f"  📂 {len(lines)} geladen",C['ok'])
    def _run_blk(self):
        grp=self.blk_grp.get().strip()
        if not grp or grp.startswith("—"): messagebox.showwarning("Fehlt","Gruppe!"); return
        users=[l.strip() for l in self.blk_users.get('1.0',tk.END).strip().split("\n") if l.strip() and '@' in l.strip()]
        if not users: messagebox.showwarning("Fehlt","Benutzer!"); return
        ge=self._ge(grp); adding="Hinzufügen" in self.blk_act.get()
        if not messagebox.askyesno("Bulk",f"{'Hinzufügen' if adding else 'Entfernen'}: {len(users)} → {ge}"): return
        self.log(f"  🏷️ Bulk: {len(users)} → {ge}",C['warn'])
        def do():
            ok_c,err_c=0,0
            is_uni=any(g['e']==ge for g in self.groups.get('teams',[]))
            for u in users:
                u=u.strip()
                if not u: continue
                if adding:
                    if is_uni: r,_,_=self.ps.run(f'Add-UnifiedGroupLinks -Identity "{ge}" -LinkType Members -Links "{u}" -EA SilentlyContinue')
                    else: r,_,_=self.ps.run(f'Add-DistributionGroupMember -Identity "{ge}" -Member "{u}" -EA SilentlyContinue')
                else:
                    if is_uni:
                        r,_,_=self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Members -Links "{u}" -Confirm:$false -EA SilentlyContinue')
                        self.ps.run(f'Remove-UnifiedGroupLinks -Identity "{ge}" -LinkType Owners -Links "{u}" -Confirm:$false -EA SilentlyContinue')
                    else: r,_,_=self.ps.run(f'Remove-DistributionGroupMember -Identity "{ge}" -Member "{u}" -Confirm:$false -EA SilentlyContinue')
                if r: ok_c+=1
                else: err_c+=1
            self.root.after(0,lambda:[self._settxt(self.blk_t,f"✅ {ok_c} OK\n❌ {err_c} Fehler" if err_c else f"✅ {ok_c} OK"),
                self.log(f"  🏷️ {ok_c}✅ {err_c}❌",C['ok'] if err_c==0 else C['warn'])])
        threading.Thread(target=do,daemon=True).start()

    # ── Allgemein ────────────────────────────────────────
    def _done(self,errs,word):
        if errs:
            for e in errs: self.log(f"  ❌ {e}",C['err'])
            messagebox.showerror("Fehler","Siehe Protokoll.")
        else: messagebox.showinfo("OK",f"✅ {word}!")

    def log(self,msg,color=None):
        ts=datetime.now().strftime("%H:%M:%S")
        tag={C['ok']:'ok',C['err']:'err',C['warn']:'warn',C['accent']:'acc',C['dim']:'dim'}.get(color,'info')
        self.log_t.configure(state=tk.NORMAL)
        self.log_t.insert(tk.END,f"[{ts}] ",'info'); self.log_t.insert(tk.END,f"{msg}\n",tag)
        self.log_t.see(tk.END); self.log_t.configure(state=tk.DISABLED)

    def cleanup(self):
        try: self.ps.run("Disconnect-ExchangeOnline -Confirm:$false",10)
        except: pass
        try: self.ps.run("Disconnect-MgGraph -EA SilentlyContinue",5)
        except: pass
        self.ps.stop()

def main():
    root=tk.Tk()
    try: root.iconbitmap('exchange.ico')
    except: pass
    app=App(root)
    root.protocol("WM_DELETE_WINDOW",lambda:[app.cleanup(),root.destroy()] if messagebox.askokcancel("Beenden","Beenden?") else None)
    root.mainloop()

if __name__=="__main__": main()