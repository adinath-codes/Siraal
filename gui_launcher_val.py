"""
gui_launcher_val.py  —  Siraal Dynamic Validation Editor  (Apple Edition v2)
============================================================================
GUI for managing factory constraints (custom_rules.json).
v2: larger fonts throughout, calmer ambient glow, stronger borders.
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import simpledialog, messagebox
import json, os, math, datetime

try:
    from audit_logger import log_event
except ImportError:
    def log_event(role, action, details, is_warning=False): pass

ctk.set_appearance_mode("Dark")

C = {
    "bg":          "#090C14",
    "surface":     "#0E1220",
    "glass":       "#121828",
    "glass2":      "#16202E",
    "glass3":      "#1C2840",
    "card":        "#131C2C",
    "border":      "#263552",
    "border2":     "#2E4068",
    "border_hi":   "#3A5080",
    "violet":      "#9B7EF5",
    "violet_dim":  "#2A1C56",
    "violet_glow": "#1A1238",
    "teal":        "#22C9B5",
    "teal_dim":    "#0B2E2A",
    "teal_glow":   "#081E1A",
    "gold":        "#E8940A",
    "gold_dim":    "#382400",
    "gold_glow":   "#261800",
    "amber":       "#E86A10",
    "amber_dim":   "#361600",
    "blue":        "#3B82F6",
    "blue_dim":    "#182C50",
    "blue_glow":   "#0E1C38",
    "ok":          "#2EC98A",
    "ok_dim":      "#0A2818",
    "warn":        "#F0B429",
    "warn_dim":    "#342200",
    "error":       "#F06060",
    "error_dim":   "#381010",
    "info":        "#5B9CF5",
    "info_dim":    "#182846",
    "text":        "#EAF0FF",
    "text2":       "#9AAACE",
    "text3":       "#5A6E90",
}

RULES_FILE     = "custom_rules.json"
ADMIN_PASSWORD = "admin"
CORNER         = 12

SEV_COLOR = {"ERROR": C["error"], "WARNING": C["warn"], "INFO": C["info"]}
SEV_DIM   = {"ERROR": C["error_dim"], "WARNING": C["warn_dim"], "INFO": C["info_dim"]}

def get_all_target_types():
    base = ["ALL","Spur_Gear_3D","Helical_Gear","Ring_Gear_3D","Bevel_Gear",
            "Worm","Worm_Wheel","Box","Cylinder","Flange","Stepped_Shaft","L_Bracket"]
    customs = []
    t_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    if os.path.exists(t_path):
        for f in os.listdir(t_path):
            if f.startswith("Custom_") and f.endswith(".json"):
                customs.append(f.replace(".json",""))
    return base + sorted(customs)

GEAR_TYPES = get_all_target_types()
MATERIALS  = ["ALL","Steel-1020","Steel-4140","Al-6061","Brass-C360","Nylon-66","Ti-6Al-4V"]
SEVERITIES = ["ERROR","WARNING","INFO"]


# ── Ambient (3 subtle orbs, calmer) ──────────────────────────────────────────
def lerp_color(c1, c2, t):
    r1,g1,b1 = int(c1[1:3],16),int(c1[3:5],16),int(c1[5:7],16)
    r2,g2,b2 = int(c2[1:3],16),int(c2[3:5],16),int(c2[5:7],16)
    return "#{:02X}{:02X}{:02X}".format(
        int(r1+(r2-r1)*t), int(g1+(g2-g1)*t), int(b1+(b2-b1)*t))

class AmbientCanvas(tk.Canvas):
    ORBS  = [(0.10,0.20,180,"#0A1E3A",0.00018,0.00014),
             (0.85,0.55,160,"#160C38",-0.00015,0.00018),
             (0.50,0.85,140,"#0A2020",0.00020,-0.00016)]
    STEPS = 14

    def __init__(self, parent, **kw):
        super().__init__(parent, highlightthickness=0, bd=0, **kw)
        self._t=0.0; self._aid=None
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def _draw(self):
        if self._aid: self.after_cancel(self._aid)
        w=self.winfo_width() or 1160; h=self.winfo_height() or 760
        self.delete("all"); self.configure(bg=C["bg"]); self._t+=1
        for rx,ry,radius,color,sx,sy in self.ORBS:
            cx=(rx+math.sin(self._t*sx*60)*0.08)*w
            cy=(ry+math.cos(self._t*sy*60)*0.08)*h
            for i in range(self.STEPS,0,-1):
                t=i/self.STEPS; r=int(radius*t)
                col=lerp_color(color,C["bg"],t**0.45)
                self.create_oval(cx-r,cy-r,cx+r,cy+r,fill=col,outline="")
        for gx in range(0,w,64):
            self.create_line(gx,0,gx,h,fill=lerp_color("#18253A",C["bg"],0.82),width=1)
        for gy in range(0,h,64):
            self.create_line(0,gy,w,gy,fill=lerp_color("#18253A",C["bg"],0.82),width=1)
        self._aid=self.after(60,self._draw)

    def stop(self):
        if self._aid: self.after_cancel(self._aid)


class GlowSep(tk.Canvas):
    def __init__(self, parent, color, bg_color=None, **kw):
        kw.setdefault("height",1)
        self._bg=bg_color or C["glass"]
        super().__init__(parent,bg=self._bg,highlightthickness=0,**kw)
        self._color=color
        self.bind("<Configure>",lambda e:self._draw()); self._draw()

    def _draw(self):
        w=self.winfo_width() or 400; self.delete("all")
        for i in range(36):
            t=abs(i/36-0.5)*2
            col=lerp_color(self._color,self._bg,t**0.35)
            self.create_line(int(i/36*w),0,int((i+1)/36*w),0,fill=col,width=1)


# ── Reusable factory functions ────────────────────────────────────────────────
def GlassFrame(parent, accent=None, radius=CORNER, depth=0, **kw):
    bg=[C["glass"],C["glass2"],C["glass3"]][min(depth,2)]
    kw.setdefault("fg_color",bg)
    kw.setdefault("border_color",accent or C["border"])
    kw.setdefault("border_width",1)
    kw.setdefault("corner_radius",radius)
    return ctk.CTkFrame(parent,**kw)

_DIM = {C["violet"]:C["violet_glow"],C["teal"]:C["teal_glow"],
        C["gold"]:C["gold_glow"],C["ok"]:C["ok_dim"],
        C["error"]:C["error_dim"],C["blue"]:C["blue_glow"],
        C["warn"]:C["warn_dim"],C["amber"]:C["amber_dim"]}

def _pill_btn(parent, text, cmd, accent, height=32, width=None):
    kw=dict(text=text,fg_color=_DIM.get(accent,C["glass3"]),
            hover_color=C["glass2"],border_color=accent,border_width=1,
            text_color=accent,font=ctk.CTkFont("Segoe UI",11,"bold"),
            corner_radius=20,height=height,command=cmd)
    if width: kw["width"]=width
    return ctk.CTkButton(parent,**kw)

def _solid_btn(parent, text, cmd, accent, height=40, dark_text=False):
    return ctk.CTkButton(parent,text=text,fg_color=accent,
                         hover_color=C["border2"],border_color=accent,
                         border_width=1,
                         text_color="#050810" if dark_text else C["text"],
                         font=ctk.CTkFont("Segoe UI",12,"bold"),
                         corner_radius=20,height=height,command=cmd)

def _entry(parent, var, height=36, ph="", accent=None, mono=False):
    font = ctk.CTkFont("Cascadia Code",13) if mono else ctk.CTkFont("Segoe UI",12)
    return ctk.CTkEntry(parent,textvariable=var,fg_color=C["glass3"],
                        border_color=accent or C["border2"],
                        text_color=C["text"] if not mono else C["ok"],
                        height=height,font=font,
                        placeholder_text=ph,corner_radius=8)

def _combo(parent, var, values, height=36, width=None, accent=None):
    kw=dict(variable=var,values=values,fg_color=C["glass3"],
            button_color=C["border2"],border_color=accent or C["border2"],
            text_color=C["text"],dropdown_fg_color=C["glass2"],
            dropdown_text_color=C["text"],
            font=ctk.CTkFont("Segoe UI",12),height=height,corner_radius=8)
    if width: kw["width"]=width
    return ctk.CTkComboBox(parent,**kw)

def _slbl(parent, text, accent, padx=18, pady=(16,6)):
    fr=ctk.CTkFrame(parent,fg_color="transparent")
    fr.pack(fill="x",padx=padx,pady=pady)
    ctk.CTkFrame(fr,fg_color=accent,width=3,height=14,corner_radius=2).pack(side="left",padx=(0,8))
    ctk.CTkLabel(fr,text=text.upper(),
                 font=ctk.CTkFont("Segoe UI",10,"bold"),
                 text_color=C["text3"]).pack(side="left")

def _flbl(parent, text):
    ctk.CTkLabel(parent,text=text,
                 font=ctk.CTkFont("Segoe UI",11),
                 text_color=C["text2"]).pack(anchor="w",pady=(0,3))


# ── Rule Card ─────────────────────────────────────────────────────────────────
class RuleCard(ctk.CTkFrame):
    def __init__(self, parent, rule, index, on_select, selected=False):
        sev   = rule.get("severity","INFO")
        color = SEV_COLOR.get(sev,C["info"])
        dim   = SEV_DIM.get(sev,C["info_dim"])
        super().__init__(parent,
                         fg_color=C["glass2"] if selected else C["glass"],
                         corner_radius=10,
                         border_color=color if selected else C["border"],
                         border_width=2 if selected else 1)
        self.pack(fill="x",pady=(0,5))
        ctk.CTkFrame(self,fg_color=color,height=2,corner_radius=1).pack(fill="x",side="top")
        body=ctk.CTkFrame(self,fg_color="transparent")
        body.pack(fill="x",padx=12,pady=10)
        # badge
        badge=ctk.CTkFrame(body,fg_color=dim,corner_radius=6)
        badge.pack(side="left",padx=(0,12))
        ctk.CTkLabel(badge,text=sev,
                     font=ctk.CTkFont("Segoe UI",10,"bold"),
                     text_color=color,padx=10,pady=4).pack()
        # info
        info=ctk.CTkFrame(body,fg_color="transparent")
        info.pack(side="left",fill="x",expand=True)
        ctk.CTkLabel(info,text=rule.get("rule_id","UNNAMED"),
                     font=ctk.CTkFont("Segoe UI",12,"bold"),
                     text_color=C["text"],anchor="w").pack(anchor="w")
        cond=rule.get("condition","")[:52]
        if len(rule.get("condition",""))>52: cond+="…"
        ctk.CTkLabel(info,text=cond or "No condition",
                     font=ctk.CTkFont("Cascadia Code",10),
                     text_color=C["text3"],anchor="w").pack(anchor="w")
        # chip
        ttype=rule.get("target_type","ALL")
        if ttype!="ALL":
            chip=ctk.CTkFrame(body,fg_color=C["glass3"],corner_radius=6)
            chip.pack(side="right")
            ctk.CTkLabel(chip,text=ttype.replace("_"," "),
                         font=ctk.CTkFont("Segoe UI",10),
                         text_color=C["text2"],padx=8,pady=3).pack()
        cb = lambda e,i=index: on_select(i)
        self.bind("<Button-1>",cb)
        for c in self.winfo_children():
            c.bind("<Button-1>",cb)
            for gc in c.winfo_children():
                gc.bind("<Button-1>",cb)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MAIN APP
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class RulesEditor(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SIRAAL  |  Factory Rules Engine")
        self.geometry("1160x760")
        self.minsize(960,620)
        self.configure(fg_color=C["bg"])
        self.rules=[]; self.current_role="Viewer"
        self.selected_index=-1; self.form_widgets=[]
        self._sev_filter_val="ALL"
        self._load_rules_from_disk()
        self._build_ambient()
        self._build_header()
        self._build_body()
        self._build_statusbar()
        self._refresh_rule_list()
        self._enforce_access_control()
        self.protocol("WM_DELETE_WINDOW",self._on_close)

    # ── Data ──────────────────────────────────────────────────────────────────
    def _load_rules_from_disk(self):
        if not os.path.exists(RULES_FILE):
            self.rules=[]; self._save_rules_to_disk(); return
        try:
            with open(RULES_FILE,"r") as f:
                self.rules=json.load(f).get("rules",[])
        except Exception as e:
            messagebox.showerror("Error",f"Failed to load: {e}"); self.rules=[]

    def _save_rules_to_disk(self):
        try:
            with open(RULES_FILE,"w") as f:
                json.dump({"rules":self.rules},f,indent=4)
        except Exception as e:
            messagebox.showerror("Error",f"Failed to save: {e}")

    # ── Ambient ───────────────────────────────────────────────────────────────
    def _build_ambient(self):
        self._amb=AmbientCanvas(self,bg=C["bg"])
        self._amb.place(x=0,y=0,relwidth=1,relheight=1)

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self):
        hdr=ctk.CTkFrame(self,fg_color=C["surface"],corner_radius=0,height=64)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkFrame(hdr,fg_color=C["violet"],height=2,corner_radius=0).pack(fill="x",side="top")
        inner=ctk.CTkFrame(hdr,fg_color="transparent")
        inner.pack(fill="both",expand=True,padx=20)
        # logo
        ib=ctk.CTkFrame(inner,fg_color=C["violet_dim"],corner_radius=10,width=40,height=40)
        ib.pack(side="left",pady=10); ib.pack_propagate(False)
        ctk.CTkLabel(ib,text="🛠️",font=ctk.CTkFont("Segoe UI Emoji",20),
                     text_color=C["violet"]).pack(expand=True)
        # title
        nf=ctk.CTkFrame(inner,fg_color="transparent")
        nf.pack(side="left",padx=(14,0),pady=10)
        ctk.CTkLabel(nf,text="Factory Rules Engine",
                     font=ctk.CTkFont("Segoe UI",17,"bold"),
                     text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(nf,text="ISO 9001 / AS9100  ·  Audit Compliant  ·  TN-IMPACT 2026",
                     font=ctk.CTkFont("Segoe UI",10),
                     text_color=C["text3"]).pack(anchor="w")
        # right controls
        right=ctk.CTkFrame(inner,fg_color="transparent")
        right.pack(side="right",pady=12)
        role_fr=GlassFrame(right,accent=C["border2"],radius=20,depth=1)
        role_fr.pack(side="left",padx=(0,16))
        ri=ctk.CTkFrame(role_fr,fg_color="transparent")
        ri.pack(padx=14,pady=8)
        ctk.CTkLabel(ri,text="ROLE",font=ctk.CTkFont("Segoe UI",10,"bold"),
                     text_color=C["text3"]).pack(side="left",padx=(0,8))
        self.role_var=ctk.StringVar(value="Viewer")
        self.role_lbl=ctk.CTkLabel(ri,textvariable=self.role_var,
                                    font=ctk.CTkFont("Segoe UI",11,"bold"),
                                    text_color=C["text2"])
        self.role_lbl.pack(side="left",padx=(0,10))
        self._btn_role=_pill_btn(ri,"Elevate →",self._on_elevate,C["violet"],height=28,width=96)
        self._btn_role.pack(side="left")
        for txt,acc in [("IS 9001",C["teal"]),("AS9100",C["gold"]),("Audit Log",C["violet"])]:
            glow={C["teal"]:C["teal_glow"],C["gold"]:C["gold_glow"],C["violet"]:C["violet_glow"]}[acc]
            fr=ctk.CTkFrame(right,fg_color=glow,corner_radius=6,border_color=acc,border_width=1)
            fr.pack(side="left",padx=3)
            ctk.CTkLabel(fr,text=txt,font=ctk.CTkFont("Segoe UI",10,"bold"),
                         text_color=acc,padx=10,pady=4).pack()
        GlowSep(self,C["border2"],bg_color=C["surface"]).pack(fill="x")

    # ── Body ──────────────────────────────────────────────────────────────────
    def _build_body(self):
        body=ctk.CTkFrame(self,fg_color="transparent")
        body.pack(fill="both",expand=True,padx=16,pady=14)
        body.columnconfigure(0,weight=0,minsize=324)
        body.columnconfigure(1,weight=1)
        body.rowconfigure(0,weight=1)
        lp=GlassFrame(body,accent=C["border2"],radius=14,width=324)
        lp.grid(row=0,column=0,sticky="nsew",padx=(0,12)); lp.pack_propagate(False)
        self._build_left(lp)
        rp=GlassFrame(body,accent=C["border2"],radius=14)
        rp.grid(row=0,column=1,sticky="nsew")
        self._build_right(rp)

    # ── Left ──────────────────────────────────────────────────────────────────
    def _build_left(self,p):
        ctk.CTkFrame(p,fg_color=C["violet"],height=2,corner_radius=1).pack(fill="x",side="top")
        hdr=ctk.CTkFrame(p,fg_color="transparent")
        hdr.pack(fill="x",padx=14,pady=(14,6))
        ib=ctk.CTkFrame(hdr,fg_color=C["violet_dim"],corner_radius=8,width=34,height=34)
        ib.pack(side="left"); ib.pack_propagate(False)
        ctk.CTkLabel(ib,text="📋",font=ctk.CTkFont("Segoe UI Emoji",16),
                     text_color=C["violet"]).pack(expand=True)
        tl=ctk.CTkFrame(hdr,fg_color="transparent"); tl.pack(side="left",padx=(10,0))
        ctk.CTkLabel(tl,text="Active Rules",font=ctk.CTkFont("Segoe UI",14,"bold"),
                     text_color=C["text"]).pack(anchor="w")
        self._count_lbl=ctk.CTkLabel(tl,text="0 rules defined",
                                      font=ctk.CTkFont("Segoe UI",10),
                                      text_color=C["text3"])
        self._count_lbl.pack(anchor="w")
        GlowSep(p,C["border2"],bg_color=C["glass"]).pack(fill="x")
        _slbl(p,"Search",C["violet"])
        self._search_var=ctk.StringVar()
        self._search_var.trace_add("write",lambda *_:self._refresh_rule_list())
        _entry(p,self._search_var,height=32,ph="🔍  Filter by ID or type…"
               ).pack(fill="x",padx=14,pady=(0,10))
        # sev chips
        cf=ctk.CTkFrame(p,fg_color="transparent"); cf.pack(fill="x",padx=14,pady=(0,10))
        ctk.CTkLabel(cf,text="Filter:",font=ctk.CTkFont("Segoe UI",10),
                     text_color=C["text3"]).pack(side="left",padx=(0,6))
        for sev,col in [("ALL",C["text2"]),("ERROR",C["error"]),
                         ("WARN",C["warn"]),("INFO",C["info"])]:
            ctk.CTkButton(cf,text=sev,width=56,height=26,
                          fg_color=SEV_DIM.get(sev,C["glass3"]) if sev!="ALL" else C["glass3"],
                          hover_color=C["glass2"],
                          border_color=col,border_width=1,text_color=col,
                          font=ctk.CTkFont("Segoe UI",10,"bold"),corner_radius=12,
                          command=lambda s=sev:self._set_sev_filter(s)
                          ).pack(side="left",padx=(0,4))
        self._list_sf=ctk.CTkScrollableFrame(p,fg_color="transparent",
                                              scrollbar_button_color=C["border2"])
        self._list_sf.pack(fill="both",expand=True,padx=8,pady=(0,8))
        GlowSep(p,C["border2"],bg_color=C["glass"]).pack(fill="x")
        self._btn_new=_solid_btn(p,"＋  New Rule",self._on_new_rule,C["violet"],height=40)
        self._btn_new.pack(fill="x",padx=14,pady=12)

    # ── Right ─────────────────────────────────────────────────────────────────
    def _build_right(self,p):
        ctk.CTkFrame(p,fg_color=C["teal"],height=2,corner_radius=1).pack(fill="x",side="top")
        hdr=ctk.CTkFrame(p,fg_color="transparent")
        hdr.pack(fill="x",padx=18,pady=(14,6))
        ib=ctk.CTkFrame(hdr,fg_color=C["teal_dim"],corner_radius=8,width=34,height=34)
        ib.pack(side="left"); ib.pack_propagate(False)
        ctk.CTkLabel(ib,text="⚙",font=ctk.CTkFont("Segoe UI Emoji",16),
                     text_color=C["teal"]).pack(expand=True)
        tl=ctk.CTkFrame(hdr,fg_color="transparent"); tl.pack(side="left",padx=(10,0))
        ctk.CTkLabel(tl,text="Rule Configuration",
                     font=ctk.CTkFont("Segoe UI",14,"bold"),text_color=C["text"]).pack(anchor="w")
        self._editor_sub=ctk.CTkLabel(tl,text="Select a rule or create a new one",
                                       font=ctk.CTkFont("Segoe UI",10),text_color=C["text3"])
        self._editor_sub.pack(anchor="w")
        GlowSep(p,C["border2"],bg_color=C["glass"]).pack(fill="x")
        sf=ctk.CTkScrollableFrame(p,fg_color="transparent",scrollbar_button_color=C["border2"])
        sf.pack(fill="both",expand=True)
        self._build_form(sf)
        GlowSep(p,C["border2"],bg_color=C["glass"]).pack(fill="x",side="bottom")
        act=ctk.CTkFrame(p,fg_color=C["glass2"],corner_radius=0,height=68)
        act.pack(fill="x",side="bottom"); act.pack_propagate(False)
        self._build_actions(act)

    # ── Form ──────────────────────────────────────────────────────────────────
    def _build_form(self,p):
        # Identity
        _slbl(p,"Identity",C["teal"])
        r1=ctk.CTkFrame(p,fg_color="transparent"); r1.pack(fill="x",padx=18,pady=(0,12))
        id_fr=ctk.CTkFrame(r1,fg_color="transparent")
        id_fr.pack(side="left",padx=(0,16),fill="x",expand=True)
        _flbl(id_fr,"Rule ID")
        self.v_id=ctk.StringVar()
        e_id=_entry(id_fr,self.v_id,height=38,ph="e.g.  SHOP_001",accent=C["teal"])
        e_id.pack(fill="x"); self.form_widgets.append(e_id)
        sev_fr=ctk.CTkFrame(r1,fg_color="transparent"); sev_fr.pack(side="left")
        _flbl(sev_fr,"Severity")
        self.v_sev=ctk.StringVar(value="ERROR")
        sc=_combo(sev_fr,self.v_sev,SEVERITIES,height=38,width=170,accent=C["teal"])
        sc.pack(); self.form_widgets.append(sc)
        # Scope
        _slbl(p,"Scope",C["teal"])
        r2=ctk.CTkFrame(p,fg_color="transparent"); r2.pack(fill="x",padx=18,pady=(0,12))
        tf=ctk.CTkFrame(r2,fg_color="transparent")
        tf.pack(side="left",padx=(0,16),fill="x",expand=True)
        _flbl(tf,"Target Part Type")
        self.v_type=ctk.StringVar(value="ALL")
        tc=_combo(tf,self.v_type,GEAR_TYPES,height=38,accent=C["teal"])
        tc.pack(fill="x"); self.form_widgets.append(tc)
        mf=ctk.CTkFrame(r2,fg_color="transparent"); mf.pack(side="left")
        _flbl(mf,"Target Material")
        self.v_mat=ctk.StringVar(value="ALL")
        mc=_combo(mf,self.v_mat,MATERIALS,height=38,width=200,accent=C["teal"])
        mc.pack(); self.form_widgets.append(mc)
        # Condition
        _slbl(p,"Condition  (Python Expression)",C["teal"])
        cc=GlassFrame(p,accent=C["teal_dim"],depth=1,radius=10)
        cc.pack(fill="x",padx=18,pady=(0,8))
        hint=ctk.CTkFrame(cc,fg_color=C["teal_glow"],corner_radius=0)
        hint.pack(fill="x")
        ctk.CTkLabel(hint,
                     text="  Variables:  P1 (Teeth)  ·  P2 (Module)  ·  "
                          "P3 (Face Width)  ·  P4 (Bore)  ·  QTY",
                     font=ctk.CTkFont("Cascadia Code",10),
                     text_color=C["teal"],anchor="w").pack(fill="x",padx=10,pady=6)
        self.v_cond=ctk.StringVar()
        ce=ctk.CTkEntry(cc,textvariable=self.v_cond,
                        fg_color=C["glass3"],border_color=C["border"],
                        text_color=C["ok"],
                        font=ctk.CTkFont("Cascadia Code",14),
                        height=46,corner_radius=0,
                        placeholder_text="P3 > 150 and P2 < 2")
        ce.pack(fill="x"); self.form_widgets.append(ce)
        ex_fr=ctk.CTkFrame(cc,fg_color="transparent"); ex_fr.pack(fill="x",padx=10,pady=8)
        ctk.CTkLabel(ex_fr,text="EXAMPLES:",font=ctk.CTkFont("Segoe UI",10,"bold"),
                     text_color=C["text3"]).pack(side="left",padx=(2,10))
        for ex in ["P3 > 150","P2 < 2 and P1 > 80","QTY > 50"]:
            ctk.CTkButton(ex_fr,text=ex,height=26,
                          fg_color=C["glass3"],hover_color=C["glass2"],
                          border_color=C["border2"],border_width=1,
                          text_color=C["teal"],
                          font=ctk.CTkFont("Cascadia Code",10),corner_radius=6,
                          command=lambda t=ex:self.v_cond.set(t)
                          ).pack(side="left",padx=(0,6))
        # Message
        _slbl(p,"Violation Message",C["teal"])
        mc2=GlassFrame(p,accent=C["teal_dim"],depth=1,radius=10)
        mc2.pack(fill="x",padx=18,pady=(0,10))
        self.t_msg=tk.Text(mc2,bg=C["glass3"],fg=C["text"],
                           font=("Segoe UI",12),relief="flat",
                           padx=12,pady=10,height=4,wrap="word",
                           insertbackground=C["teal"],
                           selectbackground=C["teal_dim"])
        self.t_msg.pack(fill="both",expand=True,padx=1,pady=1)
        self.form_widgets.append(self.t_msg)
        # Audit preview
        _slbl(p,"Last Audit Event",C["violet"])
        ac=GlassFrame(p,accent=C["violet_dim"],depth=1,radius=10)
        ac.pack(fill="x",padx=18,pady=(0,18))
        self._audit_lbl=ctk.CTkLabel(ac,text="[--:--:--]  No action recorded yet",
                                      font=ctk.CTkFont("Cascadia Code",10),
                                      text_color=C["text3"],anchor="w",justify="left")
        self._audit_lbl.pack(fill="x",padx=14,pady=10)

    def _build_actions(self,p):
        inner=ctk.CTkFrame(p,fg_color="transparent"); inner.pack(fill="both",expand=True,padx=18)
        self._btn_del=ctk.CTkButton(inner,text="🗑  Delete Rule",
                                     fg_color=C["error_dim"],hover_color=C["error"],
                                     border_color=C["error"],border_width=1,
                                     text_color=C["error"],
                                     font=ctk.CTkFont("Segoe UI",11,"bold"),
                                     corner_radius=20,height=38,width=150,
                                     command=self._on_delete)
        self._btn_del.pack(side="left",pady=14)
        self._btn_save=_solid_btn(inner,"💾  Save Rule",self._on_save,
                                   C["teal"],height=38,dark_text=True)
        self._btn_save.pack(side="right",pady=14,ipadx=20)

    # ── Status Bar ────────────────────────────────────────────────────────────
    def _build_statusbar(self):
        GlowSep(self,C["border2"],bg_color=C["surface"]).pack(fill="x",side="bottom")
        bar=ctk.CTkFrame(self,fg_color=C["surface"],corner_radius=0,height=38)
        bar.pack(fill="x",side="bottom"); bar.pack_propagate(False)
        inner=ctk.CTkFrame(bar,fg_color="transparent"); inner.pack(fill="both",expand=True,padx=16)
        self._sdot=ctk.CTkLabel(inner,text="●",font=ctk.CTkFont("Segoe UI",12),text_color=C["teal"])
        self._sdot.pack(side="left",padx=(0,6),pady=8)
        self._slbl=ctk.CTkLabel(inner,text="Rules Editor  ·  Viewer Mode  ·  Read-Only",
                                 font=ctk.CTkFont("Segoe UI",10),text_color=C["text3"])
        self._slbl.pack(side="left")
        self._rchip=ctk.CTkFrame(inner,fg_color=C["glass3"],corner_radius=6)
        self._rchip.pack(side="right",pady=8)
        self._rchip_lbl=ctk.CTkLabel(self._rchip,text="VIEWER",
                                      font=ctk.CTkFont("Segoe UI",10,"bold"),
                                      text_color=C["text3"],padx=12,pady=3)
        self._rchip_lbl.pack()

    def _set_status(self,msg,color=None):
        self._slbl.configure(text=msg,text_color=color or C["text3"])

    # ── Access Control ────────────────────────────────────────────────────────
    def _on_elevate(self):
        if self.current_role=="Admin":
            log_event("ADMIN","LOGOUT","Reverted to Viewer.")
            self.current_role="Viewer"; self.role_var.set("Viewer")
            self._btn_role.configure(text="Elevate →")
            self._set_status("Reverted to Viewer mode — read-only",C["text3"])
        else:
            pwd=simpledialog.askstring("Admin Access","Enter Admin Password:",show="*")
            if pwd==ADMIN_PASSWORD:
                self.current_role="Admin"; self.role_var.set("Admin")
                self._btn_role.configure(text="Demote ↓")
                log_event("ADMIN","LOGIN_SUCCESS","Elevated to Admin.")
                self._set_status("Admin access granted — full edit enabled",C["ok"])
                self._flash_audit("LOGIN  Admin session started")
            else:
                log_event("VIEWER","LOGIN_FAILED","Bad password.",is_warning=True)
                messagebox.showerror("Denied","Incorrect password.")
                self._set_status("Access denied — incorrect password",C["error"])
        self._enforce_access_control()

    def _enforce_access_control(self):
        is_admin=(self.current_role=="Admin")
        state="normal" if is_admin else "disabled"
        for w in self.form_widgets:
            try: w.configure(state=state)
            except: pass
        self._btn_new.configure(state=state)
        self._btn_save.configure(state=state)
        self._btn_del.configure(state=state)
        if is_admin:
            self._rchip.configure(fg_color=C["violet_dim"])
            self._rchip_lbl.configure(text="ADMIN",text_color=C["violet"])
            self.role_lbl.configure(text_color=C["violet"])
            self._sdot.configure(text_color=C["violet"])
        else:
            self._rchip.configure(fg_color=C["glass3"])
            self._rchip_lbl.configure(text="VIEWER",text_color=C["text3"])
            self.role_lbl.configure(text_color=C["text2"])
            self._sdot.configure(text_color=C["teal"])

    def _flash_audit(self,msg):
        ts=datetime.datetime.now().strftime("%H:%M:%S")
        self._audit_lbl.configure(text=f"[{ts}]  {msg}",text_color=C["violet"])

    # ── Rule List ─────────────────────────────────────────────────────────────
    def _set_sev_filter(self,sev):
        self._sev_filter_val=sev; self._refresh_rule_list()

    def _refresh_rule_list(self,*_):
        for c in self._list_sf.winfo_children(): c.destroy()
        query=self._search_var.get().lower() if hasattr(self,"_search_var") else ""
        visible=[]
        for i,rule in enumerate(self.rules):
            sev=rule.get("severity","INFO")
            f=self._sev_filter_val
            if f not in ("ALL",""):
                if f=="WARN":
                    if sev!="WARNING": continue
                elif sev!=f: continue
            if query and query not in rule.get("rule_id","").lower() \
                     and query not in rule.get("target_type","").lower(): continue
            visible.append((i,rule))
        total=len(self.rules); count=len(visible)
        self._count_lbl.configure(
            text=f"{count} rule{'s' if count!=1 else ''}"
                 +(f"  (of {total})" if count!=total else " defined"))
        if not visible:
            ef=ctk.CTkFrame(self._list_sf,fg_color="transparent")
            ef.pack(expand=True,pady=40)
            ctk.CTkLabel(ef,text="No rules match",
                         font=ctk.CTkFont("Segoe UI",12,"italic"),
                         text_color=C["text3"]).pack()
            return
        for ri,rule in visible:
            RuleCard(self._list_sf,rule,ri,on_select=self._select_rule,
                     selected=(ri==self.selected_index))

    def _select_rule(self,index):
        self.selected_index=index; rule=self.rules[index]
        self._editor_sub.configure(text=f"Editing  ›  {rule.get('rule_id','?')}",
                                    text_color=C["teal"])
        for w in self.form_widgets:
            try: w.configure(state="normal")
            except: pass
        self.v_id.set(rule.get("rule_id",""))
        self.v_type.set(rule.get("target_type","ALL"))
        self.v_mat.set(rule.get("target_material","ALL"))
        self.v_cond.set(rule.get("condition",""))
        self.v_sev.set(rule.get("severity","ERROR"))
        self.t_msg.delete("1.0","end"); self.t_msg.insert("end",rule.get("message",""))
        self._enforce_access_control(); self._refresh_rule_list()
        self._flash_audit(f"VIEW  Rule '{rule.get('rule_id','')}' opened")

    # ── CRUD ──────────────────────────────────────────────────────────────────
    def _on_new_rule(self):
        self.selected_index=-1
        self._editor_sub.configure(text="Creating new rule…",text_color=C["violet"])
        for w in self.form_widgets:
            try: w.configure(state="normal")
            except: pass
        self.v_id.set("NEW_RULE_001"); self.v_type.set("ALL"); self.v_mat.set("ALL")
        self.v_cond.set("P3 > 100"); self.v_sev.set("ERROR")
        self.t_msg.delete("1.0","end"); self.t_msg.insert("end","Describe the violation here…")
        self._enforce_access_control()
        self._flash_audit("NEW  Blank rule template loaded")

    def _on_save(self):
        if self.current_role!="Admin": return
        r_id=self.v_id.get().strip()
        if not r_id:
            messagebox.showwarning("Validation","Rule ID cannot be empty."); return
        new_rule={"rule_id":r_id,"target_type":self.v_type.get(),
                  "target_material":self.v_mat.get(),
                  "condition":self.v_cond.get().strip(),
                  "severity":self.v_sev.get(),
                  "message":self.t_msg.get("1.0","end").strip()}
        if self.selected_index>=0:
            self.rules[self.selected_index]=new_rule
        else:
            self.rules.append(new_rule); self.selected_index=len(self.rules)-1
        self._save_rules_to_disk()
        log_event(self.current_role,"RULE_SAVED",f"Saved: {r_id}")
        self._refresh_rule_list()
        self._set_status(f"Rule '{r_id}' saved successfully",C["ok"])
        self._editor_sub.configure(text=f"Saved  ›  {r_id}",text_color=C["ok"])
        self._flash_audit(f"SAVE  Rule '{r_id}' committed to disk")

    def _on_delete(self):
        if self.current_role!="Admin": return
        if self.selected_index<0 or self.selected_index>=len(self.rules): return
        r_id=self.rules[self.selected_index].get("rule_id","Unknown")
        if messagebox.askyesno("Confirm Delete",f"Delete rule '{r_id}'?\nThis cannot be undone."):
            self.rules.pop(self.selected_index)
            self._save_rules_to_disk()
            log_event(self.current_role,"RULE_DELETED",f"Deleted: {r_id}",is_warning=True)
            self.selected_index=-1; self._on_new_rule(); self._refresh_rule_list()
            self._set_status(f"Rule '{r_id}' deleted",C["warn"])
            self._flash_audit(f"DELETE  Rule '{r_id}' permanently removed")

    def _on_close(self):
        if hasattr(self,"_amb"): self._amb.stop()
        self.destroy()


if __name__=="__main__":
    app=RulesEditor()
    app.mainloop()