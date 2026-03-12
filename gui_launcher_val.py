"""
rules_editor_gui.py  —  Siraal Dynamic Validation Editor
========================================================
GUI for managing factory constraints (custom_rules.json).
Features Role-Based Access Control (Admin vs Viewer).
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import simpledialog, messagebox
import json
import os

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DESIGN SYSTEM & CONSTANTS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

C = {
    "void":      "#080C10", "base":      "#0D1117", "surface":   "#111820",
    "elevated":  "#161E28", "card":      "#1C2534", "border":    "#2E3E52",
    "gold":      "#F0B429", "teal":      "#1ABC9C", "violet":    "#8B5CF6",
    "error":     "#EF4444", "warn":      "#F59E0B", "ok":        "#22C55E",
    "text":      "#E8EDF2", "text2":     "#8B99AA", "text3":     "#5A6A7A",
}

RULES_FILE = "custom_rules.json"
ADMIN_PASSWORD = "admin" # Hardcoded for the hackathon prototype

# --- NEW: Dynamically fetch all base types + custom templates ---
def get_all_target_types():
    base_types = [
        "ALL", "Spur_Gear_3D", "Helical_Gear", "Ring_Gear_3D", "Bevel_Gear", 
        "Worm", "Worm_Wheel", "Box", "Cylinder", "Flange", "Stepped_Shaft", "L_Bracket"
    ]
    customs = []
    t_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    if os.path.exists(t_path):
        for f in os.listdir(t_path):
            if f.startswith("Custom_") and f.endswith(".json"):
                customs.append(f.replace(".json", ""))
    return base_types + sorted(customs)

GEAR_TYPES = get_all_target_types()
MATERIALS  = ["ALL", "Steel-1020", "Steel-4140", "Al-6061", "Brass-C360", "Nylon-66", "Ti-6Al-4V"]
SEVERITIES = ["ERROR", "WARNING", "INFO"]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MAIN APPLICATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class RulesEditor(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SIRAAL  |  Factory Rules Engine")
        self.geometry("950x650")
        self.minsize(800, 500)
        self.configure(fg_color=C["base"])
        
        self.rules = []
        self.current_role = "Viewer" # Default to read-only
        self.selected_index = -1
        
        self._load_rules_from_disk()
        self._build_ui()
        self._refresh_rule_list()
        self._enforce_access_control()

    # ── DATA HANDLING ────────────────────────────────────────────────────────
    def _load_rules_from_disk(self):
        if not os.path.exists(RULES_FILE):
            # Create default if missing
            self.rules = []
            self._save_rules_to_disk()
        else:
            try:
                with open(RULES_FILE, "r") as f:
                    data = json.load(f)
                    self.rules = data.get("rules", [])
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load rules: {e}")
                self.rules = []

    def _save_rules_to_disk(self):
        try:
            with open(RULES_FILE, "w") as f:
                json.dump({"rules": self.rules}, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save rules: {e}")

    # ── UI BUILDING ──────────────────────────────────────────────────────────
    def _build_ui(self):
        # HEADER
        hdr = ctk.CTkFrame(self, fg_color=C["void"], corner_radius=0, height=60)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        
        ctk.CTkLabel(hdr, text="🛠️ SIRAAL FACTORY RULES ENGINE", font=ctk.CTkFont("Segoe UI", 16, "bold"), text_color=C["violet"]).pack(side="left", padx=20)
        
        # Access Control Toggle
        ac_frame = ctk.CTkFrame(hdr, fg_color="transparent")
        ac_frame.pack(side="right", padx=20, pady=15)
        
        ctk.CTkLabel(ac_frame, text="Current Role:", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(side="left", padx=5)
        self.role_var = ctk.StringVar(value="Viewer")
        self.role_dropdown = ctk.CTkComboBox(ac_frame, variable=self.role_var, values=["Viewer", "Admin"], 
                                             width=100, height=28, command=self._on_role_change,
                                             fg_color=C["elevated"], border_color=C["border"])
        self.role_dropdown.pack(side="left")

        ctk.CTkFrame(self, fg_color=C["border"], height=1).pack(fill="x")

        # MAIN WORKSPACE (Split Left/Right)
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=15, pady=15)
        body.columnconfigure(0, weight=1, minsize=300) # Left List
        body.columnconfigure(1, weight=2) # Right Editor
        body.rowconfigure(0, weight=1)

        # LEFT PANEL: Rule List
        left_p = ctk.CTkFrame(body, fg_color=C["surface"], corner_radius=8)
        left_p.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        lbl_fr = ctk.CTkFrame(left_p, fg_color="transparent")
        lbl_fr.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(lbl_fr, text="Active Rules", font=ctk.CTkFont("Segoe UI", 12, "bold"), text_color=C["text"]).pack(side="left")
        
        self.btn_new = ctk.CTkButton(lbl_fr, text="＋ New Rule", width=80, height=24, fg_color=C["violet"], hover_color="#6D48C4", font=ctk.CTkFont("Segoe UI", 10, "bold"), command=self._on_new_rule)
        self.btn_new.pack(side="right")
        
        self.list_frame = ctk.CTkScrollableFrame(left_p, fg_color=C["card"], corner_radius=6)
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=(0,10))

        # RIGHT PANEL: Editor Form
        right_p = ctk.CTkFrame(body, fg_color=C["surface"], corner_radius=8)
        right_p.grid(row=0, column=1, sticky="nsew")
        
        ctk.CTkLabel(right_p, text="Rule Configuration", font=ctk.CTkFont("Segoe UI", 14, "bold"), text_color=C["text"]).pack(anchor="w", padx=20, pady=(15, 10))
        
        # Form Variables
        self.v_id = ctk.StringVar()
        self.v_type = ctk.StringVar(value="ALL")
        self.v_mat = ctk.StringVar(value="ALL")
        self.v_cond = ctk.StringVar()
        self.v_sev = ctk.StringVar(value="ERROR")
        
        self.form_widgets = [] # Keep track to disable/enable based on role
        
        # Row 1: ID & Severity
        r1 = ctk.CTkFrame(right_p, fg_color="transparent"); r1.pack(fill="x", padx=20, pady=5)
        self._build_field(r1, "Rule ID (e.g., SHOP_001):", self.v_id, 200)
        self._build_combo(r1, "Severity:", self.v_sev, SEVERITIES, 150)
        
        # Row 2: Target Type & Material
        r2 = ctk.CTkFrame(right_p, fg_color="transparent"); r2.pack(fill="x", padx=20, pady=5)
        
        # We pass GEAR_TYPES (which now includes your custom templates)
        self._build_combo(r2, "Target Part Type:", self.v_type, GEAR_TYPES, 200)
        self._build_combo(r2, "Target Material:", self.v_mat, MATERIALS, 200)
        
        # Row 3: Condition
        r3 = ctk.CTkFrame(right_p, fg_color="transparent"); r3.pack(fill="x", padx=20, pady=15)
        ctk.CTkLabel(r3, text="Logical Condition (Python math):", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(anchor="w")
        cond_entry = ctk.CTkEntry(r3, textvariable=self.v_cond, font=ctk.CTkFont("Cascadia Code", 12), fg_color=C["card"], border_color=C["border"])
        cond_entry.pack(fill="x", pady=2)
        self.form_widgets.append(cond_entry)
        
        # The font slant issue fixed here as well!
        ctk.CTkLabel(r3, text="Variables: P1 (Teeth), P2 (Module), P3 (Width), P4 (Bore), QTY. Example: P3 > 150 and P2 < 2", font=ctk.CTkFont("Segoe UI", 9, slant="italic"), text_color=C["text3"]).pack(anchor="w")
        
        # Row 4: Message
        r4 = ctk.CTkFrame(right_p, fg_color="transparent"); r4.pack(fill="both", expand=True, padx=20, pady=5)
        ctk.CTkLabel(r4, text="Violation Message:", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(anchor="w")
        self.t_msg = ctk.CTkTextbox(r4, font=ctk.CTkFont("Segoe UI", 12), fg_color=C["card"], border_color=C["border"], border_width=1)
        self.t_msg.pack(fill="both", expand=True, pady=2)
        self.form_widgets.append(self.t_msg)
        
        # Actions
        act_fr = ctk.CTkFrame(right_p, fg_color="transparent", height=60)
        act_fr.pack(fill="x", side="bottom", padx=20, pady=15)
        
        self.btn_del = ctk.CTkButton(act_fr, text="🗑 Delete Rule", width=120, height=36, fg_color="#3A1B20", hover_color="#672529", text_color=C["error"], command=self._on_delete)
        self.btn_del.pack(side="left")
        
        self.btn_save = ctk.CTkButton(act_fr, text="💾 Save Rule", width=140, height=36, fg_color=C["ok"], hover_color="#16A34A", text_color=C["void"], font=ctk.CTkFont("Segoe UI", 12, "bold"), command=self._on_save)
        self.btn_save.pack(side="right")

    def _build_field(self, parent, label, var, width):
        fr = ctk.CTkFrame(parent, fg_color="transparent")
        fr.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fr, text=label, font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(anchor="w")
        w = ctk.CTkEntry(fr, textvariable=var, width=width, fg_color=C["card"], border_color=C["border"])
        w.pack()
        self.form_widgets.append(w)
        
    def _build_combo(self, parent, label, var, values, width):
        fr = ctk.CTkFrame(parent, fg_color="transparent")
        fr.pack(side="left", padx=(0, 15))
        ctk.CTkLabel(fr, text=label, font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(anchor="w")
        w = ctk.CTkComboBox(fr, variable=var, values=values, width=width, fg_color=C["card"], border_color=C["border"], button_color=C["border"])
        w.pack()
        self.form_widgets.append(w)

    # ── LOGIC ────────────────────────────────────────────────────────────────
    def _on_role_change(self, choice):
        if choice == "Admin":
            # Prompt for password
            pwd = simpledialog.askstring("Admin Access", "Enter Admin Password:", show='*')
            if pwd == ADMIN_PASSWORD:
                self.current_role = "Admin"
                messagebox.showinfo("Success", "Admin access granted. You can now edit rules.")
            else:
                messagebox.showerror("Denied", "Incorrect password.")
                self.role_var.set("Viewer")
                self.current_role = "Viewer"
        else:
            self.current_role = "Viewer"
            
        self._enforce_access_control()

    def _enforce_access_control(self):
        state = "normal" if self.current_role == "Admin" else "disabled"
        
        for w in self.form_widgets:
            w.configure(state=state)
            
        self.btn_new.configure(state=state)
        self.btn_save.configure(state=state)
        self.btn_del.configure(state=state)

    def _refresh_rule_list(self):
        for child in self.list_frame.winfo_children():
            child.destroy()
            
        for idx, rule in enumerate(self.rules):
            sev = rule.get("severity", "INFO")
            color = C["error"] if sev == "ERROR" else C["warn"] if sev == "WARNING" else C["info"]
            
            btn = ctk.CTkButton(self.list_frame, text=f"[{sev}] {rule.get('rule_id', 'UNNAMED')}", 
                                fg_color=C["elevated"], hover_color=C["border"], text_color=color,
                                anchor="w", height=32,
                                command=lambda i=idx: self._select_rule(i))
            btn.pack(fill="x", pady=2)

    def _select_rule(self, index):
        self.selected_index = index
        rule = self.rules[index]
        
        # Temporarily enable widgets to insert data even if viewer
        for w in self.form_widgets: w.configure(state="normal")
        
        self.v_id.set(rule.get("rule_id", ""))
        self.v_type.set(rule.get("target_type", "ALL"))
        self.v_mat.set(rule.get("target_material", "ALL"))
        self.v_cond.set(rule.get("condition", ""))
        self.v_sev.set(rule.get("severity", "ERROR"))
        
        self.t_msg.delete("1.0", "end")
        self.t_msg.insert("end", rule.get("message", ""))
        
        self._enforce_access_control() # Re-apply constraints

    def _on_new_rule(self):
        self.selected_index = -1
        self.v_id.set("NEW_RULE")
        self.v_type.set("ALL")
        self.v_mat.set("ALL")
        self.v_cond.set("P3 > 100")
        self.v_sev.set("ERROR")
        self.t_msg.delete("1.0", "end")
        self.t_msg.insert("end", "Describe the violation here...")

    def _on_save(self):
        if self.current_role != "Admin": return
        
        r_id = self.v_id.get().strip()
        if not r_id:
            messagebox.showwarning("Validation", "Rule ID cannot be empty.")
            return
            
        new_rule = {
            "rule_id": r_id,
            "target_type": self.v_type.get(),
            "target_material": self.v_mat.get(),
            "condition": self.v_cond.get().strip(),
            "severity": self.v_sev.get(),
            "message": self.t_msg.get("1.0", "end").strip()
        }
        
        if self.selected_index >= 0:
            self.rules[self.selected_index] = new_rule # Update existing
        else:
            self.rules.append(new_rule) # Add new
            self.selected_index = len(self.rules) - 1
            
        self._save_rules_to_disk()
        self._refresh_rule_list()
        messagebox.showinfo("Saved", f"Rule '{r_id}' saved successfully.")

    def _on_delete(self):
        if self.current_role != "Admin": return
        if self.selected_index < 0 or self.selected_index >= len(self.rules): return
        
        r_id = self.rules[self.selected_index].get("rule_id", "Unknown")
        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete rule '{r_id}'?")
        
        if confirm:
            self.rules.pop(self.selected_index)
            self._save_rules_to_disk()
            self._on_new_rule()
            self._refresh_rule_list()

if __name__ == "__main__":
    app = RulesEditor()
    app.mainloop()