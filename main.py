"""
siraal_hub.py  —  Siraal Master Command Center
==============================================
The central launcher for the Siraal Manufacturing Engine.
Spawns 2D and 3D GUIs as independent processes to prevent Tkinter thread-locking.
"""

import customtkinter as ctk
import subprocess
import sys
import os
import json
from tkinter import messagebox

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DESIGN SYSTEM (Matching your existing theme)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ctk.set_appearance_mode("Dark")

C = {
    "base":      "#0D1117", "surface":   "#111820",
    "card":      "#1C2534", "border":    "#2E3E52",
    "gold":      "#F0B429", "teal":      "#1ABC9C",
    "violet":    "#8B5CF6", "text":      "#E8EDF2",
    "text2":     "#8B99AA", "text3":     "#5A6A7A",
}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MASTER HUB APP
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class SiraalHub(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SIRAAL  |  Command Center")
        self.geometry("600x480")
        self.resizable(False, False)
        self.configure(fg_color=C["base"])
        
        # Center the window on the screen
        self.eval('tk::PlaceWindow . center')

        self._build_ui()
        self._ensure_rules_file()

    def _build_ui(self):
        # Header
        hdr = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0, height=80)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        
        ctk.CTkLabel(hdr, text="⚙", font=ctk.CTkFont("Segoe UI", 36), text_color=C["gold"]).pack(side="left", padx=(30, 10))
        
        title_fr = ctk.CTkFrame(hdr, fg_color="transparent")
        title_fr.pack(side="left", pady=15)
        ctk.CTkLabel(title_fr, text="SIRAAL COMMAND CENTER", font=ctk.CTkFont("Segoe UI", 20, "bold"), text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(title_fr, text="TN-IMPACT 2026  |  Master App Launcher", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"]).pack(anchor="w")

        ctk.CTkFrame(self, fg_color=C["border"], height=2, corner_radius=0).pack(fill="x")

        # Main Body
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=40, pady=40)

        # BUTTON 1: 2D ENGINE
        btn_2d = ctk.CTkButton(body, text="📐  WORK ON 2D", height=60, 
                               fg_color=C["card"], hover_color=C["border"], border_color=C["teal"], border_width=2,
                               text_color=C["teal"], font=ctk.CTkFont("Segoe UI", 16, "bold"), 
                               command=self._launch_2d)
        btn_2d.pack(fill="x", pady=(0, 15))

        # BUTTON 2: 3D ENGINE
        btn_3d = ctk.CTkButton(body, text="🧊  WORK ON 3D", height=60, 
                               fg_color=C["card"], hover_color=C["border"], border_color=C["gold"], border_width=2,
                               text_color=C["gold"], font=ctk.CTkFont("Segoe UI", 16, "bold"), 
                               command=self._launch_3d)
        btn_3d.pack(fill="x", pady=(0, 15))

        # Divider
        div = ctk.CTkFrame(body, fg_color="transparent", height=30)
        div.pack(fill="x", pady=(10, 10))
        ctk.CTkFrame(div, fg_color=C["border"], height=1).place(relx=0, rely=0.5, relwidth=1.0)
        ctk.CTkLabel(div, text="  ADMINISTRATION  ", font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=C["text3"], fg_color=C["base"]).place(relx=0.5, rely=0.5, anchor="center")

        # BUTTON 3: RULES EDITOR
        btn_rules = ctk.CTkButton(body, text="🛠️  MODIFY FACTORY RULES", height=50, 
                                  fg_color=C["card"], hover_color=C["border"], border_color=C["violet"], border_width=1,
                                  text_color=C["violet"], font=ctk.CTkFont("Segoe UI", 14, "bold"), 
                                  command=self._launch_rules)
        btn_rules.pack(fill="x")

    # ── Launcher Logic ────────────────────────────────────────────────────────
    def _launch_2d(self):
        """Launches the 2D GUI as an independent background process."""
        if not os.path.exists("gui_launcher.py"):
            messagebox.showerror("Error", "gui_launcher.py (2D) not found in directory.")
            return
        # sys.executable ensures it uses your current 'uv' virtual environment!
        subprocess.Popen([sys.executable, "gui_launcher.py"])

    def _launch_3d(self):
        """Launches the 3D GUI as an independent background process."""
        if not os.path.exists("gui_launcher_3d.py"):
            messagebox.showerror("Error", "gui_launcher_3d.py not found in directory.")
            return
        subprocess.Popen([sys.executable, "gui_launcher_3d.py"])

    def _ensure_rules_file(self):
        """Creates the custom_rules.json file if it doesn't exist yet."""
        if not os.path.exists("custom_rules.json"):
            default_rules = {
                "rules": [
                    {
                        "rule_id": "EXAMPLE_RULE_01",
                        "target_type": "Spur_Gear_3D",
                        "target_material": "ALL",
                        "condition": "P3 > 200",
                        "severity": "ERROR",
                        "message": "Face width (P3) exceeds maximum 200mm limit for our machines."
                    }
                ]
            }
            with open("custom_rules.json", "w") as f:
                json.dump(default_rules, f, indent=4)

    def _launch_rules(self):
        """Opens the JSON rulebook in the system's default text editor."""
        try:
            # This automatically opens the file in VS Code or Notepad on Windows
            os.startfile("custom_rules.json")
        except AttributeError:
            # Fallback for Mac/Linux if needed later
            subprocess.call(['open', 'custom_rules.json'])

if __name__ == "__main__":
    app = SiraalHub()
    app.mainloop()