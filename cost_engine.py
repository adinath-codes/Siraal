"""
cost_engine.py — Siraal Manufacturing Economics & ESG Module
=============================================================
Calculates true manufacturing costs including raw billet sizing, 
machining time (MHR), scrap resale recovery, and Carbon Footprint (CO2).
Generates Enterprise-grade PDF Reports with TARGETED AI Value Engineering.
"""

import math
import os
import json
import requests
import logging
from typing import Dict, List, Tuple

try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

try:
    import matplotlib.pyplot as plt
    import matplotlib
    # Use Agg backend to prevent popping up window frames during generation
    matplotlib.use('Agg')
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

logger = logging.getLogger("Siraal.CostEngine")

# ── 1. ECONOMIC & ESG DATABASE ───────────────────────────────────────────────
MATERIAL_ECO_DB = {
    "Steel-1020": {"density": 7.87, "base_cost": 125.0,  "scrap_recovery": 0.15, "co2_kg": 1.8,  "machinability": 1.0},
    "Steel-4140": {"density": 7.85, "base_cost": 185.0,  "scrap_recovery": 0.15, "co2_kg": 1.9,  "machinability": 0.75}, 
    "Al-6061":    {"density": 2.70, "base_cost": 265.0,  "scrap_recovery": 0.40, "co2_kg": 16.0, "machinability": 3.0},  
    "Brass-C360": {"density": 8.50, "base_cost": 520.0,  "scrap_recovery": 0.45, "co2_kg": 2.5,  "machinability": 4.0},  
    "Nylon-66":   {"density": 1.14, "base_cost": 415.0,  "scrap_recovery": 0.05, "co2_kg": 6.5,  "machinability": 5.0},
    "Ti-6Al-4V":  {"density": 4.43, "base_cost": 3800.0, "scrap_recovery": 0.10, "co2_kg": 35.0, "machinability": 0.25}, 
}

CNC_MILL_RATE_HR = 1500.0 
BASE_MRR_CM3_MIN = 25.0 


# ── 2. CUSTOM PDF CLASS ──────────────────────────────────────────────────────
class SiraalReport(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 15)
        self.set_text_color(26, 188, 156) # Siraal Teal
        self.cell(0, 10, 'SIRAAL MANUFACTURING ENGINE', 0, 1, 'R')
        self.set_font('Helvetica', 'I', 10)
        self.set_text_color(128, 128, 128)
        self.cell(0, 5, 'Economics & ESG Production Report | TN-IMPACT 2026', 0, 1, 'R')
        self.line(10, 28, 200, 28)
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')


# ── 3. CORE ENGINE ───────────────────────────────────────────────────────────
class CostEngine:
    def __init__(self, metal_api_key: str = "", gemini_api_key: str = ""):
        self.metal_api_key = metal_api_key
        self.gemini_api_key = gemini_api_key
        self.db = {k: v.copy() for k, v in MATERIAL_ECO_DB.items()}

    def fetch_live_metal_prices(self):
        if not self.metal_api_key:
            logger.info("No Metal API key. Using offline standard pricing.")
            return False, "No Metal API key provided. Using offline standard pricing."
            
        try:
            url = f"https://api.metalpriceapi.com/v1/latest?api_key={self.metal_api_key}&base=USD&currencies=INR,ALU,XCU&unit=kg"
            r = requests.get(url, timeout=8)
            data = r.json()
            
            if data.get("success"):
                rates = data.get("rates", {})
                inr_rate = rates.get("INR", 83.0)
                msg_parts = []
                
                if "ALU" in rates and rates["ALU"] > 0:
                    alu_inr_kg = (1.0 / rates["ALU"]) * inr_rate
                    self.db["Al-6061"]["base_cost"] = round(alu_inr_kg, 2)
                    msg_parts.append(f"Al-6061: Rs.{alu_inr_kg:.2f}/kg")
                    
                if "XCU" in rates and rates["XCU"] > 0:
                    cu_inr_kg = (1.0 / rates["XCU"]) * inr_rate
                    self.db["Brass-C360"]["base_cost"] = round(cu_inr_kg * 1.2, 2)
                    msg_parts.append(f"Brass: Rs.{self.db['Brass-C360']['base_cost']:.2f}/kg")
                
                if msg_parts:
                    return True, "Live Prices Updated: " + " | ".join(msg_parts)
                return False, "API succeeded but returned no relevant metal rates."
            else:
                err = data.get("error", {}).get("info", "Unknown API Error")
                return False, f"API Error: {err}"
        except Exception as e:
            return False, f"API Fetch Failed: {e}"

    def _get_standard_billet(self, target_od: float, target_length: float) -> Tuple[float, float]:
        allowance_d = 5.0 
        allowance_l = 5.0 
        raw_od = target_od + allowance_d
        std_od = math.ceil(raw_od / 5.0) * 5.0 
        std_len = target_length + allowance_l
        return std_od, std_len

    def analyze_part(self, part: dict) -> dict:
        ptype = part.get("Part_Type", "Spur_Gear_3D")
        mat_key = part.get("Material", "Steel-1020")
        qty = int(part.get("Qty", 1))
        mat = self.db.get(mat_key, self.db["Steel-1020"])
        
        try:
            p1 = float(part.get("Param_1", 20))
            p2 = float(part.get("Param_2", 3))
            p3 = float(part.get("Param_3", 30))
            p4 = float(part.get("Param_4", 20))
        except ValueError:
            return {"error": "Invalid parameters"}

        # Volumes
        final_vol_mm3, outer_od, length = 0.0, 0.0, p3
        
        if ptype in ("Spur_Gear_3D", "Helical_Gear"):
            outer_od = p1 * p2 + 2 * p2
            bore_r = p4 / 2.0
            final_vol_mm3 = math.pi * ((outer_od/2)**2 - bore_r**2) * p3
        elif ptype == "Ring_Gear_3D":
            inner_r = (p1 * p2) / 2.0 - p2
            outer_od = (inner_r + p4) * 2
            final_vol_mm3 = math.pi * ((outer_od/2)**2 - inner_r**2) * p3
        elif ptype == "Box":
            outer_od = max(p1, p2) 
            final_vol_mm3 = p1 * p2 * p3
        else:
            outer_od = p1 * 2 if p1 > 10 else 100
            final_vol_mm3 = math.pi * (p1**2) * p3

        final_vol_cm3 = final_vol_mm3 / 1000.0
        final_mass_kg = final_vol_cm3 * (mat["density"] / 1000.0)

        # Raw Billet & Scrap
        billet_od, billet_len = self._get_standard_billet(outer_od, length)
        raw_vol_cm3 = (math.pi * ((billet_od/2)**2) * billet_len) / 1000.0
        raw_mass_kg = raw_vol_cm3 * (mat["density"] / 1000.0)
        scrap_mass_kg = max(0.0, raw_mass_kg - final_mass_kg)
        buy_to_fly = round(raw_mass_kg / final_mass_kg, 2) if final_mass_kg > 0 else 0

        # Machining & Cost
        scrap_vol_cm3 = max(0.0, raw_vol_cm3 - final_vol_cm3)
        actual_mrr = BASE_MRR_CM3_MIN * mat["machinability"]
        machining_time_mins = scrap_vol_cm3 / actual_mrr
        machining_cost = (machining_time_mins / 60.0) * CNC_MILL_RATE_HR

        raw_material_cost = raw_mass_kg * mat["base_cost"]
        scrap_resale_value = scrap_mass_kg * (mat["base_cost"] * mat["scrap_recovery"])
        net_unit_cost = raw_material_cost + machining_cost - scrap_resale_value
        co2_emissions_kg = raw_mass_kg * mat["co2_kg"]

        return {
            "Part_Number": part.get("Part_Number", "UNK"),
            "Part_Type": ptype,
            "Material": mat_key,
            "Qty": qty,
            "Buy_to_Fly_Ratio": buy_to_fly,
            "Raw_Billet": f"Ø{int(billet_od)}mm x {int(billet_len)}mm",
            "Mass_Final_kg": round(final_mass_kg, 3),
            "Mass_Scrap_kg": round(scrap_mass_kg, 3),
            "Machining_Time_mins": round(machining_time_mins, 1),
            "Cost_Raw_Material": round(raw_material_cost, 2),
            "Cost_Machining": round(machining_cost, 2),
            "Scrap_Resale_Value": round(scrap_resale_value, 2),
            "Net_Cost_Per_Unit": round(net_unit_cost, 2),
            "Total_Cost_Line_Item": round(net_unit_cost * qty, 2),
            "CO2_Emissions_kg": round(co2_emissions_kg * qty, 2)
        }

    def generate_bom_report(self, parts: List[dict]) -> dict:
        report = {
            "Total_Parts_Produced": 0, "Total_Net_Cost": 0.0, "Total_CO2_Emissions_kg": 0.0,
            "Total_Scrap_Recovered_kg": 0.0, "Total_Scrap_Value": 0.0, "Total_Machine_Hours": 0.0,
            "Line_Items": []
        }
        for p in parts:
            res = self.analyze_part(p)
            if "error" in res: continue
            
            qty = res["Qty"]
            report["Total_Parts_Produced"] += qty
            report["Total_Net_Cost"] += res["Total_Cost_Line_Item"]
            report["Total_CO2_Emissions_kg"] += res["CO2_Emissions_kg"]
            report["Total_Scrap_Recovered_kg"] += (res["Mass_Scrap_kg"] * qty)
            report["Total_Scrap_Value"] += (res["Scrap_Resale_Value"] * qty)
            report["Total_Machine_Hours"] += (res["Machining_Time_mins"] * qty) / 60.0
            report["Line_Items"].append(res)
            
        for k in ["Total_Net_Cost", "Total_CO2_Emissions_kg", "Total_Scrap_Recovered_kg", "Total_Scrap_Value", "Total_Machine_Hours"]:
            report[k] = round(report[k], 2)
            
        return report

    # ── CHART GENERATOR (MATPLOTLIB) ──────────────────────────────────────────
    def _generate_charts(self, report: dict, output_dir: str) -> tuple:
        if not MATPLOTLIB_AVAILABLE:
            return None, None
            
        pie_path = os.path.join(output_dir, "temp_cost_pie.png")
        bar_path = os.path.join(output_dir, "temp_co2_bar.png")

        # 1. Cost Breakdown Pie Chart
        total_mat = sum(item['Cost_Raw_Material'] * item['Qty'] for item in report['Line_Items'])
        total_mach = sum(item['Cost_Machining'] * item['Qty'] for item in report['Line_Items'])
        
        if total_mat == 0 and total_mach == 0: return None, None # Prevent div by zero
        
        plt.figure(figsize=(5, 5))
        plt.pie([total_mat, total_mach], labels=['Raw Material', 'Machining Overhead'], 
                autopct='%1.1f%%', startangle=140, colors=['#2980b9', '#e67e22'],
                textprops={'fontsize': 12, 'weight': 'bold'})
        plt.title("Gross Expenses Breakdown", fontsize=14, weight='bold', pad=15)
        plt.savefig(pie_path, bbox_inches='tight', dpi=150)
        plt.close()

        # 2. CO2 Emissions by Material (Bar Chart)
        mat_co2 = {}
        for item in report['Line_Items']:
            mat_co2[item['Material']] = mat_co2.get(item['Material'], 0) + item['CO2_Emissions_kg']
            
        plt.figure(figsize=(6, 4.5))
        plt.bar(mat_co2.keys(), mat_co2.values(), color='#8e44ad') # Purple
        plt.title("ESG Impact: CO2 Emissions by Material", fontsize=14, weight='bold', pad=15)
        plt.ylabel("Emissions (kg CO2)", fontsize=11, weight='bold')
        plt.xticks(rotation=15, ha='right')
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()
        plt.savefig(bar_path, bbox_inches='tight', dpi=150)
        plt.close()

        return pie_path, bar_path

    # ── TARGETED AI VALUE ENGINEERING ──────────────────────────────────────────
    def get_ai_insights(self, report_dict: dict) -> str:
        """Sends specific 'Worst Offender' data to Gemini to get targeted consulting advice."""
        if not GENAI_AVAILABLE or not self.gemini_api_key:
            return "AI Insights disabled. Gemini API Key or google-genai module missing."
            
        line_items = report_dict.get("Line_Items", [])
        if not line_items:
            return "Not enough data to provide AI insights."
            
        try:
            client = genai.Client(api_key=self.gemini_api_key)
            
            # ISOLATE WORST OFFENDERS
            highest_cost_part = max(line_items, key=lambda x: x["Total_Cost_Line_Item"])
            worst_waste_part = max(line_items, key=lambda x: x["Buy_to_Fly_Ratio"])
            highest_co2_part = max(line_items, key=lambda x: x["CO2_Emissions_kg"])
            
            avg_btf = round(sum([i["Buy_to_Fly_Ratio"] for i in line_items]) / len(line_items), 2)
            
            ai_payload = {
                "Global_Metrics": {
                    "Total_Net_Cost_Rs": report_dict["Total_Net_Cost"],
                    "Total_CO2_kg": report_dict["Total_CO2_Emissions_kg"],
                    "Average_Buy_to_Fly_Ratio": avg_btf
                },
                "Worst_Offenders": {
                    "Highest_Cost_Driver": {
                        "Part_Number": highest_cost_part["Part_Number"],
                        "Material": highest_cost_part["Material"],
                        "Total_Cost": highest_cost_part["Total_Cost_Line_Item"]
                    },
                    "Most_Wasteful_Part": {
                        "Part_Number": worst_waste_part["Part_Number"],
                        "Material": worst_waste_part["Material"],
                        "Buy_to_Fly_Ratio": worst_waste_part["Buy_to_Fly_Ratio"],
                        "Scrap_kg": worst_waste_part["Mass_Scrap_kg"]
                    },
                    "Highest_Carbon_Footprint": {
                        "Part_Number": highest_co2_part["Part_Number"],
                        "Material": highest_co2_part["Material"],
                        "CO2_kg": highest_co2_part["CO2_Emissions_kg"]
                    }
                }
            }
            
            prompt = f"""You are the Siraal Senior Value Engineering Expert & ESG Auditor.
            I have analyzed a manufacturing batch and isolated the worst offending parts:
            
            {json.dumps(ai_payload, indent=2)}
            
            Write 3 distinct paragraphs of targeted advice addressing these exact parts. 
            Do NOT give generic advice. You MUST explicitly name the Part Numbers in your response and suggest precise redesigns or material substitutions (e.g., swapping Titanium for Steel to cut costs, or proposing additive manufacturing/casting for high-scrap parts).
            
            Paragraph 1: Target the Highest Cost Driver part.
            Paragraph 2: Target the Most Wasteful part (Buy-to-Fly ratio).
            Paragraph 3: Target the Carbon Footprint (ESG) offender.
            
            CRITICAL RULES:
            - Call out the part numbers explicitly.
            - Do NOT use markdown like asterisks (**) or hashes (#). Use standard plain text.
            - Use 'Rs.' instead of the Rupee symbol.
            - Be concise, highly technical, and professional.
            """
            
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(temperature=0.4)
            )
            return response.text.strip()
        except Exception as e:
            logger.error(f"AI Insights failed: {e}")
            return "AI Insights could not be generated at this time due to an API timeout or error."

    # ── 4. PDF GENERATOR ─────────────────────────────────────────────────────
    def export_pdf_report(self, parts: List[dict], output_path: str):
        if not FPDF_AVAILABLE:
            logger.error("fpdf2 is not installed. Please run: uv pip install fpdf2")
            return False
            
        out_dir = os.path.dirname(os.path.abspath(output_path))
        os.makedirs(out_dir, exist_ok=True)
            
        pdf = SiraalReport()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        full_report = self.generate_bom_report(parts)
        if not full_report["Line_Items"]:
            return False
        
        # --- PHASE 1: INDIVIDUAL PART PAGES ---
        for item in full_report["Line_Items"]:
            pdf.add_page()
            
            pdf.set_font('Helvetica', 'B', 16); pdf.set_text_color(0, 0, 0)
            pdf.cell(0, 10, f"PART SPECIFICATION: {item['Part_Number']}", 0, 1)
            
            pdf.set_font('Helvetica', '', 12); pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 8, f"Type: {item['Part_Type']}   |   Material: {item['Material']}   |   Quantity: {item['Qty']}", 0, 1)
            pdf.ln(5)
            
            def add_row(label, value, is_bold=False, is_highlight=False):
                pdf.set_font('Helvetica', 'B' if is_bold else '', 11)
                pdf.set_text_color(0,0,0)
                pdf.set_fill_color(240, 248, 255) if is_highlight else pdf.set_fill_color(255, 255, 255)
                pdf.cell(90, 10, f"  {label}", border=1, fill=is_highlight)
                safe_val = str(value).encode('latin-1', 'replace').decode('latin-1')
                pdf.cell(90, 10, f"  {safe_val}", border=1, ln=1, fill=is_highlight)

            pdf.set_font('Helvetica', 'B', 12); pdf.set_text_color(41, 128, 185)
            pdf.cell(0, 10, "1. Manufacturing & ESG Metrics", 0, 1)
            add_row("Required Billet Size", item['Raw_Billet'])
            add_row("Buy-to-Fly Scrap Ratio", f"{item['Buy_to_Fly_Ratio']} (1.0 is ideal)")
            add_row("Final Part Mass", f"{item['Mass_Final_kg']} kg")
            add_row("Waste Scrap Mass", f"{item['Mass_Scrap_kg']} kg")
            add_row("Estimated Machining Time", f"{item['Machining_Time_mins']} mins")
            add_row("Carbon Footprint (Total Qty)", f"{item['CO2_Emissions_kg']} kg CO2", is_highlight=True)
            pdf.ln(8)
            
            pdf.set_font('Helvetica', 'B', 12); pdf.set_text_color(41, 128, 185)
            pdf.cell(0, 10, "2. Financial Breakdown (Per Unit)", 0, 1)
            add_row("Raw Material Cost", f"Rs. {item['Cost_Raw_Material']:,.2f}")
            add_row("Machining Cost (Overhead)", f"Rs. {item['Cost_Machining']:,.2f}")
            add_row("Scrap Resale Recovery", f"- Rs. {item['Scrap_Resale_Value']:,.2f}")
            add_row("Net Cost (Per Unit)", f"Rs. {item['Net_Cost_Per_Unit']:,.2f}", is_bold=True)
            
            pdf.ln(5)
            pdf.set_font('Helvetica', 'B', 14); pdf.set_text_color(231, 76, 60)
            pdf.cell(0, 12, f"Total Cost for this Line Item (Qty: {item['Qty']}): Rs. {item['Total_Cost_Line_Item']:,.2f}", border=1, ln=1, align='C', fill=False)

        # --- PHASE 2: FINAL SUMMARY PAGE ---
        pdf.add_page()
        pdf.set_font('Helvetica', 'B', 20); pdf.set_text_color(44, 62, 80)
        pdf.cell(0, 15, "EXECUTIVE PRODUCTION SUMMARY", 0, 1, 'C')
        pdf.line(10, pdf.get_y(), 200, pdf.get_y()); pdf.ln(10)
        
        def summary_block(title, value, color_rgb):
            pdf.set_font('Helvetica', 'B', 12); pdf.set_text_color(100, 100, 100)
            pdf.cell(0, 8, title, 0, 1, 'C')
            pdf.set_font('Helvetica', 'B', 24); pdf.set_text_color(*color_rgb)
            pdf.cell(0, 15, value, 0, 1, 'C'); pdf.ln(5)

        summary_block("TOTAL PRODUCTION RUN COST", f"Rs. {full_report['Total_Net_Cost']:,.2f}", (231, 76, 60))
        summary_block("TOTAL PARTS MANUFACTURED", str(full_report['Total_Parts_Produced']), (41, 128, 185))
        summary_block("TOTAL MACHINE TIME REQUIRED", f"{full_report['Total_Machine_Hours']:,.1f} Hours", (243, 156, 18))
        summary_block("TOTAL CARBON EMISSIONS", f"{full_report['Total_CO2_Emissions_kg']:,.1f} kg CO2", (142, 68, 173))
        
        pdf.ln(10)
        pdf.set_font('Helvetica', 'I', 10); pdf.set_text_color(100, 100, 100)
        pdf.multi_cell(0, 6, f"Sustainability Note: By recycling the swarf from this production run, you will recover {full_report['Total_Scrap_Recovered_kg']:,.1f} kg of raw material, injecting Rs. {full_report['Total_Scrap_Value']:,.2f} back into your operational budget.", align='C')

        # --- PHASE 3: VISUAL ANALYTICS (CHARTS) ---
        pie_path, bar_path = self._generate_charts(full_report, out_dir)
        if pie_path and bar_path:
            pdf.add_page()
            pdf.set_font('Helvetica', 'B', 20)
            pdf.set_text_color(41, 128, 185)
            pdf.cell(0, 15, "VISUAL PRODUCTION ANALYTICS", 0, 1, 'C')
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.ln(10)
            
            y_pos = pdf.get_y()
            pdf.image(pie_path, x=15, y=y_pos, w=85)
            pdf.image(bar_path, x=110, y=y_pos, w=85)
            pdf.set_y(y_pos + 90)

        # --- PHASE 4: TARGETED AI VALUE ENGINEERING INSIGHTS ---
        pdf.add_page()
        pdf.set_font('Helvetica', 'B', 20)
        pdf.set_text_color(142, 68, 173) # Purple for AI
        pdf.cell(0, 15, "TARGETED AI VALUE ENGINEERING", 0, 1, 'C')
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(10)
        
        pdf.set_font('Helvetica', 'I', 11)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 8, "Automated 'Worst-Offender' Audit by Gemini 2.5 Flash", 0, 1, 'C')
        pdf.ln(5)

        ai_advice = self.get_ai_insights(full_report)
        safe_advice = ai_advice.encode('latin-1', 'replace').decode('latin-1')
        
        pdf.set_font('Helvetica', '', 12)
        pdf.set_text_color(0, 0, 0)
        pdf.multi_cell(0, 8, safe_advice)

        # Save to disk and cleanup temp images
        pdf.output(output_path)
        logger.info(f"PDF Report successfully generated: {output_path}")
        
        if pie_path and os.path.exists(pie_path): os.remove(pie_path)
        if bar_path and os.path.exists(bar_path): os.remove(bar_path)
        
        return True

# ── Quick Test Block ────────────────────────────────────────────────────────
if __name__ == "__main__":
    m_api = os.environ.get("METALPRICE_API_KEY", "")
    g_api = os.environ.get("GEMINI_API_KEY", "")
    engine = CostEngine(metal_api_key=m_api, gemini_api_key=g_api)
    
    # Intentionally bad design to trigger the AI:
    sample_bom = [
        {
            "Part_Number": "GR-SPUR-001", "Part_Type": "Spur_Gear_3D",
            "Material": "Al-6061", "Param_1": 40, "Param_2": 3, "Param_3": 40, "Param_4": 25, "Qty": 10
        },
        {
            "Part_Number": "SH-STEP-002", "Part_Type": "Stepped_Shaft",
            "Material": "Ti-6Al-4V", "Param_1": 150, "Param_2": 50, "Param_3": 80, "Param_4": 30, "Qty": 5
        }
    ]
    
    report_path = "output_reports/Production_Cost_Report.pdf"
    engine.export_pdf_report(sample_bom, report_path)
    print(f"Test complete! Open the file at: {report_path}")