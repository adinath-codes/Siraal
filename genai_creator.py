import os
import json
from google import genai

def generate_siraal_shape(part_name, description, api_key=None, model_name="gemini-2.5-flash", log_cb=None):
    """Uses the Google GenAI SDK to generate a Siraal-compatible JSON CSG recipe."""
    
    # Custom print function that routes to the GUI if available, or terminal if standalone
    def _log(msg):
        if log_cb: log_cb(msg)
        else: print(msg)

    if part_name.startswith("Custom_"): 
        part_name = part_name[7:] # Clean it up

    _log(f"\n[AI] Thinking about how to build: Custom_{part_name}...")
    
    prompt = f"""
    You are the AI CAD Architect for the Siraal Automation Engine.
    Your job is to convert natural language descriptions into a specific JSON Constructive Solid Geometry (CSG) recipe.

    Target Part Name: Custom_{part_name}
    Description: {description}

    SIRAAL CSG RULES:
    1. Variables: You MUST use P1, P2, P3, and P4 for dimensions (these come from Excel).
    2. Shapes allowed: "box", "cylinder", "sphere".
    3. Actions allowed: "BASE" (must be step 1), "ADD", "SUBTRACT".
    4. You can use math expressions as strings (e.g., "P1/2", "P3+10").
    5. Always make subtraction tools slightly longer than the part to ensure clean cuts (e.g., height "P3+10", z "-5").

    SCHEMA EXAMPLE:
    {{
      "Part_Name": "Custom_{part_name}",
      "Steps": [
        {{ "action": "BASE", "shape": "box", "length": "P1", "width": "P2", "height": "P3", "z": "0" }},
        {{ "action": "SUBTRACT", "shape": "cylinder", "radius": "P4/2", "height": "P3+10", "z": "-5" }}
      ]
    }}

    CRITICAL INSTRUCTION: Output ONLY the raw JSON object. Do not include markdown code blocks (```json). Do not include any conversational text.
    """

    try:
        # Initialize client (uses provided key, or environment variable if None)
        client = genai.Client(api_key=api_key) if api_key else genai.Client()
        
        response = client.models.generate_content(
            model=model_name,
            contents=prompt
        )
        
        json_str = response.text.strip()
        _log(f"[AI] Received CSG logic ({len(json_str)} chars)")

        # --- REPAIRED JSON CLEANING LOGIC ---
        if json_str.startswith("```json"): 
            json_str = json_str[7:-3].strip()
        elif json_str.startswith("```"): 
            json_str = json_str[3:-3].strip()
        # ------------------------------------

        recipe = json.loads(json_str)
        if "Steps" not in recipe: raise ValueError("Expected 'Steps' array in JSON")
        
        os.makedirs("templates", exist_ok=True)
        file_path = os.path.join("templates", f"Custom_{part_name}.json")
        
        with open(file_path, "w") as f:
            json.dump(recipe, f, indent=2)
            
        _log(f"[AI] ✔ Success! Generated recipe saved to: {file_path}")
        return True
        
    except Exception as e:
        _log(f"[AI] ❌ Failed to generate part: {e}")
        return False

# ==========================================
# Run Standalone (If you double-click this file directly)
# ==========================================
if __name__ == "__main__":
    print("=======================================")
    print(" SIRAAL AI - TEXT-TO-CAD GENERATOR")
    print("=======================================")
    p_name = input("Enter a short name for the part (e.g., Sensor_Mount): ")
    p_desc = input("Describe how to build it using P1, P2, P3, P4:\n> ")
    generate_siraal_shape(p_name, p_desc)