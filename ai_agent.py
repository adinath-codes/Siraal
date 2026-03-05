import google.generativeai as genai
import os

class SiraalAI_Copilot:
    def __init__(self, api_key):
        # Configure the Gemini API
        genai.configure(api_key=api_key)
        # Using the flash model for maximum speed during the hackathon
        self.model = genai.GenerativeModel('gemini-2.5-flash')
        
    def generate_bom_from_prompt(self, user_prompt, output_filename="ai_generated_batch.csv"):
        """Translates natural language into a strict manufacturing CSV."""
        
        system_instructions = """
        You are an expert mechanical engineering AI.
        Your job is to translate the user's natural language request into a strict manufacturing CSV file.
        
        You MUST use exactly these columns in this exact order:
        Part_Number,Description,Material,Plate_Length,Plate_Width,Thickness,Hole_Diameter,Hole_Count,Revision,Tolerance
        
        Rules:
        1. If the user does not specify a Part_Number, generate one (e.g., AI-001, AI-002).
        2. If the user does not specify a Material, assume 'Steel-1020'.
        3. If the user does not specify a Revision or Tolerance, assume 'A' and '+/-0.1mm'.
        4. Hole_Count defaults to 4 unless specified.
        5. DO NOT output any markdown, conversational text, or explanations. 
        6. OUTPUT ONLY THE RAW CSV TEXT.
        """
        
        full_prompt = system_instructions + "\n\nUser Request: " + user_prompt
        
        try:
            response = self.model.generate_content(full_prompt)
            raw_text = response.text.strip()
            
            # Sanitize output (LLMs sometimes wrap CSVs in markdown blocks despite instructions)
            clean_csv = raw_text.replace("```csv", "").replace("```", "").strip()
            
            # Save the AI's response directly to a file
            with open(output_filename, "w", encoding="utf-8") as file:
                file.write(clean_csv)
                
            return True, output_filename
            
        except Exception as e:
            return False, str(e)