import os
import json
from extractor import extract_text_from_docx, extract_text_from_pdf, needs_permission

# Try to import config, but handle if it doesn't exist or is different in this project
try:
    from config import app_config
except ImportError:
    # Fallback or check environment directly
    class MockConfig:
        GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    app_config = MockConfig

def extract_from_file_ai(file_path, api_key=None):
    """
    Extracts figure captions and credit lines using the new google-genai SDK.
    """
    # Use provided key, or fallback to system key
    final_api_key = api_key or app_config.GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    
    if not final_api_key:
        raise ValueError("Gemini API Key is required. Please provide it or configure GEMINI_API_KEY in .env")

    try:
        from google import genai
        from google.genai import types
    except ImportError:
        raise RuntimeError("google-genai library not installed. Please install it using: pip install google-genai>=0.3.0")

    client = genai.Client(api_key=final_api_key)

    # Extract text content
    if file_path.lower().endswith('.pdf'):
        paragraphs = extract_text_from_pdf(file_path)
    else:
        paragraphs = extract_text_from_docx(file_path)
    
    full_text = "\n".join(paragraphs)
    
    prompt = f"""
    You are an expert editorial assistant. Your task is to extract figure, table, and box captions along with their credit lines from the provided manuscript text.

    STRICTLY FOLLOW THESE RULES:
    1. Identify all FIGURES, TABLES, BOXES, including UNNUMBERED FIGURES and UNNUMBERED BOXES.
    2. Extract the 'Item Type' (e.g., Figure, Table, Box), 'Item Number' (e.g., 1.1, 1-2), 'Caption' (Title/Legend), and 'Credit Line'.
    3. If an item (Figure, Table, or Box) does not have a number, set 'item_no' to "Unnumbered".
    4. The 'Credit Line' is usually at the end of the caption for figures, or at the bottom of tables/boxes.
    5. Look for credit keywords: Source, Reprint, Adapted from, Modified from, Copyright, Courtesy of, matches for '©', etc.
    6. 'Chapter' number usually appears at the beginning of the text (e.g., "Chapter 1"). Extract it if available.
    7. Return the output STRICTLY as a JSON list of objects.

    The JSON objects must have these keys:
    - "chapter": The chapter number or title.
    - "item_type": "Figure", "Table", "Box", etc.
    - "item_no": The number (e.g. "1.2") or "Unnumbered".
    - "caption": The text of the caption/title.
    - "credit": The extracted credit line. If none, leave empty string "".
    
    ONLY include items that have a detected credit line.
    
    MANUSCRIPT TEXT:
    {full_text}
    """

    try:
        response = client.models.generate_content(
            model='gemini-3-pro-preview',
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type='application/json'
            )
        )
        
        # Parse JSON response directly
        try:
            data = json.loads(response.text)
        except json.JSONDecodeError:
            # Fallback if the response text contains markdown code blocks
            text = response.text
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0]
            elif "```" in text:
                text = text.split("```")[1]
            data = json.loads(text)
        
        # Post-process to add 'needs_permission' calculation locally
        processed_results = []
        for item in data:
            # Ensure all keys exist
            caption = item.get("caption", "")
            credit = item.get("credit", "").strip()
            
            # Post-processing: Clean up common prefixes to match manual extractor behavior
            if credit.startswith('('):
                credit = credit[1:].strip()
            
            if credit.lower().startswith('used '):
                credit = credit[5:].strip()
            
            processed_results.append({
                "chapter": item.get("chapter", ""),
                "item_type": item.get("item_type", "Figure"),
                "item_no": item.get("item_no", ""),
                "caption": caption,
                "credit": credit,
                "needs_permission": needs_permission(caption, credit)
            })
            
        return processed_results

    except Exception as e:
        print(f"AI Extraction Error: {e}")
        # Fallback or re-raise? specific error message is better
        raise RuntimeError(f"Gemini API Error: {str(e)}")
