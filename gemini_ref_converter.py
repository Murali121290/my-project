import os
import logging
from typing import Optional, Dict, Any
import google.generativeai as genai

logger = logging.getLogger(__name__)


def convert_to_granular_style(raw_text: str, target_style: str) -> Optional[Dict[str, Any]]:
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        logger.error("GOOGLE_API_KEY not found")
        return None

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-2.0-flash")

        # 🔹 Define ALL bib fields here
        bib_fields = [
            "bib_surname", "bib_fname", "bib_title", "bib_journal", "bib_year",
            "bib_volume", "bib_issue", "bib_fpage", "bib_lpage", "bib_doi",
            "bib_url", "bib_book", "bib_chaptertitle", "bib_editionno",
            "bib_ed_surname", "bib_publisher", "bib_location",
            "bib_institution", "bib_school", "bib_conference",
            "bib_confacronym", "bib_conflocation", "bib_deg",
            "bib_reportnum"
        ]

        # 🔹 Create metadata schema dynamically
        metadata_properties = {
            field: {"type": ["string", "null"]}
            for field in bib_fields
        }

        response_schema = {
            "type": "object",
            "properties": {
                "formatted_output": {"type": "string"},
                "metadata": {
                    "type": "object",
                    "properties": metadata_properties,
                    "required": bib_fields
                }
            },
            "required": ["formatted_output", "metadata"]
        }

        prompt = f"""
Parse the following reference.

1. Format it in {target_style} style.
2. Extract metadata into the defined bib_ fields.

Return only structured output.

Reference:
{raw_text}
"""

        response = model.generate_content(
            prompt,
            generation_config={
                "response_mime_type": "application/json",
                "response_schema": response_schema,
                "temperature": 0.1
            }
        )

        if not response:
            return None

        return response.candidates[0].content.parts[0].text

    except Exception as e:
        logger.error(f"Structured mapping failed: {e}")
        return None
