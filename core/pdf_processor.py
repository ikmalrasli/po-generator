import json
import os
from dotenv import load_dotenv
from google import genai
from google.genai import types
import pathlib
from datetime import datetime

from config.settings import GEMINI_MODEL, GEMINI_TEMPERATURE, GEMINI_TOP_P, JSONS_DIR

class PDFProcessor:
    def __init__(self):
        load_dotenv()
        self.client = genai.Client(api_key=os.getenv("GOOGLE_API_KEY"))
        self.model = GEMINI_MODEL

    def extract_po_data(self, filepath, gui_data):
        print("Sending request to Gemini...")

        prompt = """
        Extract the following information from the quotation provided below and return it as a valid JSON object. Do not include any text or formatting outside of the JSON object.

        JSON Structure:
        {
        "companyName": "The name of the company providing the quotation.",
        "address": "The full mailing address of the company.",
        "quotationNumber": "The unique quotation number or reference ID.",
        "pic": {
            "name": "The name of the Person-in-Charge or contact person. Use null if not found.",
            "email": "The contact person's email address. Use null if not found.",
            "phone": "The contact person's phone number. Use null if not found."
        },
        "terms": {
            "payment": "The payment terms (e.g., 'COD', '30 Days', '50% Upfront').",
            "deliveryWeeks": "The delivery lead time, converted to a number of weeks (e.g., '14 days' becomes 2, '4 weeks' becomes 4). If the delivery time is a range (e.g., '2-4 weeks', '6-8 weeks'), **always select the lower number** to represent the shortest possible lead time (e.g., '2-4 weeks' becomes 2, '6-8 weeks' becomes 6). Use null if not specified or cannot be converted."
        },
        "items": [
            {
            "quantity": "The numerical quantity of the item.",
            "unit": "The unit of measure (e.g., 'pcs', 'kgs', 'lot').",
            "description": "The full description of the item.",
            "unitPrice": "The price per unit as a number."
            }
        ]
        }
        """

        response = self.client.models.generate_content(
            model=self.model,
            contents=[
                types.Part.from_bytes(
                    data=filepath.read_bytes(),
                    mime_type='application/pdf',
                ),
                prompt
            ],
            config=types.GenerateContentConfig(
                temperature=GEMINI_TEMPERATURE,
                top_p=GEMINI_TOP_P,
                response_mime_type="application/json"
            )
        )

        try:
            cleaned_response = response.text.strip().replace('```json', '').replace('```', '')
            extracted_data = json.loads(cleaned_response)
            
            # Add GUI data to the extracted data
            extracted_data['gui_data'] = gui_data
            
            # Save extracted JSON for reference
            self._save_extracted_json(extracted_data)
            
            return extracted_data

        except json.JSONDecodeError as e:
            print(f"Error decoding JSON: {e}")
            print("Raw response from API:", response.text)
            raise

    def _save_extracted_json(self, extracted_data):
        """Save extracted JSON data for reference"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_path = JSONS_DIR / f"output_{timestamp}.json"
        with open(json_path, 'w') as f:
            json.dump(extracted_data, f, indent=2)