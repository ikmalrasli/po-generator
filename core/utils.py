import re
from num2words import num2words
from datetime import datetime

def format_address_for_excel(address, max_length=45):
    address = address.strip()

    if len(address) <= max_length:
        return address, ""

    best_split_index = -1
    for i in range(min(len(address), max_length), 0, -1):
        if address[i] == ',':
            best_split_index = i + 1
            break

    if best_split_index == -1:
        for i in range(min(len(address), max_length), 0, -1):
            if address[i].isspace():
                best_split_index = i
                break
    
    if best_split_index == -1:
        best_split_index = max_length
        
    line1 = address[:best_split_index].strip()
    line2 = address[best_split_index:].strip()
    
    if line1.endswith(',') and line2:
        line1 = line1[:-1]
    
    return line1, line2

def validate_po_number_format(po_number):
    po_pattern = r'^P-\d{6}-\d{3}M$'
    return re.match(po_pattern, po_number) is not None

def number_to_ringgit(amount):
    try:
        amount_str = f"{amount:.2f}"
    except ValueError:
        return "Invalid input amount."

    ringgit_part, cents_part = amount_str.split('.')
    ringgit_int = int(ringgit_part)
    cents_int = int(cents_part)

    if ringgit_int > 0:
        ringgit_words = num2words(ringgit_int, lang='en').replace(',', '').replace(' and', '')
        ringgit_words_formal = ringgit_words.title()
        output = f"{ringgit_words_formal} Ringgit"
    else:
        output = ""

    if cents_int > 0:
        cents_words = num2words(cents_int, lang='en').replace(',', '').replace(' and', '')
        cents_words_formal = cents_words.title()

        if ringgit_int > 0:
            output += f" and {cents_words_formal} Cents"
        else:
            output = f"{cents_words_formal} Cents"
            
    if not output:
        return "Zero Ringgit Only"

    output += " Only"
    return output

def extract_project_number(po_number):
    """Extract project number from PO number by removing the last part"""
    return re.sub(r'-\d{3}M$', '', po_number)