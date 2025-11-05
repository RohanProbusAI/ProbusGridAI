import os
from pathlib import Path
from PyPDF2 import PdfReader
import pandas as pd
from PIL import Image
import base64
from openai import OpenAI
from dotenv import load_dotenv
import json
import re
from typing import List, Dict, Optional, Tuple

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Embedded Formula Data (as provided)
FORMULA_DATA = [
    {"LOB": "TW", "SEGMENT": "1+5", "INSURER": "LIBERTY", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "INSURER": "LIBERTY", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "LIBERTY", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "LIBERTY", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "LIBERTY", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "LIBERTY", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "INSURER": "LIBERTY", "PO": "90% of Payin", "REMARKS": "All Fuel"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "LIBERTY", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "TATA", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "TATA", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "TATA", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "TATA", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "LIBERTY", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "LIBERTY", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "LIBERTY", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "LIBERTY", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "LIBERTY", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "LIBERTY", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "LIBERTY", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "LIBERTY", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "LIBERTY", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "LIBERTY", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "LIBERTY", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "LIBERTY", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "LIBERTY", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "INSURER": "LIBERTY", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "BUS", "SEGMENT": "STAFF BUS", "INSURER": "LIBERTY", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "LIBERTY", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "LIBERTY", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "LIBERTY", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "LIBERTY", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "MISD", "SEGMENT": "MISD", "INSURER": "LIBERTY", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "TRACTOR", "INSURER": "LIBERTY", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "MISC", "INSURER": "LIBERTY", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "Misc, Misd, Tractor", "INSURER": "LIBERTY", "PO": "88% of Payin", "REMARKS": "NIL"}
]

COMPANY_MAPPING = {
    'icici': 'ICICI', 'reliance': 'Reliance', 'chola': 'Chola', 'sompo': 'Sompo',
    'kotak': 'Kotak', 'magma': 'Magma', 'bajaj': 'Bajaj', 'digit': 'Digit', 
    'liberty': 'Liberty', 'future': 'Future', 'tata': 'Tata', 'iffco': 'IFFCO', 
    'royal': 'Royal', 'sbi': 'SBI', 'zuno': 'Zuno', 'hdfc': 'HDFC', 'shriram': 'Shriram'
}

# Known Indian cities and states for location detection
KNOWN_LOCATIONS = [
    'ahmedabad', 'surat', 'vadodara', 'rajkot', 'mumbai', 'pune', 'nagpur', 'thane',
    'delhi', 'bangalore', 'bengaluru', 'chennai', 'hyderabad', 'kolkata', 'jaipur',
    'lucknow', 'kanpur', 'patna', 'indore', 'bhopal', 'chandigarh', 'coimbatore',
    'kochi', 'guwahati', 'bhubaneswar', 'visakhapatnam', 'goa', 'mysore', 'nashik',
    'guj', 'gujarat', 'mh', 'maharashtra', 'apts', 'andhra pradesh', 'karnataka',
    'tamil nadu', 'tn', 'assam', 'tripura', 'meghalaya', 'mizoram', 'west bengal',
    'wb', 'up', 'uttar pradesh', 'mp', 'madhya pradesh', 'raj', 'rajasthan',
    'bihar', 'jharkhand', 'kerala', 'odisha', 'chhattisgarh', 'uttarakhand',
    'hubli', 'bhubaneshwar', 'coimbatore', 'madurai', 'tiruchirappalli', 'salem',
    'bhubaneswar'
]

# RTO code patterns
RTO_CODE_PATTERN = r'^[A-Z]{2}\d{1,2}$|^AS\d{1,2}$|^[A-Z]{2,4}\d{1,2}$'

def extract_from_filename(file_path):
    """Extract company name and location from filename intelligently"""
    filename = str(file_path).lower()
    base_name = Path(file_path).stem.lower()

    # Try known mapping first
    for key, value in COMPANY_MAPPING.items():
        if key in filename:
            return value, _extract_location(filename)

    # Intelligent fallback: Look for capitalized words (likely company name)
    words = re.findall(r'\b[a-zA-Z]{3,}\b', base_name)
    potential_companies = [w.title() for w in words if w.title() not in {'Insurance', 'General', 'Grid', 'Payout', 'Payin', 'Excel', 'Pdf', 'Img'}]
    if potential_companies:
        company = potential_companies[0]
        return company, _extract_location(filename)

    # Final fallback
    return None, _extract_location(filename)


def _extract_location(filename):
    """Helper to extract location from filename"""
    location_patterns = [
        r'\b(guj|gujarat)\b', r'\b(mh|maharashtra|mumbai)\b', 
        r'\b(apts)\b', 
        r'\b(dl|delhi)\b', r'\b(kar|karnataka|bangalore|bengaluru)\b',
        r'\b(tn|tamilnadu|tamil nadu|chennai)\b', r'\b(assam)\b',
        r'\b(tripura)\b', r'\b(meghalaya)\b', r'\b(mizoram)\b',
        r'\b(wb|west bengal)\b', r'\b(up|uttar pradesh)\b',
        r'\b(mp|madhya pradesh)\b', r'\b(raj|rajasthan)\b',
        r'\b(bihar)\b', r'\b(jharkhand)\b'
    ]
    for pattern in location_patterns:
        match = re.search(pattern, filename, re.IGNORECASE)
        if match:
            return match.group(0).upper()
    return None


def is_location_column(column_name: str) -> bool:
    """Detect if a column name represents a location"""
    col_lower = str(column_name).lower().strip()
    
    # Check if it matches RTO code pattern
    if re.match(RTO_CODE_PATTERN, str(column_name).strip(), re.IGNORECASE):
        return True
    
    # Check if it's a known location
    if any(loc in col_lower for loc in KNOWN_LOCATIONS):
        return True
    
    # Check common location-related terms
    location_terms = ['location', 'city', 'state', 'region', 'rto', 'zone', 'area', 'pradesh', 'geo segment new', 'geo segment old']
    if any(term in col_lower for term in location_terms):
        return True
    
    return False


def get_all_specific_insurers_for_lob_segment(lob: str, segment: str) -> List[str]:
    """Collect all specific insurer names (not All/Rest) for given LOB+Segment"""
    specific = set()
    for rule in FORMULA_DATA:
        if rule.get("LOB").upper() != lob.upper() or rule.get("SEGMENT").upper() != segment.upper():
            continue
        insurer_field = rule.get("INSURER")
        if not insurer_field:
            continue
        insurers = [i.strip().upper() for i in insurer_field.split(',')]
        for ins in insurers:
            if ins not in ["ALL COMPANIES", "REST OF COMPANIES"]:
                specific.add(ins)
    return list(specific)


def is_company_in_specific_list(company: str, specific_insurers: List[str]) -> bool:
    """Check if company matches any in the specific insurer list"""
    if not company:
        return False
    company_norm = company.upper()
    for ins in specific_insurers:
        if ins in company_norm or company_norm in ins:
            return True
        if any(word in company_norm for word in ins.split()):
            return True
    return False


def calculate_payout(normalized_data):
    company_name = normalized_data.get("COMPANY NAME", "Unknown").strip()
    segment = normalized_data.get("SEGMENT", "Unknown").upper().strip()
    original_segment = normalized_data.get("ORIGINAL_SEGMENT", segment)
    policy_type = (normalized_data.get("POLICY TYPE", "") or "").upper().strip()
    payin = normalized_data.get("PAYIN", 0)
    location = normalized_data.get("LOCATION", "N/A")

    if payin is None or payin == "":
        return "0.00%", "Payin missing → Payout 0%"

    payin_value = float(str(payin).replace('%', '').strip() or 0)

    if payin_value == 0:
        return "0.00%", "Payin is 0% → Payout 0% (no calculation applied)"

    payin_category = (
        "Payin Below 20%" if payin_value <= 20 else
        "Payin 21% to 30%" if 21 <= payin_value <= 30 else
        "Payin 31% to 50%" if 31 <= payin_value <= 50 else
        "Payin Above 50%"
    )

    # Determine LOB
    segment_norm = segment.replace(" ", "").upper()

    lob = "TW"
    if "BUS" in segment_norm or "SCHOOL" in segment_norm or "STAFF" in segment_norm:
        lob = "BUS"
    elif "PVT" in segment_norm or "CAR" in segment_norm:
        lob = "PVT CAR"
    elif "CV" in segment_norm or "GVW" in segment_norm or "PCV" in segment_norm or "GCV" in segment_norm or 'PCV Bus' in segment_norm or 'PCV School' in segment_norm:
        lob = "CV"
    elif "TAXI" in segment_norm:
        lob = "TAXI"
    elif "MISD" in segment_norm or "TRACTOR" in segment_norm or "MISC" in segment_norm:
        lob = "MISD"

    # Get policy type from segment
    segment_policy = None
    if "COMP" in segment or "SAOD" in segment or "TP" in segment:
        segment_policy = next((pt for pt in ['COMP', 'SAOD', 'TP'] if pt in segment), None)
    else:
        segment_policy = policy_type

    # Precompute specific insurers for this LOB+Segment
    specific_insurers = get_all_specific_insurers_for_lob_segment(lob, segment)

    candidate_rules = []
    company_upper = company_name.upper()

    for rule in FORMULA_DATA:
        if rule["LOB"].upper() != lob.upper() or rule["SEGMENT"].upper() != segment.upper():
            continue

        insurers = [i.strip().upper() for i in rule["INSURER"].split(',')]
        insurer_type = None
        specificity = 0

        # Priority 1: Specific Company Match
        for ins in insurers:
            if ins in ["ALL COMPANIES", "REST OF COMPANIES"]:
                continue
            if (ins in company_upper or company_upper in ins or
                    any(word in company_upper for word in ins.split())):
                insurer_type = "SPECIFIC"
                specificity = 3
                break

        # Priority 2: Rest of Companies
        if not insurer_type and "REST OF COMPANIES" in insurers:
            if not is_company_in_specific_list(company_name, specific_insurers):
                insurer_type = "REST"
                specificity = 2

        # Priority 3: All Companies
        if not insurer_type and "ALL COMPANIES" in insurers:
            insurer_type = "ALL"
            specificity = 1

        if not insurer_type:
            continue

        remarks = rule.get("REMARKS", "").upper().strip()
        remarks_score = 0

        if remarks == "NIL":
            remarks_score = 10
        elif remarks == "ALL FUEL":
            remarks_score = 5
        elif payin_value > 0 and remarks == payin_category.upper():
            remarks_score = 8
        elif payin_value > 20 and "PAYIN ABOVE 20%" in remarks:
            remarks_score = 6
        elif "ALL COMPANIES" in insurers and remarks in ["ALL FUEL", "NIL", ""]:
            remarks_score = 5
        elif "ZUNO" in insurers and "ZUNO - 21" in remarks:
            remarks_score = 7
        else:
            continue

        total_score = (specificity * 100) + remarks_score
        candidate_rules.append((rule, total_score, specificity, insurer_type))

    # Select best rule
    if candidate_rules:
        candidate_rules.sort(key=lambda x: x[1], reverse=True)
        matched_rule = candidate_rules[0][0]
        explanation = f"Match: LOB={lob}, Segment={segment}, Insurer={matched_rule['INSURER']}, Remarks={matched_rule.get('REMARKS', 'NIL')}, Formula={matched_rule['PO']}"
    else:
        matched_rule = {"PO": "90% of Payin", "INSURER": "Default", "REMARKS": "NIL"}
        explanation = f"No Specific Match: Default Formula=90% of Payin"

    if original_segment.upper() != segment.upper():
        explanation += f"; Input Segment: {original_segment} → Matched: {segment}"

    # Apply formula
    po = matched_rule["PO"]
    payout = payin_value
    op_str = ""

    if "flat" in po.lower():
        flat_val = float(po.split('%')[0].strip())
        payout = flat_val
        op_str = f"= {payout:.2f}% (flat)"
    elif "% of Payin" in po:
        if "90" in po:
            factor = 0.9
        elif "88" in po:
            factor = 0.88
        elif "Less 2" in po:
            factor = 0.98
        else:
            factor = 1.0
        payout = payin_value * factor
        op_str = f"* {factor} = {payout:.2f}%"
    elif po.startswith("-"):
        deduction = float(po.replace('%', ''))
        payout = payin_value + deduction
        op_str = f"{po} = {payout:.2f}%"

    payout = max(0, payout)
    explanation += f"; Calculated: {payin_value:.2f}% {op_str}"

    return f"{payout:.2f}%", explanation


def extract_percentage(value) -> Optional[float]:
    if pd.isna(value) or value is None:
        return None

    value_str = str(value).strip().replace('%', '')

    try:
        num = float(value_str)
        if 0 <= num <= 1:        # 0.18 → 18%
            return round(num * 100, 2)
        elif 0 <= num <= 100:    # 18 → 18%
            return round(num, 2)
    except:
        pass

    match = re.search(r'(\d+(?:\.\d+)?)', value_str)
    if match:
        num = float(match.group(1))
        if 0 <= num <= 100:
            return round(num, 2)

    return None

def normalize_segment(segment_str):
    """Normalize segment names to standard format"""
    
    if not segment_str or pd.isna(segment_str):
        return "Unknown"
    
    segment_lower = re.sub(r'\s+', ' ', str(segment_str).lower().strip())

    # Don't normalize if it's a location
    if segment_lower == "apts":
        return segment_str
    
    if segment_lower == "Misc":
        return "MISD"

    # Bus
    if 'school buses' in segment_lower or 'school bus' in segment_lower or 'school buses' in segment_lower:
        return "SCHOOL BUS"
    if 'staff bus' in segment_lower:
        return "STAFF BUS"
    if 'bus' in segment_lower and 'school' not in segment_lower and 'staff' not in segment_lower:
        return "SCHOOL BUS"
    
    # Taxi
    if 'taxi' in segment_lower:
        return "TAXI"
    
    # 3 Wheeler
    if any(x in segment_lower for x in ['3w', '3 w', 'three wheeler', '3wheeler']):
        if 'pcv' in segment_lower or 'passenger' in segment_lower:
            return "PCV 3W"
        elif 'gcv' in segment_lower or 'goods' in segment_lower:
            return "GCV 3W"
        return "PCV 3W"
    
    # Two Wheeler
    if any(x in segment_lower for x in ['tw', '2w', 'two wheeler', 'bike', 'scooter', 'ev', 'motor cycle', 'tw new']):
        if 'saod' in segment_lower or 'comp' in segment_lower:
            return "TW SAOD + COMP"
        
        elif 'tp' in segment_lower or 'satp' in segment_lower:
            return "TW TP"
        elif '1+5' in segment_lower or 'grid' in segment_lower or 'new' in segment_lower or 'tw new' in segment_lower:
            return "1+5"
        return "TW"
    
    # Private Car
    if any(x in segment_lower for x in ['pvt car', 'private car', 'car']):
        # Priority 1: Check for explicit TP mentions (but NOT if COMP is also present)
        if (('pvt car tp' in segment_lower or 
            'private car tp' in segment_lower or 
            'car tp' in segment_lower or
            'satp' in segment_lower) and 
            'comp' not in segment_lower and 
            'saod' not in segment_lower):
            return "PVT CAR TP"
        
        # Priority 2: Check for COMP/SAOD/Package
        elif ('comp' in segment_lower or 
            'saod' in segment_lower or 
            'package' in segment_lower or
            'pvt car comp' in segment_lower):
            return "PVT CAR COMP + SAOD"
        
        # Priority 3: If only "TP" appears (without COMP/SAOD)
        elif 'tp' in segment_lower:
            return "PVT CAR TP"
        
        # Default: Assume comprehensive
        return "PVT CAR COMP + SAOD"

    # Commercial Vehicle
    if any(x in segment_lower for x in ['cv', 'gcv', 'pcv', 'lcv', 'gvw', 'tn', 'tonnage', 'pcv bus', 'pcv school', 'pcv 4w']):
        if 'upto 2.5' in segment_lower or '0-2.5' in segment_lower or '2.5 gvw' in segment_lower:
            return "Upto 2.5 GVW"
        elif 'pcv 3w' in segment_lower or 'gcv 3w' in segment_lower or 'pcv 4w non school' in segment_lower or 'pcv bus non school' in segment_lower or 'pcv bus school':
            return "All GVW & PCV 3W, GCV 3W"
        elif '2.5-3.5' in segment_lower:
            return "2.5-3.5 GVW"
        elif '3.5-12' in segment_lower:
            return "3.5-12 GVW"
        elif '12-45' in segment_lower:
            if 'satp' in segment_lower or 'tp' in segment_lower:
                return "12-45 GVW SATP"
            return "12-45 GVW"
        elif '>45' in segment_lower or '45t' in segment_lower:
            return ">45 GVW"
        return "All GVW & PCV 3W, GCV 3W"
    
    # Misc
    if any(x in segment_lower for x in ['misc', 'misd', 'tractor']):
        return "MISD"
    
    return segment_str


def normalize_policy_type(policy_str):
    """Normalize policy type - Package = Comprehensive"""
    if not policy_str or pd.isna(policy_str):
        return None
    
    policy_lower = str(policy_str).lower().strip()
    
    if 'package' in policy_lower:
        return "Comprehensive"
    
    if 'comp' in policy_lower and 'saod' in policy_lower:
        return "COMP + SAOD"
    elif 'comp' in policy_lower:
        return "Comprehensive"
    elif 'saod' in policy_lower or 'od' in policy_lower:
        return "SAOD"
    elif 'tp' in policy_lower or 'satp' in policy_lower:
        return "TP"
    elif 'aotp' in policy_lower:
        return "AOTP"
    
    return policy_str.upper()


def detect_table_structure(sheet_data) -> Dict:
    """
    Intelligently detect the structure of the Excel table.
    Returns: {
        'header_row_index': int,
        'data_start_index': int,
        'segment_col': str,
        'policy_col': str or None,
        'fuel_col': str or None,
        'vehicle_age_col': str or None,
        'location_cols': list of str (column keys that represent locations)
    }
    """
    header_row_index = -1
    header_map = {}
    location_cols = []
    
    # Step 1: Find header row
    for i, row in enumerate(sheet_data):
        row_vals_lower = {k: str(v).lower().strip() for k, v in row.items() if v}
        
        # Look for typical header keywords
        header_keywords = ['segment', 'product', 'policy', 'fuel', 'vehicle age', 'age', 
                          'location', 'region', 'rto', 'zone']
        
        if any(any(keyword in val for keyword in header_keywords) for val in row_vals_lower.values()):
            header_row_index = i
            
            # Map important columns
            for col_key, val in row.items():
                if not val:
                    continue
                val_lower = str(val).lower().strip()
                
                if 'segment' in val_lower or 'product line' in val_lower:
                    header_map['segment'] = col_key
                elif 'policy' in val_lower:
                    header_map['policy'] = col_key
                elif 'fuel' in val_lower:
                    header_map['fuel'] = col_key
                elif 'vehicle age' in val_lower or 'age' in val_lower:
                    header_map['vehicle_age'] = col_key
                elif is_location_column(val):
                    location_cols.append(col_key)
            
            break
    
    # Step 2: If no header found, make educated guess
    if header_row_index == -1 and sheet_data:
        # Check if first row looks like headers
        first_row = sheet_data[0]
        first_row_vals = list(first_row.values())
        
        # If first few cells are text (not numbers), likely headers
        text_cells = sum(1 for v in first_row_vals[:5] if v and not str(v).replace('.','').isdigit())
        if text_cells >= 2:
            header_row_index = 0
            
            # Auto-detect columns by position and content
            for col_key, val in first_row.items():
                if not val:
                    continue
                val_str = str(val)
                
                # Check if this column is a location
                if is_location_column(val_str):
                    location_cols.append(col_key)
                elif 'segment' in val_str.lower():
                    header_map['segment'] = col_key
                elif 'policy' in val_str.lower():
                    header_map['policy'] = col_key
                elif 'fuel' in val_str.lower():
                    header_map['fuel'] = col_key
    
    # Step 3: If still no segment column, guess by position
    if 'segment' not in header_map and sheet_data:
        header_map['segment'] = 'col_0'  # First column usually segment
    
    # Step 4: Detect location columns if not found in headers
    if not location_cols and header_row_index >= 0:
        # Check next row after header for location patterns
        for col_key in sheet_data[0].keys():
            if col_key in header_map.values():
                continue  # Skip known non-location columns
            
            # Sample first few data rows
            sample_values = []
            for i in range(header_row_index + 1, min(header_row_index + 5, len(sheet_data))):
                val = sheet_data[i].get(col_key)
                if val and str(val).strip():
                    sample_values.append(str(val))
            
            # If values look like locations or numbers, it's likely a location column
            if sample_values:
                # Check if it's numeric (payin values) or location names
                is_numeric = any(str(v).replace('.','').isdigit() for v in sample_values)
                is_location = any(is_location_column(v) for v in sample_values)
                
                if is_numeric or is_location:
                    location_cols.append(col_key)
    
    data_start_index = header_row_index + 1 if header_row_index >= 0 else 0
    
    return {
        'header_row_index': header_row_index,
        'data_start_index': data_start_index,
        'segment_col': header_map.get('segment'),
        'policy_col': header_map.get('policy'),
        'fuel_col': header_map.get('fuel'),
        'vehicle_age_col': header_map.get('vehicle_age'),
        'location_cols': location_cols
    }


def user_confirm_structure(structure_info, sheet_name):
    """Interact with user to confirm/edit detected structure"""
    print(f"\n🔍 Detected Structure for Sheet '{sheet_name}':")
    
    # Segment
    detected_segment = structure_info.get('segment_col', 'None')
    print(f"   Segment Column: {detected_segment}")
    confirm = input("   Is this correct? (y/n): ").strip().lower()
    if confirm != 'y':
        new_segment = input("   Enter correct segment column key (e.g., col_0): ").strip()
        if new_segment:
            structure_info['segment_col'] = new_segment
    
    # Policy
    detected_policy = structure_info.get('policy_col', 'None')
    print(f"   Policy Column: {detected_policy}")
    confirm = input("   Is this correct? (y/n): ").strip().lower()
    if confirm != 'y':
        new_policy = input("   Enter correct policy column key (or leave blank if none): ").strip()
        structure_info['policy_col'] = new_policy if new_policy else None
    
    # Location
    detected_locations = ', '.join(structure_info.get('location_cols', [])) or 'None'
    print(f"   Location Columns: {detected_locations}")
    confirm = input("   Is this correct? (y/n): ").strip().lower()
    if confirm != 'y':
        new_locations = input("   Enter correct location column keys (comma-separated, e.g., col_2,col_3): ").strip()
        if new_locations:
            structure_info['location_cols'] = [col.strip() for col in new_locations.split(',')]
    
    return structure_info


def restructure_pivoted_data(sheet_data, structure_info, sheet_name):
    """
    Restructure pivoted table data where locations are columns.
    Converts wide format to long format with one entry per location.
    """
    segment_col = structure_info['segment_col']
    policy_col = structure_info['policy_col']
    fuel_col = structure_info['fuel_col']
    vehicle_age_col = structure_info['vehicle_age_col']
    location_cols = structure_info['location_cols']
    data_start_index = structure_info['data_start_index']
    
    restructured_data = []
    
    # Get location names from header row if exists
    location_names = {}
    if structure_info['header_row_index'] >= 0:
        header_row = sheet_data[structure_info['header_row_index']]
        for col_key in location_cols:
            location_names[col_key] = str(header_row.get(col_key, col_key))
    else:
        # Use column keys as location names
        for col_key in location_cols:
            location_names[col_key] = col_key
    
    # Process each data row
    last_segment = None
    last_policy = None
    last_fuel = None
    last_vehicle_age = None
    
    for i in range(data_start_index, len(sheet_data)):
        row = sheet_data[i]
        
        # Skip completely empty rows
        if all(v is None or str(v).strip() == "" for v in row.values()):
            continue
        
        # Get segment (with fill-down logic)
        current_segment = row.get(segment_col) if segment_col else None
        if current_segment and str(current_segment).strip():
            last_segment = str(current_segment).strip()
        elif last_segment:
            current_segment = last_segment
        else:
            continue  # No segment, skip row
        
        # Get other attributes (with fill-down)
        current_policy = None
        if policy_col:
            current_policy = row.get(policy_col)
            if current_policy and str(current_policy).strip():
                last_policy = str(current_policy).strip()
            elif last_policy:
                current_policy = last_policy
        
        current_fuel = None
        if fuel_col:
            current_fuel = row.get(fuel_col)
            if current_fuel and str(current_fuel).strip():
                last_fuel = str(current_fuel).strip()
            elif last_fuel:
                current_fuel = last_fuel
        
        current_vehicle_age = None
        if vehicle_age_col:
            current_vehicle_age = row.get(vehicle_age_col)
            if current_vehicle_age and str(current_vehicle_age).strip():
                last_vehicle_age = str(current_vehicle_age).strip()
            elif last_vehicle_age:
                current_vehicle_age = last_vehicle_age
        
        # Create an entry for each location column
        for col_key in location_cols:
            payin_value = row.get(col_key)
            
            # Skip if no payin value
            if payin_value is None or str(payin_value).strip() == "":
                continue
            
            # Create structured entry
            entry = {
                'segment': current_segment,
                'policy_type': current_policy,
                'location': location_names.get(col_key, col_key),
                'payin': str(payin_value).strip(),
                'other_info': []
            }
            
            # Add fuel and vehicle age to other_info
            if current_fuel:
                entry['other_info'].append(current_fuel)
            if current_vehicle_age:
                entry['other_info'].append(current_vehicle_age)
            
            # Join other_info
            entry['other_info'] = ', '.join(entry['other_info']) if entry['other_info'] else None
            
            restructured_data.append(entry)
    
    return restructured_data


def extract_text_from_sheet(sheet_data, sheet_name):
    """Converts a list of row-dictionaries into text for the AI"""
    text = f"Sheet Name: {sheet_name}\n"
    for idx, row in enumerate(sheet_data, 1):
        row_values = [str(v) for v in row.values() if v]
        row_str = ", ".join(row_values)
        text += f"row{idx}: {row_str}\n"
    return text


def excel_to_json_converter(excel_path, output_json_path=None):
    """Convert Excel to JSON with SMART header detection"""
    print("\n" + "="*60)
    print("EXCEL TO JSON CONVERTER (SMART HEADER DETECTION)")
    print("="*60)
    
    excel_file = pd.ExcelFile(excel_path)
    sheet_names = excel_file.sheet_names
    
    print(f"\nFound {len(sheet_names)} worksheet(s):")
    for i, sheet in enumerate(sheet_names, 1):
        print(f"   {i}. {sheet}")
    
    all_sheets_data = {}
    
    for sheet_name in sheet_names:
        print(f"\nProcessing sheet: '{sheet_name}'")
        
        # Read without header
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        print(f"   Raw shape: {df.shape}")
        
        # SMART: Find header row (first row with text in first 5 cols)
        header_row_idx = 0
        for idx, row in df.iterrows():
            text_count = sum(1 for val in row[:min(5, len(row))] if pd.notna(val) and isinstance(val, str) and val.strip())
            num_count = sum(1 for val in row[:min(5, len(row))] if pd.notna(val) and not isinstance(val, str))
            if text_count >= 2 and num_count <= 1:  # Likely header
                header_row_idx = idx
                break
        
        print(f"   Detected header row: {header_row_idx}")
        
        # Re-read with correct header
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row_idx)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        # Rename columns to col_0, col_1... but keep original names in first row of JSON
        original_headers = df.columns.tolist()
        df.columns = [f"col_{i}" for i in range(len(df.columns))]
        
        # Insert original header names as first data row
        header_row = {f"col_{i}": str(name) for i, name in enumerate(original_headers)}
        json_data = [header_row] + df.to_dict('records')
        
        all_sheets_data[sheet_name] = json_data
        print(f"   Final records: {len(json_data)} (incl. header row)")
    
    if output_json_path is None:
        base_name = Path(excel_path).stem
        output_json_path = Path(excel_path).parent / f"{base_name}_converted.json"
    
    output_data = {
        "source_file": str(excel_path),
        "total_sheets": len(sheet_names),
        "sheets": all_sheets_data
    }
    
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)
    
    print(f"\nJSON file saved: {output_json_path}")
    return output_data


def process_single_sheet(sheet_name, sheet_data, source_file, company_name):
    """Process a single sheet with user interaction"""
    print(f"\n🔄 Processing sheet: '{sheet_name}'")
    print(f"   Records in sheet: {len(sheet_data)}")
    
    # Detect table structure
    structure_info = detect_table_structure(sheet_data)
    
    print(f"   Initial Structure detected:")
    print(f"     - Header row: {structure_info['header_row_index']}")
    print(f"     - Segment column: {structure_info['segment_col']}")
    print(f"     - Location columns: {structure_info['location_cols']}")
    
    # User confirmation and editing
    structure_info = user_confirm_structure(structure_info, sheet_name)
    
    # Check if this is a pivoted table (locations as columns)
    if structure_info['location_cols']:
        print(f"   📊 Pivoted table detected! Restructuring data...")
        extracted_list = restructure_pivoted_data(sheet_data, structure_info, sheet_name)
        print(f"   ✓ Restructured into {len(extracted_list)} entries")
    else:
        # Standard format - use AI extraction
        print(f"   📝 Standard format. Using AI extraction...")
        
        # Apply fill-down logic for segment column
        processed_sheet_data = []
        last_segment = None
        segment_col_key = structure_info.get('segment_col')
        data_start_index = structure_info['data_start_index']
        
        if segment_col_key:
            for i in range(data_start_index, len(sheet_data)):
                row = sheet_data[i]
                if all(v is None or str(v).strip() == "" for v in row.values()):
                    continue
                
                current_segment = row.get(segment_col_key)
                if current_segment is None and str(current_segment).strip():
                    last_segment = current_segment
                elif last_segment is not None:
                    row[segment_col_key] = last_segment
                
                processed_sheet_data.append(row)
        else:
            processed_sheet_data = sheet_data[data_start_index:]
        
        sheet_text = extract_text_from_sheet(processed_sheet_data, sheet_name)
        extracted_list = extract_structured_data_single(sheet_text, "Excel")
    
    # Normalize and calculate payout for each entry
    all_entries = []
    filename_company, filename_location = extract_from_filename(source_file)
    for extracted in extracted_list:
        normalized = normalize_extracted_data(extracted, source_file)
        normalized["COMPANY NAME"] = company_name
        if filename_location and not normalized["LOCATION"]:
            normalized["LOCATION"] = filename_location
        
        # Handle geo segment new/old in remarks
        if 'geo segment new' in sheet_name.lower() or any('geo segment new' in str(v).lower() for row in sheet_data for v in row.values()):
            if normalized.get("REMARK"):
                normalized["REMARK"].append("Geo Segment New Present")
            else:
                normalized["REMARK"] = ["Geo Segment New Present"]
        
        payout, explanation = calculate_payout(normalized)
        normalized["PAYOUT"] = payout
        normalized["CALCULATION EXPLANATION"] = explanation
        all_entries.append(normalized)
    
    print(f"   ✓ Extracted {len(extracted_list)} entries from sheet '{sheet_name}'")
    
    return all_entries


def parse_converted_json(json_path_or_data, original_excel_path=None):
    print("\n" + "="*60)
    print("PARSING CONVERTED JSON (SMART HEADER + MERGED CELLS + SHEET SELECTION)")
    print("="*60)

    # Load JSON
    if isinstance(json_path_or_data, str):
        json_path = json_path_or_data
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        data = json_path_or_data
        json_path = None

    source_file = original_excel_path or data.get("source_file") or json_path
    print(f"\nSource: {source_file}")

    company_name = input("\nEnter Company Name: ").strip()
    print(f"Company name set: {company_name}")

    # Define norm here
    def norm(s):
        return re.sub(r'\s+', ' ', str(s).strip()).lower() if s is not None else ""

    # Define the function inside
    def process_single_sheet_with_user_input(sheet_name, sheet_data, source_file, company_name, norm):
        """Process one sheet with full user input (used by parse_converted_json)"""
        print(f"   Records: {len(sheet_data)}")
        if not sheet_data:
            print("   Skipping empty sheet.")
            return []

        first_row = sheet_data[0]
        data_rows = sheet_data[1:]
        header_keys = sorted(first_row.keys())

        print("\n   Detected Headers:")
        for k in header_keys:
            print(f"      {k} → '{first_row.get(k, '')}'")

        # 1. SEGMENT
   
        manual_segment_raw = input("\n   Enter Segment (e.g. CV, TW, PVT CAR TP): ").strip().upper()
        manual_segment = normalize_segment(manual_segment_raw)
        print(f"   Segment entered: {manual_segment_raw}")
        if manual_segment != manual_segment_raw:
            print(f"   Normalized to: {manual_segment}")
        else:
            print(f"   Segment: {manual_segment}")

        # 2. PAY-IN COLUMNS
        while True:
            payin_input = input(
                "   Enter **exact header text** or **col_X** for PAY-IN column(s) – comma-separated: "
            ).strip()
            if not payin_input:
                print("   No pay-in columns? Skipping sheet.")
                return []

            wanted = [p.strip() for p in payin_input.split(',') if p.strip()]
            payin_cols = []
            payin_headers = []
            missing = []
            for want in wanted:
                found = False
                if want.startswith('col_') and want in header_keys:
                    payin_cols.append(want)
                    payin_headers.append(str(first_row[want]).strip())
                    found = True
                else:
                    for k, v in first_row.items():
                        if norm(v) == norm(want):
                            payin_cols.append(k)
                            payin_headers.append(str(v).strip())
                            found = True
                            break
                if not found:
                    missing.append(want)

            if not missing:
                print(f"   Pay-in columns: {', '.join(payin_cols)}")
                break
            else:
                print(f"   Not found: {', '.join(missing)}. Try again.")

        # 3. LOCATION COLUMN
        while True:
            loc_input = input(
                "   Enter **exact header text** or **col_X** for LOCATION column (or leave blank): "
            ).strip()
            if not loc_input:
                location_col = None
                print("   No location column – will use filename or N/A.")
                break

            found = False
            if loc_input.startswith('col_') and loc_input in header_keys:
                location_col = loc_input
                found = True
            else:
                for k, v in first_row.items():
                    if norm(v) == norm(loc_input):
                        location_col = k
                        found = True
                        break

            if found:
                print(f"   Location column → {location_col} ('{first_row[location_col]}')")
                break
            else:
                print(f"   No match for '{loc_input}'. Try again.")

        # 4. AUTO POLICY FROM HEADER
        def extract_policy_from_header(header):
            h = header.upper()
            if 'TP' in h or 'SATP' in h:
                return "TP"
            elif 'COMP' in h and 'SAOD' in h:
                return "COMP + SAOD"
            elif 'COMP' in h:
                return "COMP + SAOD"
            elif 'SAOD' in h or 'OD' in h:
                return "SAOD"
            else:
                return "COMP + SAOD"

        payin_policy_map = {col: extract_policy_from_header(first_row[col]) for col in payin_cols}
        print("   Auto-detected Policy Types:")
        for col, pol in payin_policy_map.items():
            print(f"      {first_row[col]} → {pol}")

        # 5. EXTRACT WITH FILL-DOWN
        entries = []
        last_segment = manual_segment
        last_location = None

        for row_idx, row in enumerate(data_rows, start=1):
            if all(v is None or str(v).strip() == "" for v in row.values()):
                continue

            current_segment = row.get(location_col) if location_col else None
            if current_segment and str(current_segment).strip():
                last_segment = str(current_segment).strip()

            current_location = None
            if location_col:
                val = row.get(location_col)
                if val and str(val).strip() and str(val).strip() not in {'-', ''}:
                    current_location = str(val).strip()
                    last_location = current_location
                else:
                    current_location = last_location
            loc_value = current_location or "N/A"

            for payin_col in payin_cols:
                raw = row.get(payin_col)
                if raw is None or str(raw).strip() in {'', '-'}:
                    continue
                payin = extract_percentage(raw)
                if payin is None:
                    continue

                entries.append({
                    "segment": last_segment,
                    "policy_type": payin_policy_map[payin_col],
                    "location": loc_value,
                    "payin": payin
                })

        print(f"   Extracted → {len(entries)} entries")

        # 6. NORMALIZE + PAYOUT
        sheet_entries = []
        _, filename_location = extract_from_filename(source_file)

        for item in entries:
            loc = item["location"] if item["location"] != "N/A" else (filename_location or "N/A")
            normalized = {
                "COMPANY NAME": company_name,
                "ORIGINAL_SEGMENT": manual_segment_raw,  # ← Use the raw input
                "SEGMENT": manual_segment,  # ← Use the already normalized version
                "POLICY TYPE": normalize_policy_type(item["policy_type"]),
                "LOCATION": loc,
                "PAYIN": item["payin"],
                "REMARK": ["Geo Segment New Present"] if 'geo segment new' in sheet_name.lower() else [],
                "CALCULATION EXPLANATION": ""
            }
            payout, explanation = calculate_payout(normalized)
            normalized["PAYOUT"] = payout
            normalized["CALCULATION EXPLANATION"] = explanation
            sheet_entries.append(normalized)

        print(f"   Final entries: {len(sheet_entries)}")

        # 7. SAVE INDIVIDUAL SHEET OUTPUT
        if sheet_entries:
            out_dict = {k: [] for k in [
                "COMPANY NAME","ORIGINAL_SEGMENT","SEGMENT","POLICY TYPE",
                "LOCATION","PAYIN","PAYOUT","REMARK","CALCULATION EXPLANATION"
            ]}
            for e in sheet_entries:
                for k in out_dict:
                    out_dict[k].append("; ".join(e.get(k, [])) if k == "REMARK" else e.get(k, ""))
            
            df = pd.DataFrame(out_dict)
            
            # Create safe filename
            safe_name = re.sub(r'[<>:"/\\|?*]', '_', sheet_name)
            safe_company = re.sub(r'[<>:"/\\|?*]', '_', company_name)
            out_file = f"grid_output_{safe_company}_{safe_name}.xlsx"
            
            df.to_excel(out_file, index=False)
            print(f"\n   📊 Individual sheet saved → {out_file}")
            print(f"   ✓ Contains {len(sheet_entries)} entries")
        else:
            print(f"\n   ⚠️  No entries to save for sheet '{sheet_name}'")

        return sheet_entries

    # === MAIN LOGIC ===

    all_sheet_outputs = []
    sheets = list(data["sheets"].items())
    total_sheets = len(sheets)

    print(f"\n📋 Workbook has {total_sheets} sheet(s):")
    for idx, (name, _) in enumerate(sheets, 1):
        print(f"   [{idx}] {name}")

    processed_sheets = set()  # Track processed sheets

    print(f"\n💡 Tip: You can process sheets one by one. After each sheet, you'll be asked if you want to continue.")
    while True:
        # Show available unprocessed sheets
        unprocessed = [(idx, name) for idx, (name, _) in enumerate(sheets, 1) if idx not in processed_sheets]
        
        if not unprocessed:
            print("\n✅ All sheets have been processed!")
            break
        
        print(f"\n📋 Available sheets to process:")
        for idx, name in unprocessed:
            print(f"   [{idx}] {name}")
        
        # Ask which sheet to process
        sheet_choice = input(f"\nEnter sheet number to process (1-{total_sheets}): ").strip()
        
        if not sheet_choice.isdigit():
            print("❌ Invalid input. Please enter a number.")
            continue
        
        sheet_idx = int(sheet_choice) - 1
        
        if sheet_idx < 0 or sheet_idx >= total_sheets:
            print(f"❌ Invalid sheet number. Choose between 1-{total_sheets}")
            continue
        
        if (sheet_idx + 1) in processed_sheets:
            print(f"⚠️  Sheet [{sheet_idx + 1}] already processed. Choose another.")
            continue
        
        # Process selected sheet
        sheet_name, sheet_data = sheets[sheet_idx]
        print(f"\n{'='*20} PROCESSING SHEET [{sheet_idx + 1}] '{sheet_name}' {'='*20}")
        entries = process_single_sheet_with_user_input(sheet_name, sheet_data, source_file, company_name, norm)
        all_sheet_outputs.extend(entries)
        processed_sheets.add(sheet_idx + 1)
        
        print(f"\n✅ Sheet '{sheet_name}' processed successfully!")
        print(f"   Extracted {len(entries)} entries")
        
        # Ask if user wants to process another sheet
        if len(processed_sheets) < total_sheets:
            continue_choice = input(f"\n🔄 Process another sheet? (y/n): ").strip().lower()
            if continue_choice != 'y':
                print(f"\n✅ Stopping. Processed {len(processed_sheets)} out of {total_sheets} sheets.")
                break
        else:
            print("\n✅ All sheets processed!")
            break

    print(f"\n{'='*60}")
    print(f"✅ PROCESSING COMPLETE")
    print(f"{'='*60}")
    print(f"Total entries extracted: {len(all_sheet_outputs)}")
    print(f"Sheets processed: {len(processed_sheets)}/{total_sheets}")




def normalize_extracted_data(extracted, file_path):
    """Normalize extracted data from AI"""
    normalized = {
        "COMPANY NAME": None,
        "ORIGINAL_SEGMENT": None,
        "SEGMENT": None,
        "POLICY TYPE": None,
        "LOCATION": None,
        "PAYIN": None,
        "PAYOUT": "",
        "REMARK": [],
        "CALCULATION EXPLANATION": ""
    }
    
    filename_company, filename_location = extract_from_filename(file_path)
    normalized["COMPANY NAME"] = extracted.get("company_name") or filename_company or "Unknown"
    
    normalized["ORIGINAL_SEGMENT"] = extracted.get("segment")
    
    if normalized["ORIGINAL_SEGMENT"]:
        normalized["SEGMENT"] = normalize_segment(normalized["ORIGINAL_SEGMENT"])
    else:
        normalized["SEGMENT"] = "Unknown"
    
    if extracted.get("policy_type"):
        normalized["POLICY TYPE"] = normalize_policy_type(extracted["policy_type"])
    else:
        normalized["POLICY TYPE"] = "Comprehensive"
    
    # Location priority: extracted location > filename location
    if extracted.get("location"):
        normalized["LOCATION"] = extracted["location"]
    elif filename_location:
        normalized["LOCATION"] = filename_location
    
    if extracted.get("payin"):
        payin_str = str(extracted["payin"])
        payin_match = re.search(r'(\d+(?:\.\d+)?)', payin_str)
        if payin_match:
            normalized["PAYIN"] = float(payin_match.group(1))
    
    if extracted.get("other_info"):
        other_info_str = str(extracted["other_info"]).strip()
        if other_info_str:
            normalized["REMARK"] = [x.strip() for x in other_info_str.split(',') if x.strip()]
    
    return normalized


def extract_structured_data_single(text, file_type):
    """Extract structured data using OpenAI with enhanced location detection"""
    prompt = f"""You are an expert at extracting structured data from insurance payout grids. These grids can be in tabular format, lists, or descriptive text, and may have dynamic columns/rows. Your goal is to parse the entire text intelligently, identifying all entries even if the structure is irregular, merged, or pivoted.

Key Guidelines:
- Handle dynamic structures: Columns might be locations, payins, segments, or mixed. Rows might represent segments, with sub-details for policy types, ages, etc.
- Identify company: From text or infer as 'Unknown'. Common: ICICI, Reliance, Chola, Sompo, Kotak, Magma, Bajaj, Digit, Liberty, Future, Tata, IFFCO, Royal, SBI, Zuno, HDFC, Shriram.
- Segment normalization: Map to standard names like 'PVT CAR COMP + SAOD', 'TW TP', 'GCV 3W', 'SCHOOL BUS', 'MISD'. Detect from context, e.g., 'Pvt Car Package' -> 'PVT CAR COMP + SAOD'. Also please note if MISC is given then it is MISD only but just a typo, so consider MISC as MISD and process the same way.
- Policy type: 'Package' or 'Comp' -> 'Comprehensive', 'TP' -> 'TP', 'SAOD' -> 'SAOD'. If not specified, infer from segment.
- Location: CRITICAL - Recognize city names (Ahmedabad, Surat, Mumbai, Delhi, Bangalore, Chennai, etc.), state names (Gujarat, Maharashtra, Karnataka, etc.), and RTO codes (AS1, AS2, GJ01, MH01, etc.) as LOCATIONS even if no "location" label is present. If you see these as column headers or values, treat them as locations. Also please ignore slight spelling mistakes
- Payin: Extract percentages or numbers (e.g., 0.25 -> 25%, 65 -> 65). Ignore non-numeric. If 0 then put 0 only there.
- Other info/Remarks: Capture age (e.g., '>=1<=6'), fuel (e.g., 'all'), notes (e.g., 'Brand New'), logic.
- Handle pivoted tables: If rows are segments and columns are locations with payins, create an entry per cell. Traverse the table correctly, associating each payin with the correct segment and location.
- Handle multiple entries per row: If a row has multiple segments or policy types, split into separate objects.
- Ignore empty/header rows. Output only valid data rows.
- Be robust: If structure is unclear, reason step-by-step in your mind to map to fields.

Example Input (pivoted):
row1: Segment,Policy,Fuel Type,Ahmedabad,Surat,Mumbai
row2: GCV <= 2.5,Package,all,65,60,55

Example Output:
[
  {{"company_name": "Unknown", "segment": "GCV <= 2.5", "policy_type": "Package", "location": "Ahmedabad", "payin": "65", "other_info": "all"}},
  {{"company_name": "Unknown", "segment": "GCV <= 2.5", "policy_type": "Package", "location": "Surat", "payin": "60", "other_info": "all"}},
  {{"company_name": "Unknown", "segment": "GCV <= 2.5", "policy_type": "Package", "location": "Mumbai", "payin": "55", "other_info": "all"}}
]

TEXT TO ANALYZE (may be rows from Excel or OCR text):
{text}

Return ONLY a list of JSON objects, one per extracted entry. Ensure valid JSON.
"""
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1
    )
    
    response_text = response.choices[0].message.content.strip()
    
    if response_text.startswith('```'):
        response_text = re.sub(r'^```json\n|^```\n|```$', '', response_text, flags=re.MULTILINE).strip()
    
    try:
        result = json.loads(response_text)
        if not isinstance(result, list):
            result = [result]
    except json.JSONDecodeError as e:
        print(f"\n⚠️ JSON Parse Error: {e} - Using fallback empty entry")
        print(f"Problematic text: {response_text[:500]}...")
        result = [{
            "company_name": None,
            "segment": None,
            "policy_type": None,
            "location": None,
            "payin": None,
            "other_info": "Extraction failed"
        }]
    
    return result


def parse_pdf_content(path):
    """Extract text from PDF"""
    reader = PdfReader(path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    
    print("\n" + "="*60)
    print("PDF FILE PARSED")
    print("="*60)
    print(f"\nPages: {len(reader.pages)}")
    print("\n--- Raw Text Preview ---\n")
    print(text[:1500])
    if len(text) > 1500:
        print("\n(Truncated...)")
    
    extracted_list = extract_structured_data_single(text, "PDF")
    structured_data = []
    
    for extracted in extracted_list:
        normalized = normalize_extracted_data(extracted, path)
        normalized["PAYOUT"], normalized["CALCULATION EXPLANATION"] = calculate_payout(normalized)
        structured_data.append(normalized)
    
    print("\n" + "="*60)
    print("STRUCTURED DATA EXTRACTION")
    print("="*60)
    for idx, entry in enumerate(structured_data):
        print(f"\nRow {idx + 1}:")
        for key, value in entry.items():
            if key == "REMARK":
                print(f"{key}: {', '.join(value) if value else 'None'}")
            else:
                print(f"{key}: {value if value else 'Not Found'}")
    
    return {
        "file_type": "PDF",
        "file_path": path,
        "num_pages": len(reader.pages),
        "raw_text": text,
        "structured_data": structured_data
    }


def parse_image_content(path):
    """Extract text from image using Vision API"""
    with open(path, "rb") as img_file:
        img_data = base64.b64encode(img_file.read()).decode('utf-8')
    
    img = Image.open(path)
    
    ext = Path(path).suffix.lower()
    mime_types = {
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.webp': 'image/webp'
    }
    mime_type = mime_types.get(ext, 'image/jpeg')
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "Extract all text from this image accurately, preserving structure like tables, rows, columns as much as possible. Output in a readable text format, using markdown for tables if applicable. Provide only the extracted text."
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime_type};base64,{img_data}"
                        }
                    }
                ]
            }
        ],
        max_tokens=1500
    )
    
    text = response.choices[0].message.content
    
    print("\n" + "="*60)
    print("IMAGE FILE PARSED")
    print("="*60)
    print(f"\nSize: {img.size[0]}×{img.size[1]} pixels")
    print(f"Mode: {img.mode}")
    print(f"\nFilename: {Path(path).name}")
    
    filename_company, filename_location = extract_from_filename(path)
    print(f"📁 Filename Company: {filename_company}")
    print(f"📍 Filename Location: {filename_location}")
    
    print("\n--- Raw Extracted Text (OCR) ---\n")
    print(text)
    
    extracted_list = extract_structured_data_single(text, "Image")
    structured_data = []
    
    for extracted in extracted_list:
        normalized = normalize_extracted_data(extracted, path)
        normalized["PAYOUT"], normalized["CALCULATION EXPLANATION"] = calculate_payout(normalized)
        structured_data.append(normalized)
    
    print("\n" + "="*60)
    print("STRUCTURED DATA EXTRACTION")
    print("="*60)
    for idx, entry in enumerate(structured_data):
        print(f"\nRow {idx + 1}:")
        for key, value in entry.items():
            if key == "REMARK":
                print(f"{key}: {', '.join(value) if value else 'None'}")
            else:
                print(f"{key}: {value if value else 'Not Found'}")
    
    return {
        "file_type": "Image",
        "file_path": path,
        "image_size": img.size,
        "image_mode": img.mode,
        "raw_text": text,
        "structured_data": structured_data
    }


def parse_file(file_path=None):
    """Universal file parser - handles all file types"""
    if file_path is None:
        file_path = input("Enter file path: ")
    
    if not os.path.exists(file_path):
        return {"error": f"File not found: {file_path}"}
    
    ext = Path(file_path).suffix.lower()
    
    try:
        if ext == '.pdf':
            return parse_pdf_content(file_path)
            
        elif ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            print("\n" + "="*60)
            print("EXCEL FILE DETECTED")
            print("="*60)
            
            json_data = excel_to_json_converter(file_path)
            return parse_converted_json(json_data, original_excel_path=file_path)
                
        elif ext in ['.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif']:
            return parse_image_content(file_path)
            
        elif ext == '.json':
            with open(file_path, 'r', encoding='utf-8') as f:
                json_preview = json.load(f)
            
            if isinstance(json_preview, dict) and "sheets" in json_preview:
                print("\n📊 Detected converted Excel JSON format")
                return parse_converted_json(file_path)
            else:
                return {"error": "Standard JSON format not implemented. Please use Excel or PDF."}
                
        else:
            return {"error": f"Unsupported file type: {ext}"}
            
    except Exception as e:
        import traceback
        return {"error": f"Error: {str(e)}\n{traceback.format_exc()}"}


if __name__ == "__main__":
    result = parse_file()
    
    if "error" in result:
        print(f"\n❌ Error: {result['error']}")
    else:
        print("\n" + "="*60)
        print("✅ PARSING COMPLETE")
        print("="*60)
        print(f"\nFile Type: {result['file_type']}")
        
        entries = result.get('structured_data', [])
        
        if entries:
            output_data = {
                "COMPANY NAME": [],
                "ORIGINAL_SEGMENT": [],
                "SEGMENT": [],
                "POLICY TYPE": [],
                "LOCATION": [],
                "PAYIN": [],
                "PAYOUT": [],
                "REMARK": [],
                "CALCULATION EXPLANATION": []
            }
            
            for entry in entries:
                output_data["COMPANY NAME"].append(entry.get("COMPANY NAME", ""))
                output_data["ORIGINAL_SEGMENT"].append(entry.get("ORIGINAL_SEGMENT", ""))
                output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
                output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
                output_data["LOCATION"].append(entry.get("LOCATION", ""))
                output_data["PAYIN"].append(entry.get("PAYIN", ""))
                output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
                output_data["REMARK"].append("; ".join(entry.get("REMARK", [])))
                output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
            
            output_df = pd.DataFrame(output_data)
            output_file = "grid_output.xlsx"
            output_df.to_excel(output_file, index=False)
            
            print(f"\n📊 Excel file created: {output_file}")
            print(f"   Total entries: {len(entries)}")
            print("\n✅ Summary:")
            print(f"   - Company: {output_data['COMPANY NAME'][0] if output_data['COMPANY NAME'] else 'N/A'}")
            print(f"   - Unique Segments: {len(set(s for s in output_data['SEGMENT'] if s))}")
            print(f"   - Unique Locations: {len(set(l for l in output_data['LOCATION'] if l))}")
            
            if any(output_data['PAYIN']):
                valid_payins = [p for p in output_data['PAYIN'] if p is not None and isinstance(p, (int, float))]
                if valid_payins:
                    print(f"   - Payin Range: {min(valid_payins):.1f}% - {max(valid_payins):.1f}%")
            
            print("\n📋 Sample Entries:")
            for i in range(min(5, len(entries))):
                entry = entries[i]
                print(f"   {i+1}. {entry['ORIGINAL_SEGMENT']} | {entry.get('POLICY TYPE', 'N/A')} | {entry.get('LOCATION', 'N/A')} | {entry['PAYIN']}% → {entry['PAYOUT']}")
        else:
            print("\n⚠️  No entries extracted from file")

