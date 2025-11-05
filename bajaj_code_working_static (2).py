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

# Embedded Formula Data
FORMULA_DATA = [
    {"LOB": "TW", "SEGMENT": "1+5", "INSURER": "BAJAJ", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "INSURER": "BAJAJ", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin Above 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "INSURER": "BAJAJ", "PO": "90% of Payin", "REMARKS": "All Fuel"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin Above 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "BAJAJ", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "BAJAJ", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "BAJAJ", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "BAJAJ", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "BAJAJ", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "BAJAJ", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "BAJAJ", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "BAJAJ", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "BAJAJ", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "BUS", "SEGMENT": "STAFF BUS", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "BAJAJ", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "BAJAJ", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "MISD", "SEGMENT": "MISD", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "TRACTOR", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "Misc", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"}
]

COMPANY_MAPPING = {
    'icici': 'ICICI', 'reliance': 'Reliance', 'chola': 'Chola', 'sompo': 'Sompo',
    'kotak': 'Kotak', 'magma': 'Magma', 'bajaj': 'Bajaj', 'digit': 'Digit', 
    'liberty': 'Liberty', 'future': 'Future', 'tata': 'Tata', 'iffco': 'IFFCO', 
    'royal': 'Royal', 'sbi': 'SBI', 'zuno': 'Zuno', 'hdfc': 'HDFC', 'shriram': 'Shriram'
}

def extract_from_filename(file_path):
    """Extract company name and location from filename intelligently"""
    filename = str(file_path).lower()
    base_name = Path(file_path).stem.lower()

    for key, value in COMPANY_MAPPING.items():
        if key in filename:
            return value, _extract_location(filename)

    words = re.findall(r'\b[a-zA-Z]{3,}\b', base_name)
    potential_companies = [w.title() for w in words if w.title() not in {'Insurance', 'General', 'Grid', 'Payout', 'Payin', 'Excel', 'Pdf', 'Img'}]
    if potential_companies:
        company = potential_companies[0]
        return company, _extract_location(filename)

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


def check_approval_status(final_remarks_value):
    """
    Check if the final remarks column indicates approval or rejection
    Returns: ('APPROVED', 'REJECTED', or 'UNKNOWN')
    """
    if not final_remarks_value or pd.isna(final_remarks_value):
        return 'UNKNOWN'
    
    remarks_str = str(final_remarks_value).upper().strip()
    
    # Check for approval keywords
    approval_keywords = ['OK', 'YES', 'APPROVED', 'ACCEPT', 'ACCEPTED', 'APPROVE']
    rejection_keywords = ['NO', 'DECLINE', 'DECLINED', 'REJECT', 'REJECTED', 'NOT ACCEPTED', 'NOT APPROVED']
    
    for keyword in rejection_keywords:
        if keyword in remarks_str:
            return 'REJECTED'
    
    for keyword in approval_keywords:
        if keyword in remarks_str:
            return 'APPROVED'
    
    return 'UNKNOWN'


def extract_requirement_adjustment(requirement_value):
    """
    Extract adjustment percentage from requirement column
    Returns: (adjustment_value, adjustment_type) e.g., (2.5, 'ADD') or (0, 'NONE')
    """
    if not requirement_value or pd.isna(requirement_value):
        return 0, 'NONE'
    
    req_str = str(requirement_value).strip()
    
    # Pattern to match: Grid + 2.5% or Grid +2.5% or +2.5% or -2.5%
    pattern = r'([+\-])\s*(\d+(?:\.\d+)?)\s*%?'
    match = re.search(pattern, req_str)
    
    if match:
        sign = match.group(1)
        value = float(match.group(2))
        
        if sign == '+':
            return value, 'ADD'
        elif sign == '-':
            return value, 'SUBTRACT'
    
    return 0, 'NONE'


def detect_segment_from_worksheet_name(sheet_name):
    """
    SMART SEGMENT DETECTION FROM WORKSHEET NAME
    Returns: (detected_segment, confidence_level)
    """
    sheet_lower = str(sheet_name).lower().strip()
    
    # High confidence patterns
    if any(kw in sheet_lower for kw in ['private car satp', 'pvt car satp', 'pc satp', 'car satp']):
        return "PVT CAR TP", "HIGH"
    
    if any(kw in sheet_lower for kw in ['private car', 'pvt car', 'pc ', 'car']):
        if 'package' in sheet_lower or 'comp' in sheet_lower or 'saod' in sheet_lower:
            return "PVT CAR COMP + SAOD", "HIGH"
        return "PVT CAR", "MEDIUM"
    
    if any(kw in sheet_lower for kw in ['tw_bike', 'tw-bike', 'tw bike', 'two wheeler', '2w', 'bike']):
        if 'satp' in sheet_lower or 'tp' in sheet_lower:
            return "TW TP", "HIGH"
        if 'saod' in sheet_lower or 'comp' in sheet_lower or 'package' in sheet_lower:
            return "TW SAOD + COMP", "HIGH"
        if '1+5' in sheet_lower or 'grid' in sheet_lower or 'new' in sheet_lower:
            return "1+5", "HIGH"
        return "TW", "MEDIUM"
    
    if any(kw in sheet_lower for kw in ['tw- scooter', 'tw-scooter', 'tw scooter', 'scooter']):
        if 'satp' in sheet_lower or 'tp' in sheet_lower:
            return "TW TP", "HIGH"
        if 'saod' in sheet_lower or 'comp' in sheet_lower:
            return "TW SAOD + COMP", "HIGH"
        return "TW", "MEDIUM"
    
    if any(kw in sheet_lower for kw in ['school bus', 'schoolbus']):
        return "SCHOOL BUS", "HIGH"
    
    if any(kw in sheet_lower for kw in ['staff bus', 'staffbus']):
        return "STAFF BUS", "HIGH"
    
    if 'bus' in sheet_lower:
        return "SCHOOL BUS", "LOW"
    
    if 'taxi' in sheet_lower or 'cab' in sheet_lower:
        return "TAXI", "HIGH"
    
    if any(kw in sheet_lower for kw in ['satp - cv', 'satp-cv', 'cv satp', 'cv tp']):
        return "All GVW & PCV 3W, GCV 3W", "MEDIUM"
    
    if any(kw in sheet_lower for kw in ['cv', 'gcv', 'pcv', 'commercial vehicle']):
        return "All GVW & PCV 3W, GCV 3W", "MEDIUM"
    
    if any(kw in sheet_lower for kw in ['3w', '3 w', 'three wheeler']):
        if 'pcv' in sheet_lower:
            return "PCV 3W", "HIGH"
        if 'gcv' in sheet_lower:
            return "GCV 3W", "HIGH"
        return "PCV 3W", "MEDIUM"
    
    if any(kw in sheet_lower for kw in ['misd', 'misc', 'tractor','MISCAgri-Trailer','miscagri-trailer','agri']):
        return "MISD", "HIGH"
    
    return None, "NONE"


def ask_user_for_segment_confirmation(detected_segment, sheet_name, confidence_level):
    """
    Ask user to confirm or provide segment for the worksheet.
    Returns: (final_segment, policy_type) or (None, None) if user skips
    """
    print(f"\n{'='*60}")
    print(f"📋 WORKSHEET: '{sheet_name}'")
    print(f"{'='*60}")
    
    if detected_segment:
        print(f"🔍 Detected segment from worksheet name: '{detected_segment}' (Confidence: {confidence_level})")
        print(f"\nOptions:")
        print(f"  1. Use '{detected_segment}' as segment (Press ENTER or type 'yes')")
        print(f"  2. Enter a different segment (type the segment name)")
        print(f"  3. Skip this worksheet (type 'skip')")
        
        user_input = input(f"\nYour choice: ").strip()
        
        if user_input.lower() == 'skip':
            print(f"⏭️  Skipping worksheet '{sheet_name}'")
            return None, None
        
        if user_input == "" or user_input.lower() in ['yes', 'y', '1']:
            final_segment = detected_segment
        else:
            final_segment = user_input
    else:
        print(f"⚠️ Could not detect segment from worksheet name: '{sheet_name}'")
        print(f"\nAvailable segments:")
        print("  - TW (Two Wheeler), TW SAOD + COMP, TW TP, 1+5")
        print("  - PVT CAR COMP + SAOD, PVT CAR TP")
        print("  - CV, Upto 2.5 GVW, PCV 3W, GCV 3W, All GVW")
        print("  - SCHOOL BUS, STAFF BUS, TAXI, MISD, TRACTOR")
        print("  - Type 'skip' to skip this worksheet")
        
        final_segment = input(f"\nEnter segment for '{sheet_name}': ").strip()
        
        if final_segment.lower() == 'skip':
            print(f"⏭️  Skipping worksheet '{sheet_name}'")
            return None, None
        
        if not final_segment:
            print(f"❌ No segment provided. Skipping worksheet.")
            return None, None
    
    # Normalize the segment
    normalized_segment = normalize_segment(final_segment)
    
    # Check if policy type is needed
    needs_policy_clarification = False
    segment_upper = normalized_segment.upper()
    
    if any(x in segment_upper for x in ['PVT CAR', 'PRIVATE CAR']) and \
       not any(x in segment_upper for x in ['TP', 'COMP', 'SAOD', 'PACKAGE']):
        needs_policy_clarification = True
    
    if any(x in segment_upper for x in ['TW', 'TWO WHEELER']) and \
       not any(x in segment_upper for x in ['TP', 'COMP', 'SAOD', '1+5', 'GRID']):
        needs_policy_clarification = True
    
    policy_type = None
    
    if needs_policy_clarification:
        print(f"\n🔍 Segment detected as: '{normalized_segment}'")
        print(f"⚠️ Policy type is missing. Please specify:")
        print(f"   1. TP (Third Party / SATP)")
        print(f"   2. COMP (Comprehensive)")
        print(f"   3. SAOD (Stand Alone Own Damage)")
        print(f"   4. COMP + SAOD")
        print(f"   5. Package (same as Comprehensive)")
        print(f"   6. 1+5 (for TW New/Grid)")
        
        policy_input = input(f"\nEnter policy type: ").strip()
        
        if policy_input:
            policy_type = normalize_policy_type(policy_input)
            
            if 'PVT CAR' in normalized_segment and policy_type:
                if policy_type in ['TP', 'Third Party', 'SATP', 'satp', 'Satp']:
                    normalized_segment = "PVT CAR TP"
                elif policy_type in ['Comprehensive', 'COMP + SAOD', 'SAOD', 'COMP']:
                    normalized_segment = "PVT CAR COMP + SAOD"
            
            elif 'TW' in normalized_segment and policy_type:
                if policy_type in ['TP', 'Third Party', 'SATP', 'Satp', 'satp']:
                    normalized_segment = "TW TP"
                elif policy_type in ['Comprehensive', 'COMP + SAOD', 'SAOD', 'COMP']:
                    normalized_segment = "TW SAOD + COMP"
                elif '1+5' in policy_input.upper() or 'GRID' in policy_input.upper():
                    normalized_segment = "1+5"
    
    print(f"✓ Final segment: '{normalized_segment}'")
    if policy_type:
        print(f"✓ Policy type: '{policy_type}'")
    
    return normalized_segment, policy_type


def get_all_specific_insurers_for_lob_segment(lob: str, segment: str) -> List[str]:
    """Collect all specific insurer names for given LOB+Segment"""
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
    """
    Calculate payout with support for Requirement and Final Remarks columns
    """
    company_name = normalized_data.get("COMPANY NAME", "BAJAJ").strip()
    segment = normalized_data.get("SEGMENT", "BAJAJ").upper().strip()
    policy_type = (normalized_data.get("POLICY TYPE", "") or "").upper().strip()
    payin = normalized_data.get("PAYIN", 0)
    location = normalized_data.get("LOCATION", "N/A")
    
    # NEW: Get requirement and final remarks
    requirement_value = normalized_data.get("REQUIREMENT", None)
    final_remarks_value = normalized_data.get("FINAL_REMARKS", None)

    if payin is None or payin == "" or str(payin).strip() == "" or str(payin).strip().lower() == "none":
        payin_value = 0.0
    else:
        try:
            payin_value = float(str(payin).replace('%', '').strip())
        except (ValueError, AttributeError):
            payin_value = 0.0

    if payin_value == 0:
        return "0.00%", "Payin is 0% → Payout is 0% (no calculation applied)", []

    payin_category = (
        "Payin Below 20%" if payin_value <= 20 else
        "Payin 21% to 30%" if 21 <= payin_value <= 30 else
        "Payin 31% to 50%" if 31 <= payin_value <= 50 else
        "Payin Above 50%"
    )

    segment_norm = segment.replace(" ", "")
    lob = "TW"
    if "BUS" in segment_norm or "SCHOOL" in segment_norm or "STAFF" in segment_norm:
        lob = "BUS"
    elif "PVT" in segment_norm or "CAR" in segment_norm:
        lob = "PVT CAR"
    elif "CV" in segment_norm or "GVW" in segment_norm or "PCV" in segment_norm or "GCV" in segment_norm:
        lob = "CV"
    elif "TAXI" in segment_norm:
        lob = "TAXI"
    elif "MISD" in segment_norm or "TRACTOR" in segment_norm:
        lob = "MISD"

    segment_policy = None
    if "COMP" in segment or "SAOD" in segment or "TP" in segment:
        segment_policy = next((pt for pt in ['COMP', 'SAOD', 'TP'] if pt in segment), None)
    else:
        segment_policy = policy_type

    specific_insurers = get_all_specific_insurers_for_lob_segment(lob, segment)

    candidate_rules = []
    company_upper = company_name.upper()

    for rule in FORMULA_DATA:
        rule_lob = rule["LOB"].upper()
        rule_segment = rule["SEGMENT"].upper()
        
        if rule_lob != lob.upper():
            continue
        
        segment_matches = (rule_segment == segment.upper() or 
                          rule_segment in segment.upper() or 
                          segment.upper() in rule_segment)
        
        if not segment_matches:
            continue

        insurers = [i.strip().upper() for i in rule["INSURER"].split(',')]
        insurer_type = None
        specificity = 0

        for ins in insurers:
            if ins in ["ALL COMPANIES", "REST OF COMPANIES"]:
                continue
            if (ins in company_upper or company_upper in ins or
                    any(word in company_upper for word in ins.split())):
                insurer_type = "SPECIFIC"
                specificity = 3
                break

        if not insurer_type and "REST OF COMPANIES" in insurers:
            if not is_company_in_specific_list(company_name, specific_insurers):
                insurer_type = "REST"
                specificity = 2

        if not insurer_type and "ALL COMPANIES" in insurers:
            insurer_type = "ALL"
            specificity = 1

        if not insurer_type:
            continue

        remarks = rule.get("REMARKS", "").upper().strip()
        remarks_score = 0

        if remarks == "NIL":
            remarks_score = 10
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

    if candidate_rules:
        candidate_rules.sort(key=lambda x: x[1], reverse=True)
        matched_rule = candidate_rules[0][0]
        explanation = f"Match: LOB={lob}, Segment={segment}, Insurer={matched_rule['INSURER']}, Remarks={matched_rule.get('REMARKS', 'NIL')}, Formula={matched_rule['PO']}"
    else:
        matched_rule = {"PO": "90% of Payin", "INSURER": "Default", "REMARKS": "NIL"}
        explanation = f"No Match: Default Formula=90% of Payin"

    po = matched_rule["PO"]
    
    if not any(indicator in po for indicator in ['%', 'flat', 'Payin', '-']):
        return po, f"Match: LOB={lob}, Segment={segment}, Insurer={matched_rule['INSURER']}, Remarks={matched_rule.get('REMARKS', 'NIL')}, Payout={po} (text value, no calculation)", []
    
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
        else:
            factor = 1.0
        payout = payin_value * factor
        op_str = f"* {factor} = {payout:.2f}%"
    elif "-2%" in po:
        payout = payin_value - 2
        op_str = f"- 2% = {payout:.2f}%"
    elif "-3%" in po:
        payout = payin_value - 3
        op_str = f"- 3% = {payout:.2f}%"
    elif "-4%" in po:
        payout = payin_value - 4
        op_str = f"- 4% = {payout:.2f}%"
    elif "-5%" in po:
        payout = payin_value - 5
        op_str = f"- 5% = {payout:.2f}%"

    payout = max(0, payout)
    explanation += f"; Calculated: {payin_value:.2f}% {op_str}"

    # NEW: Check Final Remarks and Requirement columns
    additional_remarks = []
    
    approval_status = check_approval_status(final_remarks_value)
    
    if approval_status == 'APPROVED':
        # Check if there's a requirement adjustment
        adjustment_value, adjustment_type = extract_requirement_adjustment(requirement_value)
        
        if adjustment_type == 'ADD' and adjustment_value > 0:
            # Add the adjustment to payout
            original_payout = payout
            payout = payout + adjustment_value
            explanation += f"; Requirement Adjustment: +{adjustment_value}% → Final Payout: {payout:.2f}%"
            additional_remarks.append(f"OK - Approved and increased {adjustment_value}%")
        elif adjustment_type == 'SUBTRACT' and adjustment_value > 0:
            # Subtract the adjustment from payout
            original_payout = payout
            payout = payout - adjustment_value
            payout = max(0, payout)
            explanation += f"; Requirement Adjustment: -{adjustment_value}% → Final Payout: {payout:.2f}%"
            additional_remarks.append(f"OK - Approved and decreased {adjustment_value}%")
        else:
            additional_remarks.append("OK - Approved")
    
    elif approval_status == 'REJECTED':
        additional_remarks.append("NO - Declined")
    
    return f"{payout:.2f}%", explanation, additional_remarks


def extract_percentage(value) -> Optional[float]:
    """Extract numeric percentage from various formats"""
    if pd.isna(value) or value is None:
        return None
    
    try:
        num_val = float(value)
        if 0 < num_val < 1:
            return num_val * 100
        elif 0 <= num_val <= 100:
            return num_val
    except:
        pass
    
    value_str = str(value)
    match = re.search(r'(\d+(?:\.\d+)?)\s*%?', value_str)
    if match:
        num_val = float(match.group(1))
        if 0 <= num_val <= 100:
            return num_val
    
    return None


def normalize_segment(segment_str):
    """Normalize segment names to standard format"""
    if not segment_str or pd.isna(segment_str):
        return "Unknown"
    
    segment_lower = re.sub(r'\s+', ' ', str(segment_str).lower().strip())

    if segment_lower == "apts":
        return segment_str

    if 'other than school bus' in segment_lower or 'other then school bus' in segment_lower:
        return "STAFF BUS"
    
    if 'pcv' in segment_lower and ('4w' in segment_lower or '4 w' in segment_lower or 'four wheeler' in segment_lower):
        if 'school' not in segment_lower:
            return "STAFF BUS"
    
    if re.search(r'sc\s*[<≤]\s*\d+', segment_lower) or re.search(r'seating.*?[<≤]\s*\d+', segment_lower):
        if 'school' not in segment_lower:
            return "STAFF BUS"

    if 'pcv school buses' in segment_lower or 'school bus' in segment_lower or 'school buses' in segment_lower:
        return "SCHOOL BUS"
    
    if 'staff bus' in segment_lower:
        return "STAFF BUS"
    
    if 'bus' in segment_lower and 'school' not in segment_lower and 'staff' not in segment_lower:
        return "SCHOOL BUS"
    
    if 'taxi' in segment_lower:
        return "TAXI"
    
    if any(x in segment_lower for x in ['3w', '3 w', 'three wheeler', '3wheeler']):
        if 'pcv' in segment_lower or 'passenger' in segment_lower:
            return "PCV 3W"
        elif 'gcv' in segment_lower or 'goods' in segment_lower:
            return "GCV 3W"
        return "PCV 3W"
    
    if any(x in segment_lower for x in ['tw', '2w', 'two wheeler', 'bike', 'scooter', 'mc', 'sc', 'ev', 'motor cycle', 'tw new']):
        if 'saod' in segment_lower or 'comp' in segment_lower:
            return "TW SAOD + COMP"
        elif 'tp' in segment_lower or 'satp' in segment_lower:
            return "TW TP"
        elif '1+5' in segment_lower or 'grid' in segment_lower or 'new' in segment_lower or 'tw new' in segment_lower:
            return "1+5"
        return "TW"
    
    if any(x in segment_lower for x in ['pvt car', 'private car', 'car', 'pvt car tp']):
        if 'comp' in segment_lower or 'saod' in segment_lower or 'package' in segment_lower:
            return "PVT CAR COMP + SAOD"
        elif 'tp' in segment_lower or 'pvt car tp' in segment_lower:
            return "PVT CAR TP"
        return "PVT CAR COMP + SAOD"
    
    if any(x in segment_lower for x in ['cv', 'gcv', 'pcv', 'lcv', 'gvw', 'tn', 'tonnage', '2.5', '3.5']):
        if 'upto 2.5' in segment_lower or '0-2.5' in segment_lower or '2.5 gvw' in segment_lower:
            return "Upto 2.5 GVW"
        elif 'pcv 3w' in segment_lower or 'gcv 3w' in segment_lower or '3w gcv' in segment_lower or '3w pcv' in segment_lower:
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


def detect_columns_intelligently(sheet_data):
    """
    Intelligently detect column mappings including SEGMENT, REQUIREMENT, and FINAL_REMARKS columns.
    Returns: dict with 'segment', 'location', 'payin', 'remarks', 'district', 'requirement', 'final_remarks'
    """
    header_map = {}
    remarks_columns = []
    
    for i in range(min(5, len(sheet_data))):
        row = sheet_data[i]
        for col_key, val in row.items():
            if val is None or str(val).strip() == "":
                continue
            
            val_str = str(val).lower().strip()
            
            if any(kw in val_str for kw in ["segment", "category", "product", "product type", "vehicle type", "class"]):
                if 'segment' not in header_map:
                    header_map['segment'] = col_key
                    print(f"   ✓ Detected SEGMENT column: {col_key} ('{val}')")
            
            if any(kw in val_str for kw in ["rto", "state", "location", "region", "zone", "rto-statename", "rto-state"]):
                if 'location' not in header_map:
                    header_map['location'] = col_key
                    print(f"   ✓ Detected LOCATION column: {col_key} ('{val}')")
            
            if any(kw in val_str for kw in ["district", "city", "sub-location", "area", "town"]):
                if 'district' not in header_map:
                    header_map['district'] = col_key
                    print(f"   ✓ Detected DISTRICT/CITY column: {col_key} ('{val}')")
            
            # NEW: Detect Requirement column
            if any(kw in val_str for kw in ["requirement", "req", "grid requirement", "adjustment"]):
                if 'requirement' not in header_map:
                    header_map['requirement'] = col_key
                    print(f"   ✓ Detected REQUIREMENT column: {col_key} ('{val}')")
            
            # NEW: Detect Final Remarks column
            if any(kw in val_str for kw in ["final remark", "final remarks", "uw guidelines to be continued", "approval status", "status"]):
                if 'final_remarks' not in header_map:
                    header_map['final_remarks'] = col_key
                    print(f"   ✓ Detected FINAL REMARKS column: {col_key} ('{val}')")
            
            if any(kw in val_str for kw in ["rsme", "rate", "max rate", "guideline", "satp", "payin", "ncb", "max rate for rsme"]):
                if not any(remark_kw in val_str for remark_kw in ["remark", "uw guidelines (satp)", "guidelines (satp)", "enabler"]):
                    if 'payin' not in header_map:
                        header_map['payin'] = col_key
                        print(f"   ✓ Detected PAYIN column: {col_key} ('{val}')")
            
            if any(kw in val_str for kw in ["remark", "uw guideline", "uw guidelines (satp)", "guidelines (satp)", "enabler", "uw guidelines"]):
                if 'final_remarks' not in header_map:  # Only add to remarks_columns if not already identified as final_remarks
                    remarks_columns.append(col_key)
                    print(f"   ✓ Detected REMARKS column {len(remarks_columns)}: {col_key} ('{val}')")
    
    if remarks_columns:
        header_map['remarks'] = remarks_columns
    
    return header_map


def extract_text_from_sheet(sheet_data, sheet_name):
    """Converts a list of row-dictionaries into text for the AI"""
    text = f"Sheet Name: {sheet_name}\n"
    for idx, row in enumerate(sheet_data, 1):
        row_values = [str(v) for v in row.values() if v is not None]
        row_str = ", ".join(row_values)
        text += f"row{idx}: {row_str}\n"
    return text


def excel_to_json_converter(excel_path, output_json_path=None):
    """Convert Excel to JSON"""
    print("\n" + "="*60)
    print("EXCEL TO JSON CONVERTER")
    print("="*60)
    
    excel_file = pd.ExcelFile(excel_path)
    sheet_names = excel_file.sheet_names
    
    print(f"\n📊 Found {len(sheet_names)} worksheet(s):")
    for i, sheet in enumerate(sheet_names, 1):
        print(f"   {i}. {sheet}")
    
    all_sheets_data = {}
    
    for sheet_name in sheet_names:
        print(f"\n🔄 Processing sheet: '{sheet_name}'")
        
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        print(f"   Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        
        df.columns = [f"col_{i}" for i in range(len(df.columns))]
        df = df.where(pd.notna(df), None)
        json_data = df.to_dict('records')
        
        all_sheets_data[sheet_name] = json_data
        print(f"   ✓ Converted to {len(json_data)} JSON records")
    
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
    
    print(f"\n✅ JSON file saved: {output_json_path}")
    print(f"   Total sheets: {len(sheet_names)}")
    print(f"   Total records: {sum(len(v) for v in all_sheets_data.values())}")
    
    return output_data


def parse_converted_json(json_path_or_data, original_excel_path=None, user_company=None):
    """
    WORKSHEET-BASED PARSING with SMART SEGMENT DETECTION and REQUIREMENT/FINAL REMARKS support
    Parse converted JSON with intelligent column detection for each worksheet.
    """
    print("\n" + "="*60)
    print("PARSING CONVERTED JSON (Worksheet-Based with Auto Segment Detection)")
    print("="*60)
    
    if isinstance(json_path_or_data, str):
        with open(json_path_or_data, 'r', encoding='utf-8') as f:
            data = json.load(f)
        json_path = json_path_or_data
    else:
        data = json_path_or_data
        json_path = None
    
    source_file = original_excel_path or data.get("source_file") or json_path
    filename_company, filename_location = extract_from_filename(source_file) if source_file else (None, None)
    
    if user_company is None:
        print("\n" + "="*60)
        print("USER INPUT REQUIRED")
        print("="*60)
        user_company = input("📝 Enter COMPANY NAME (e.g., BAJAJ, ICICI, Reliance): ").strip()
    
    print(f"\n📁 Source: {source_file}")
    print(f"📁 User-provided Company: {user_company}")
    print(f"📍 Detected Location from filename: {filename_location}")
    
    all_worksheets_data = {}
    
    if "sheets" in data:
        print(f"\n📊 Processing {data['total_sheets']} worksheet(s)...")
        
        for sheet_name, sheet_data in data["sheets"].items():
            print(f"\n{'='*60}")
            print(f"📑 WORKSHEET: '{sheet_name}'")
            print(f"{'='*60}")
            print(f"   Records in sheet: {len(sheet_data)}")
            
            detected_segment, confidence = detect_segment_from_worksheet_name(sheet_name)
            
            header_map = detect_columns_intelligently(sheet_data)
            has_segment_column = 'segment' in header_map
            
            if has_segment_column:
                print(f"   ✓ SEGMENT COLUMN found in data - will use segments from data")
                final_segment = None
                policy_type = None
            else:
                print(f"   ⚠️ No SEGMENT COLUMN found in data")
                
                final_segment, policy_type = ask_user_for_segment_confirmation(
                    detected_segment, sheet_name, confidence
                )
                
                if final_segment is None:
                    print(f"   ⏭️  Skipping worksheet '{sheet_name}'")
                    continue
            
            if not header_map:
                print("   ⚠️ Warning: Could not detect columns automatically. Using default...")
                header_map = {'location': 'col_0', 'payin': 'col_1'}
            
            data_start_index = 0
            for i in range(min(5, len(sheet_data))):
                row = sheet_data[i]
                row_vals = [str(v).lower().strip() for v in row.values() if v]
                if any(k in row_vals for k in ["rto", "state", "location", "rsme", "rate", "guideline", "segment", "category", "requirement", "final remark"]):
                    data_start_index = i + 1
                    print(f"   Found header row at index {i}. Data starts from row {data_start_index}")
                    break
            
            processed_entries = []
            for i in range(data_start_index, len(sheet_data)):
                row = sheet_data[i]
                
                if all(v is None or str(v).strip() == "" for v in row.values()):
                    continue
                
                segment = None
                location = None
                district = None
                payin = None
                payin_text = None
                remarks = []
                requirement = None
                final_remarks = None
                
                segment_col = header_map.get('segment')
                location_col = header_map.get('location')
                district_col = header_map.get('district')
                payin_col = header_map.get('payin')
                remarks_cols = header_map.get('remarks')
                requirement_col = header_map.get('requirement')
                final_remarks_col = header_map.get('final_remarks')
                
                if has_segment_column and segment_col and segment_col in row:
                    segment_raw = row[segment_col]
                    if segment_raw is not None and str(segment_raw).strip():
                        segment = str(segment_raw).strip()
                else:
                    segment = final_segment
                
                if location_col and location_col in row:
                    location = row[location_col]
                
                if district_col and district_col in row:
                    district_val = row[district_col]
                    if district_val is not None and str(district_val).strip():
                        district = str(district_val).strip()
                
                if location and district:
                    combined_location = f"{location}-{district}"
                    location = combined_location
                elif location:
                    location = str(location)
                
                # NEW: Extract requirement value
                if requirement_col and requirement_col in row:
                    requirement = row[requirement_col]
                
                # NEW: Extract final remarks value
                if final_remarks_col and final_remarks_col in row:
                    final_remarks = row[final_remarks_col]
                
                if payin_col and payin_col in row:
                    payin_raw = row[payin_col]
                    if payin_raw is not None:
                        payin_str = str(payin_raw).strip().upper()
                        
                        if "IRDA" in payin_str or (not any(char.isdigit() for char in payin_str) and len(payin_str) > 3):
                            payin_text = payin_raw
                            payin = None
                        else:
                            payin = extract_percentage(payin_raw)
                
                if remarks_cols:
                    if isinstance(remarks_cols, list):
                        for remarks_col in remarks_cols:
                            if remarks_col in row:
                                remarks_val = row[remarks_col]
                                if remarks_val is not None and str(remarks_val).strip():
                                    remarks_str = str(remarks_val).strip()
                                    if remarks_str.lower() not in ["nan", "none", "-", ""]:
                                        remarks.append(remarks_str)
                    else:
                        if remarks_cols in row:
                            remarks_val = row[remarks_cols]
                            if remarks_val is not None and str(remarks_val).strip():
                                remarks_str = str(remarks_val).strip()
                                if remarks_str.lower() not in ["nan", "none", "-", ""]:
                                    remarks.append(remarks_str)
                
                if not remarks and not remarks_cols:
                    for col_key, col_val in row.items():
                        if col_val is None or str(col_val).strip() == "":
                            continue
                        
                        if col_key == location_col or col_key == payin_col or col_key == district_col or col_key == segment_col or col_key == requirement_col or col_key == final_remarks_col:
                            continue
                        
                        val_str = str(col_val).strip()
                        val_lower = val_str.lower()
                        
                        if any(header_kw in val_lower for header_kw in ["rto", "state", "location", "rsme", "rate", "guideline", "max rate", "district", "city", "segment", "requirement", "final remark"]):
                            continue
                        
                        if len(val_str) > 0 and val_str.lower() not in ["nan", "none", "-"]:
                            remarks.append(val_str)
                
                if location is None and payin is None and payin_text is None and segment is None:
                    continue
                
                entry = {
                    "company_name": user_company,
                    "segment": segment,
                    "policy_type": policy_type,
                    "location": location,
                    "payin": payin if payin_text is None else payin_text,
                    "other_info": " | ".join(remarks) if remarks else None,
                    "requirement": requirement,
                    "final_remarks": final_remarks
                }
                
                processed_entries.append(entry)
            
            print(f"   ✓ Extracted {len(processed_entries)} entries from sheet '{sheet_name}'")
            
            final_entries = []
            for extracted in processed_entries:
                normalized = normalize_extracted_data(extracted, source_file)
                
                normalized["COMPANY NAME"] = user_company
                
                if extracted.get("segment"):
                    normalized["SEGMENT"] = normalize_segment(extracted["segment"])
                
                if not normalized["LOCATION"] and filename_location:
                    normalized["LOCATION"] = filename_location
                
                # NEW: Add requirement and final_remarks to normalized data
                normalized["REQUIREMENT"] = extracted.get("requirement")
                normalized["FINAL_REMARKS"] = extracted.get("final_remarks")
                
                payout, explanation, additional_remarks = calculate_payout(normalized)
                normalized["PAYOUT"] = payout
                normalized["CALCULATION EXPLANATION"] = explanation
                
                # NEW: Append additional remarks from requirement/final_remarks check
                if additional_remarks:
                    for add_remark in additional_remarks:
                        if add_remark not in normalized["REMARK"]:
                            normalized["REMARK"].append(add_remark)
                
                final_entries.append(normalized)
            
            all_worksheets_data[sheet_name] = final_entries
            print(f"   ✓ Processed {len(final_entries)} entries for worksheet '{sheet_name}'")
    
    return {
        "file_type": "JSON (Converted from Excel - Worksheet-Based Processing)",
        "file_path": json_path or source_file,
        "worksheets": all_worksheets_data
    }


def extract_policy_type_from_segment(segment_str):
    """Extract policy type from segment string"""
    if not segment_str:
        return None
    
    segment_upper = segment_str.upper()
    
    if "TP" in segment_upper and "SATP" not in segment_upper:
        return "TP"
    
    if "SATP" in segment_upper:
        return "SATP"
    
    if "COMP" in segment_upper:
        return "Comprehensive"
    
    if "SAOD" in segment_upper or "OD" in segment_upper:
        return "SAOD"
    
    if "PACKAGE" in segment_upper:
        return "Comprehensive"
    
    return "Comprehensive"


def normalize_extracted_data(extracted, file_path):
    """Normalize extracted data from AI"""
    normalized = {
        "COMPANY NAME": "Unknown",
        "SEGMENT": None,
        "POLICY TYPE": None,
        "LOCATION": None,
        "PAYIN": None,
        "PAYOUT": "",
        "REMARK": [],
        "CALCULATION EXPLANATION": "",
        "REQUIREMENT": None,
        "FINAL_REMARKS": None
    }
    
    normalized["COMPANY NAME"] = extracted.get("company_name") or "Unknown"
    
    if extracted.get("segment"):
        normalized["SEGMENT"] = normalize_segment(extracted["segment"])
    else:
        normalized["SEGMENT"] = "Unknown"
    
    policy_from_segment = extract_policy_type_from_segment(normalized["SEGMENT"])
    
    if policy_from_segment:
        normalized["POLICY TYPE"] = policy_from_segment
    elif extracted.get("policy_type"):
        normalized["POLICY TYPE"] = normalize_policy_type(extracted["policy_type"])
    else:
        normalized["POLICY TYPE"] = "Comprehensive"
    
    if extracted.get("location"):
        normalized["LOCATION"] = extracted["location"]
    elif extracted.get("Region/RTO/Zone"):
        normalized["LOCATION"] = extracted["Region/RTO/Zone"]

    if extracted.get("payin"):
        payin_raw = extracted["payin"]
        
        if isinstance(payin_raw, str):
            payin_upper = str(payin_raw).upper()
            if "IRDA" in payin_upper or not any(char.isdigit() for char in str(payin_raw)):
                normalized["PAYIN"] = payin_raw
                normalized["REMARK"].append(f"As per IRDA: {payin_raw}")
            else:
                payin_str = str(payin_raw)
                payin_match = re.search(r'(\d+(?:\.\d+)?)', payin_str)
                if payin_match:
                    normalized["PAYIN"] = float(payin_match.group(1))
        else:
            normalized["PAYIN"] = payin_raw
    
    if extracted.get("other_info"):
        other_info_str = str(extracted["other_info"]).strip()
        if other_info_str and other_info_str.lower() not in ["none", "nan", ""]:
            remarks_list = [r.strip() for r in other_info_str.split("|") if r.strip()]
            
            for remark in remarks_list:
                if not any("As per IRDA" in r for r in normalized["REMARK"]):
                    normalized["REMARK"].append(remark)
    
    # NEW: Add requirement and final_remarks
    normalized["REQUIREMENT"] = extracted.get("requirement")
    normalized["FINAL_REMARKS"] = extracted.get("final_remarks")
    
    return normalized


def parse_file(file_path=None, user_company=None):
    """Universal file parser - handles all file types"""
    if file_path is None:
        file_path = input("Enter file path: ")
    
    if not os.path.exists(file_path):
        return {"error": f"File not found: {file_path}"}
    
    ext = Path(file_path).suffix.lower()
    
    try:
        if ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            print("\n" + "="*60)
            print("EXCEL FILE DETECTED (Worksheet-Based Processing)")
            print("="*60)
            
            json_data = excel_to_json_converter(file_path)
            return parse_converted_json(json_data, original_excel_path=file_path, user_company=user_company)
                
        else:
            return {"error": f"Unsupported file type: {ext}. Currently only Excel files are supported for worksheet-based processing."}
            
    except Exception as e:
        import traceback
        return {"error": f"Error: {str(e)}\n{traceback.format_exc()}"}


if __name__ == "__main__":
    result = parse_file(user_company=None)
    
    if "error" in result:
        print(f"\n❌ Error: {result['error']}")
    else:
        print("\n" + "="*60)
        print("✅ PARSING COMPLETE")
        print("="*60)
        print(f"\nFile Type: {result['file_type']}")
        
        worksheets_data = result.get('worksheets', {})
        
        if worksheets_data:
            output_file = "grid_output.xlsx"
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                total_entries = 0
                
                for sheet_name, entries in worksheets_data.items():
                    if not entries:
                        continue
                    
                    output_data = {
                        "COMPANY NAME": [],
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
                        output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
                        output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
                        output_data["LOCATION"].append(entry.get("LOCATION", ""))
                        output_data["PAYIN"].append(entry.get("PAYIN", ""))
                        output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
                        output_data["REMARK"].append("; ".join(entry.get("REMARK", [])))
                        output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
                    
                    output_df = pd.DataFrame(output_data)
                    
                    safe_sheet_name = sheet_name[:31]
                    output_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    
                    print(f"\n   ✓ Worksheet '{sheet_name}': {len(entries)} entries")
                    total_entries += len(entries)
            
            print(f"\n📊 Excel file created: {output_file}")
            print(f"   Total worksheets: {len(worksheets_data)}")
            print(f"   Total entries: {total_entries}")
            
            print(f"\n✅ Summary:")
            for sheet_name, entries in worksheets_data.items():
                if entries:
                    print(f"\n   📋 {sheet_name}:")
                    print(f"      - Entries: {len(entries)}")
                    unique_segments = set(e['SEGMENT'] for e in entries if e['SEGMENT'])
                    if unique_segments:
                        print(f"      - Segments: {', '.join(unique_segments)}")
                    unique_locations = set(e['LOCATION'] for e in entries if e['LOCATION'])
                    if unique_locations:
                        print(f"      - Locations: {len(unique_locations)} unique")
        else:
            print("\n⚠️  No entries extracted from file")
