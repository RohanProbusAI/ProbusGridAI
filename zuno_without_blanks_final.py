# zuno.py: doesnt works for blannk or 0 values
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
    {"LOB": "TW", "SEGMENT": "1+5", "INSURER": "Zuno", "PO": "90% of Payin", "REMARKS": "NIL"},

    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "INSURER": "Zuno", "PO": "90% of Payin", "REMARKS": "NIL"},

    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "Zuno", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "Zuno", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "Zuno", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},    
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "Zuno", "PO": "-5%", "REMARKS": "Payin Above 50%"},

    # Rule 1: Zuno gets flat 21% payout
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "Zuno", "PO": "21% flat", "REMARKS": "NIL"},

    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "INSURER": "Zuno", "PO": "90% of Payin", "REMARKS": "NIL"},
    
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "Zuno", "PO": "90% of Payin", "REMARKS": "Zuno -  21"},

    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "Zuno", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "Zuno", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "Zuno", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "Zuno", "PO": "-5%", "REMARKS": "Payin Above 50%"},

    {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "INSURER": "Zuno", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "BUS", "SEGMENT": "STAFF BUS", "INSURER": "Zuno", "PO": "88% of Payin", "REMARKS": "NIL"},

    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "Zuno", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "Zuno", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "Zuno", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "Zuno", "PO": "-5%", "REMARKS": "Payin Above 50%"},

    {"LOB": "MISD", "SEGMENT": "Misd, Tractor", "INSURER": "All Companies", "PO": "88% of Payin", "REMARKS": "NIL"}
]

COMPANY_MAPPING = {
    'icici': 'ICICI', 'reliance': 'Reliance', 'chola': 'Chola', 'sompo': 'Sompo',
    'kotak': 'Kotak', 'magma': 'Magma', 'bajaj': 'Bajaj', 'digit': 'Digit', 
    'liberty': 'Liberty', 'future': 'Future', 'tata': 'Tata', 'iffco': 'IFFCO', 
    'royal': 'Royal', 'sbi': 'SBI', 'zuno': 'Zuno', 'hdfc': 'HDFC', 'shriram': 'Shriram'
}

def extract_from_filename(file_path):
    """Extract company name and location from filename"""
    filename = str(file_path).lower()
    
    company = None
    for key, value in COMPANY_MAPPING.items():
        if key in filename:
            company = value
            break
    
    location = None
    location_patterns = [
        r'\b(guj|gujarat)\b', r'\b(mh|maharashtra|mumbai)\b', 
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
            location = match.group(0).upper()
            break
    
    return company, location

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
    
    # Bus (check first for PCV School BUSES and other bus types)
    if 'pcv school buses' in segment_lower or 'school bus' in segment_lower or 'school buses' in segment_lower:
        return "SCHOOL BUS"
    if 'staff bus' in segment_lower:
        return "STAFF BUS"
    if 'bus' in segment_lower and 'school' not in segment_lower and 'staff' not in segment_lower:
        return "SCHOOL BUS"
    
    # Taxi
    if 'taxi' in segment_lower:
        return "TAXI"
    
    # 3 Wheeler (before PCV/GCV)
    if any(x in segment_lower for x in ['3w', '3 w', 'three wheeler', '3wheeler']):
        if 'pcv' in segment_lower or 'passenger' in segment_lower:
            return "PCV 3W"
        elif 'gcv' in segment_lower or 'goods' in segment_lower:
            return "GCV 3W"
        return "PCV 3W"
    
    # Two Wheeler
    if any(x in segment_lower for x in ['tw', '2w', 'two wheeler', 'bike', 'scooter', 'mc', 'sc', 'ev', 'motor cycle']):
        if 'saod' in segment_lower or 'comp' in segment_lower:
            return "TW SAOD + COMP"
        elif 'tp' in segment_lower or 'satp' in segment_lower:
            return "TW TP"
        elif '1+5' in segment_lower or 'grid' in segment_lower:
            return "1+5"
        return "TW"
    
    # Private Car
    if any(x in segment_lower for x in ['pvt car', 'private car', 'car']):
        if 'comp' in segment_lower or 'saod' in segment_lower:
            return "PVT CAR COMP + SAOD"
        elif 'tp' in segment_lower:
            return "PVT CAR TP"
        return "PVT CAR COMP + SAOD"
    
    # Commercial Vehicle
    if any(x in segment_lower for x in ['cv', 'gcv', 'pcv', 'lcv', 'gvw', 'tn', 'tonnage']):
        if 'upto 2.5' in segment_lower or '0-2.5' in segment_lower or '2.5 gvw' in segment_lower:
            return "Upto 2.5 GVW"
        elif 'pcv 3w' in segment_lower or 'gcv 3w' in segment_lower:
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

def calculate_payout(normalized_data):
    """
    Calculate payout based on formula rules
    Returns: (payout, calculation_explanation)
    """
    # Get policy details from normalized_data
    company_name = normalized_data.get("COMPANY NAME", "Unknown")
    segment = normalized_data.get("SEGMENT", 'Unknown').upper().strip()
    policy_type = (normalized_data.get("POLICY TYPE", '') or '').upper().strip()
    payin = normalized_data.get("PAYIN", 0)
    payin_str = str(payin) + "%" if payin else "0%"

    if payin is None or payin == 0:
        payin_value = 0.0
    else:
        try:
            payin_value = float(str(payin).replace('%', '').replace(' ', '').strip() or 0)
        except (ValueError, TypeError):
            payin_value = 0.0

    location = normalized_data.get("LOCATION", 'N/A')

    # Classify payin category based on payin_value
    if payin_value <= 20:
        payin_category = "Payin Below 20%"
    elif 21 <= payin_value <= 30:
        payin_category = "Payin 21% to 30%"
    elif 31 <= payin_value <= 50:
        payin_category = "Payin 31% to 50%"
    else:
        payin_category = "Payin Above 50%"

    # Normalize segment for matching (ignore spaces)
    segment_normalized = segment.replace(' ', '')

    # Determine LOB from segment
    lob = ""
    if any(keyword.replace(' ', '') in segment_normalized for keyword in ['BUS', 'SCHOOL BUS', 'STAFF BUS', 'SCHOOL BUSES']):
        lob = "BUS"
    elif any(keyword.replace(' ', '') in segment_normalized for keyword in ['TW', 'MOTOR CYCLE', '2W', 'TWO WHEELER', 'BIKE', 'MC', 'SC', 'EV']):
        lob = "TW"
    elif any(keyword.replace(' ', '') in segment_normalized for keyword in ['PVT CAR', 'PRIVATE CAR', 'CAR']):
        lob = "PVT CAR"
    elif any(keyword.replace(' ', '') in segment_normalized for keyword in ['CV', 'COMMERCIAL', 'LCV', 'GVW', 'TN', 'PCV', 'GCV']):
        lob = "CV"
    
    elif any(keyword.replace(' ', '') in segment_normalized for keyword in ['TAXI', 'PVT TAXI']):
        lob = "TAXI"
    elif any(keyword.replace(' ', '') in segment_normalized for keyword in ['MISD', 'TRACTOR', 'MISC', 'CRANE', 'GARBAGE']):
        lob = "MISD"
    else:
        lob = "TW"

    # Extract policy type from segment if present
    segment_policy_type = None
    if 'COMP' in segment or 'SAOD' in segment or 'TP' in segment:
        segment_policy_type = next((pt for pt in ['COMP', 'SAOD', 'TP'] if pt in segment), None)
    else:
        segment_policy_type = policy_type

    # Find matching formula rule based on segment, insurer, and payin category
    matched_rule = None
    rule_explanation = ""
    
    # Normalize company name for matching (handle case variations)
    company_normalized = company_name.upper().replace('GENERAL', '').replace('INSURANCE', '').strip() if company_name and company_name != "Unknown" else "UNKNOWN"
    
    print(f"🔍 DEBUG: Company='{company_name}' -> Normalized='{company_normalized}', Segment='{segment}', LOB='{lob}', Payin={payin_value}")

    # First pass: Look for exact segment matches (including School Bus)
    candidate_rules = []
    
    for rule in FORMULA_DATA:
        print(f"🔍 Checking rule: LOB={rule['LOB']}, Segment={rule['SEGMENT']}, Insurer={rule['INSURER']}")
        
        if rule["LOB"] != lob:
            print(f"   ❌ LOB mismatch: {lob} != {rule['LOB']}")
            continue

        # Check segment match
        rule_segment = rule["SEGMENT"].upper().strip()
        segment_match = False
        
        if lob == "TW":
            if rule_segment == "1+5" and "1+5" in segment:
                segment_match = True
            elif rule_segment == "TW SAOD + COMP" and segment_policy_type in ["COMP", "SAOD", "COMP + SAOD"]:
                segment_match = True
            elif rule_segment == "TW TP" and segment_policy_type == "TP":
                segment_match = True
        elif lob == "PVT CAR":
            if rule_segment == "PVT CAR COMP + SAOD" and segment_policy_type in ["COMP", "SAOD", "COMP + SAOD"]:
                segment_match = True
            elif rule_segment == "PVT CAR TP" and segment_policy_type == "TP":
                segment_match = True
        elif lob == "CV":
            if rule_segment == "UPTO 2.5 GVW" and ("UPTO 2.5 GVW" in segment or segment == "CV"):
                segment_match = True
            elif rule_segment == "ALL GVW & PCV 3W, GCV 3W" and any(gvw in segment for gvw in ["GVW", "PCV 3W", "GCV 3W", "CV"]):
                segment_match = True
        elif lob == "BUS":
            if rule_segment == "SCHOOL BUS" and segment == "SCHOOL BUS":
                segment_match = True
                print(f"   ✓ SCHOOL BUS segment match found!")
            elif rule_segment == "STAFF BUS" and segment == "STAFF BUS":
                segment_match = True
        elif lob == "TAXI" and rule_segment == "TAXI" and segment == "TAXI":
            segment_match = True
        elif lob == "MISD" and rule_segment == "MISD, TRACTOR" and segment == "MISD":
            segment_match = True

        if not segment_match:
            print(f"   ❌ Segment mismatch: '{segment}' != '{rule_segment}'")
            continue

        # Check insurer match
        insurers = [ins.strip().upper() for ins in rule["INSURER"].split(',')]
        company_match = False
        insurer_specificity = 0
        
        print(f"   📋 Rule insurers: {insurers}")
        
        for insurer in insurers:
            clean_insurer = insurer.strip()
            if clean_insurer not in ["ALL COMPANIES", "REST OF COMPANIES"]:
                if (clean_insurer in company_normalized or 
                    company_normalized in clean_insurer or
                    any(part in company_normalized for part in clean_insurer.split()) or
                    any(part in clean_insurer for part in company_normalized.split())):
                    specific_match = True
                    insurer_specificity = 3
                    company_match = True
                    print(f"   ✓ Company match: '{company_normalized}' matches '{clean_insurer}'")
                    break
        
        if not company_match:
            if "ALL COMPANIES" in insurers:
                company_match = True
                insurer_specificity = 1
                print(f"   ✓ Company match: All Companies rule")
            elif "REST OF COMPANIES" in insurers:
                is_in_specific_list = False
                for other_rule in FORMULA_DATA:
                    if (other_rule["LOB"] == rule["LOB"] and 
                        other_rule["SEGMENT"] == rule["SEGMENT"] and
                        "REST OF COMPANIES" not in other_rule["INSURER"] and
                        "ALL COMPANIES" not in other_rule["INSURER"]):
                        other_insurers = [ins.strip().upper() for ins in other_rule["INSURER"].split(',')]
                        for other_insurer in other_insurers:
                            if (other_insurer in company_normalized or 
                                company_normalized in other_insurer or
                                any(part in company_normalized for part in other_insurer.split())):
                                is_in_specific_list = True
                                break
                        if is_in_specific_list:
                            break
                if not is_in_specific_list:
                    company_match = True
                    insurer_specificity = 2
                    print(f"   ✓ Company match: Rest of Companies rule")

        if not company_match:
            print(f"   ❌ Company mismatch: '{company_normalized}' not in {insurers}")
            continue

        rule_remarks = rule.get("REMARKS", "").upper().strip()
        remarks_specificity = 0
        
        if rule_remarks == "NIL":
            remarks_specificity = 10
            print(f"   ✓ NIL remarks - highest priority match")
        elif payin_value > 0:
            if rule_remarks == payin_category.upper():
                remarks_specificity = 8
                print(f"   ✓ Payin category match: {payin_category}")
            elif rule_remarks == "PAYIN ABOVE 20%" and payin_value > 20:
                remarks_specificity = 6
                print(f"   ✓ Payin above 20% fallback match")
            else:
                if payin_value > 0:
                    continue
        else:
            if rule_remarks == "NIL":
                remarks_specificity = 10
            else:
                continue

        total_score = (insurer_specificity * 20) + remarks_specificity
        candidate_rules.append((rule, total_score, insurer_specificity, remarks_specificity))
        print(f"   ✓ Candidate rule added - Score: {total_score} (Insurer: {insurer_specificity}, Remarks: {remarks_specificity})")

    if candidate_rules:
        candidate_rules.sort(key=lambda x: x[1], reverse=True)
        matched_rule = candidate_rules[0][0]
        # explanation = f"Match: LOB={lob}, Segment={segment}, Insurer={matched_rule['INSURER']}, Remarks={matched_rule.get('REMARKS', 'NIL')}, Formula={matched_rule['PO']}"
    else:
        matched_rule = {"PO": "90% of Payin", "INSURER": "Default", "REMARKS": "NIL"}
        # explanation = f"No specific rule matched - Using fallback: 90% of Payin"
        # print(f"   ⚠️ No rules matched - using default 90% of Payin")
# 

    po_formula = matched_rule["PO"].strip()
    calculated_payout = 0.0
    
    print(f"   💰 Formula: '{po_formula}', Payin: {payin_value}")
    
    if "21% flat" in po_formula or "21% FLAT" in po_formula:
        calculated_payout = 21.0
        op_str = "= 21.00% (flat)"
    elif "90% of Payin" in po_formula or "90% OF PAYIN" in po_formula:
        calculated_payout = payin_value * 0.9
        op_str = f"* 0.9 = {calculated_payout:.2f}%"
    elif "88% of Payin" in po_formula or "88% OF PAYIN" in po_formula:
        calculated_payout = payin_value * 0.88
        op_str = f"* 0.88 = {calculated_payout:.2f}%"
    elif "Less 2% of Payin" in po_formula or "LESS 2% OF PAYIN" in po_formula or "-2%" in po_formula:
        calculated_payout = payin_value - 2
        op_str = f"- 2% = {calculated_payout:.2f}%"
    elif "-3%" in po_formula:
        calculated_payout = payin_value - 3
        op_str = f"- 3% = {calculated_payout:.2f}%"
    elif "-4%" in po_formula:
        calculated_payout = payin_value - 4
        op_str = f"- 4% = {calculated_payout:.2f}%"
    elif "-5%" in po_formula:
        calculated_payout = payin_value - 5
        op_str = f"- 5% = {calculated_payout:.2f}%"
    else:
        calculated_payout = payin_value
        op_str = f"(using original payin)"
    
    calculated_payout = max(0, calculated_payout)
    rule_explanation = f"Match: LOB={lob}, Segment={segment}, Insurer={matched_rule['INSURER']}, Remarks={matched_rule.get('REMARKS', 'NIL')}, Formula={matched_rule['PO']}; Calculated: {payin_value:.2f}% {op_str}"
    print(f"   ✅ Final Payout: {calculated_payout:.2f}%")

    return f"{calculated_payout:.2f}%", rule_explanation

def extract_text_from_sheet(sheet_data, sheet_name):
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

# def parse_converted_json(json_path_or_data, original_excel_path=None):
#     """Parse converted JSON"""
#     print("\n" + "="*60)
#     print("PARSING CONVERTED JSON")
#     print("="*60)
    
#     if isinstance(json_path_or_data, str):
#         with open(json_path_or_data, 'r', encoding='utf-8') as f:
#             data = json.load(f)
#         json_path = json_path_or_data
#     else:
#         data = json_path_or_data
#         json_path = None
    
#     source_file = original_excel_path or data.get("source_file") or json_path
#     filename_company, filename_location = extract_from_filename(source_file) if source_file else (None, None)
    
#     print(f"\n📁 Source: {source_file}")
#     print(f"📁 Detected Company: {filename_company}")
#     print(f"📍 Detected Location: {filename_location}")
    
#     all_entries = []
    
#     if "sheets" in data:
#         print(f"\n📊 Processing {data['total_sheets']} sheet(s)...")
        
#         for sheet_name, sheet_data in data["sheets"].items():
#             print(f"\n🔄 Processing sheet: '{sheet_name}'")
#             print(f"   Records in sheet: {len(sheet_data)}")
            
#             sheet_text = extract_text_from_sheet(sheet_data, sheet_name)
#             extracted_list = extract_structured_data_single(sheet_text, "Excel")
#             for extracted in extracted_list:
#                 normalized = normalize_extracted_data(extracted, source_file)
#                 if filename_company:
#                     normalized["COMPANY NAME"] = filename_company
#                 if filename_location:
#                     normalized["LOCATION"] = filename_location or normalized["LOCATION"]
#                 payout, explanation = calculate_payout(normalized)
#                 normalized["PAYOUT"] = payout
#                 normalized["CALCULATION EXPLANATION"] = explanation
#                 all_entries.append(normalized)
            
#             print(f"   ✓ Extracted {len(extracted_list)} entries from sheet '{sheet_name}'")
    
#     print(f"\n✅ Total entries extracted: {len(all_entries)}")
    
#     return {
#         "file_type": "JSON (Converted from Excel)",
#         "file_path": json_path or source_file,
#         "structured_data": all_entries
#     }
def parse_converted_json(json_path_or_data, original_excel_path=None):
    """
    Parse converted JSON, applying pre-processing to handle merged cells
    (fill-down logic) before sending to AI.
    """
    print("\n" + "="*60)
    print("PARSING CONVERTED JSON (with Fill-Down Logic)")
    print("="*60)
    
    if isinstance(json_path_or_data, str):
        with open(json_path_or_data, 'r', encoding='utf-8') as f:
            data = json.load(f)
        json_path = json_path_or_data
    else:
        data = json_path_or_data
        json_path = None
    
    source_file = original_excel_path or data.get("source_file") or json_path
    # filename_company, filename_location = extract_from_filename(source_file) if source_file else (None, None)
    
    # Force Zuno as company, ignore filename detection
    filename_company = "Zuno"
    filename_location = extract_from_filename(source_file)[1] if source_file else None


    print(f"\n📁 Source: {source_file}")
    print(f"📁 Detected Company: {filename_company}")
    print(f"📍 Detected Location: {filename_location}")
    
    all_entries = []
    
    if "sheets" in data:
        print(f"\n📊 Processing {data['total_sheets']} sheet(s)...")
        
        for sheet_name, sheet_data in data["sheets"].items():
            print(f"\n🔄 Processing sheet: '{sheet_name}'")
            print(f"   Records in sheet: {len(sheet_data)}")

            # --- START: Fill-Down Logic ---
            
            # 1. Find header row and key columns (Segment, Location, Payin)
            header_row_index = -1
            header_map = {}
            data_start_index = 0

            for i, row in enumerate(sheet_data):
                row_vals = [str(v).lower().strip() for v in row.values() if v]
                # Look for keywords that indicate a header row
                if any(k in row_vals for k in ["product line", "segment", "region/rto/zone", "with ncb", "payin"]):
                    header_row_index = i
                    data_start_index = i + 1
                    for col_key, val in row.items():
                        if val:
                            val_str = str(val).lower().strip()
                            if "product line" in val_str or "segment" in val_str:
                                header_map['segment'] = col_key
                            elif "region" in val_str or "location" in val_str or "rto" in val_str:
                                header_map['location'] = col_key
                            elif "payin" in val_str or "ncb" in val_str:
                                header_map['payin'] = col_key
                    print(f"   Found header row at index {header_row_index}. Key columns: {header_map}")
                    break # Found header
            
            # 2. If no header found, make educated guesses
            if header_row_index == -1 and sheet_data:
                print("   No header row found. Guessing columns...")
                first_row_vals = [str(v).lower() for v in sheet_data[0].values()]
                if "imd code" in first_row_vals:
                    # This is likely the Zuno sheet structure
                    header_map = {'segment': 'col_1', 'location': 'col_2', 'payin': 'col_3'}
                    data_start_index = 1 # Skip the "header"
                    print(f"   Guessed Zuno structure (IMD Code): {header_map}")
                else:
                    # Default guess
                    header_map = {'segment': 'col_0', 'location': 'col_1', 'payin': 'col_2'}
                    print(f"   Guessed default structure: {header_map}")

            # 3. Apply fill-down logic for the identified 'segment' column
            processed_sheet_data = []
            last_segment = None
            segment_col_key = header_map.get('segment')

            if segment_col_key:
                print(f"   Applying fill-down logic on column: {segment_col_key}")
                for i in range(data_start_index, len(sheet_data)):
                    row = sheet_data[i]
                    # Check if row is mostly empty
                    if all(v is None or str(v).strip() == "" for v in row.values()):
                        continue # Skip empty rows

                    current_segment = row.get(segment_col_key)
                    
                    if current_segment is not None and str(current_segment).strip():
                        last_segment = current_segment
                    elif last_segment is not None:
                        # This is a 'merged' row. Fill it.
                        row[segment_col_key] = last_segment
                    
                    processed_sheet_data.append(row)
                print(f"   Fill-down complete. Processed {len(processed_sheet_data)} data rows.")
            else:
                print("   ⚠️ Warning: Could not find segment column. Proceeding without fill-down.")
                processed_sheet_data = sheet_data[data_start_index:]
            
            # --- END: Fill-Down Logic ---

            # Now, convert the *processed* data to text for the AI
            sheet_text = extract_text_from_sheet(processed_sheet_data, sheet_name)
            
            extracted_list = extract_structured_data_single(sheet_text, "Excel")
            
            for extracted in extracted_list:
                normalized = normalize_extracted_data(extracted, source_file)
                if filename_company:
                    normalized["COMPANY NAME"] = filename_company
                if filename_location:
                    normalized["LOCATION"] = filename_location or normalized["LOCATION"]
                
                payout, explanation = calculate_payout(normalized)
                normalized["PAYOUT"] = payout
                normalized["CALCULATION EXPLANATION"] = explanation
                all_entries.append(normalized)
            
            print(f"   ✓ Extracted {len(extracted_list)} entries from sheet '{sheet_name}'")
    
    print(f"\n✅ Total entries extracted: {len(all_entries)}")
    
    return {
        "file_type": "JSON (Converted from Excel)",
        "file_path": json_path or source_file,
        "structured_data": all_entries
    }


def normalize_extracted_data(extracted, file_path):
    """Normalize extracted data from AI"""
    normalized = {
        "COMPANY NAME": None,
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
    
    if extracted.get("segment"):
        normalized["SEGMENT"] = normalize_segment(extracted["segment"])
    else:
        normalized["SEGMENT"] = "Unknown"
    
    if extracted.get("policy_type"):
        normalized["POLICY TYPE"] = normalize_policy_type(extracted["policy_type"])
    else:
        normalized["POLICY TYPE"] = "Comprehensive"
    
    if filename_location:
        normalized["LOCATION"] = filename_location
    elif extracted.get("location"):
        normalized["LOCATION"] = extracted["location"]
    
    if extracted.get("payin"):
        payin_str = str(extracted["payin"])
        payin_match = re.search(r'(\d+(?:\.\d+)?)', payin_str)
        if payin_match:
            normalized["PAYIN"] = float(payin_match.group(1))
    
    if extracted.get("other_info"):
        other_info_str = str(extracted["other_info"]).strip()
        if other_info_str:
            normalized["REMARK"].append(other_info_str)
    
    return normalized

def extract_structured_data_single(text, file_type):
    """Extract structured data using OpenAI"""
    prompt = f"""You are an expert at extracting structured data from insurance payout grids. These grids can be in tabular format, lists, or descriptive text, and may have dynamic columns/rows. Your goal is to parse the entire text intelligently, identifying all entries even if the structure is irregular, merged, or pivoted.

Key Guidelines:
- Handle dynamic structures: Columns might be locations, payins, segments, or mixed. Rows might represent segments, with sub-details for policy types, ages, etc.
- Identify company: From text or infer as 'Unknown'. Common: ICICI, Reliance, Chola, Sompo, Kotak, Magma, Bajaj, Digit, Liberty, Future, Tata, IFFCO, Royal, SBI, Zuno, HDFC, Shriram.
- Segment normalization: Map to standard names like 'PVT CAR COMP + SAOD', 'TW TP', 'GCV 3W', 'SCHOOL BUS', 'MISD'. Detect from context, e.g., 'Pvt Car Package' -> 'PVT CAR COMP + SAOD'.
- Policy type: 'Package' or 'Comp' -> 'Comprehensive', 'TP' -> 'TP', 'SAOD' -> 'SAOD'. If not specified, infer from segment.
- Location: Cities, states, RTO codes, or 'Pan India'. Handle lists like 'North only Delhi NCR, Chandigarh'. Please know APTS is a location, so keep it in the location column in output file if u find it and calculate payout based on the respective row.
- Payin: Extract percentages or numbers (e.g., 0.25 -> 25%). Ignore non-numeric.
- Other info/Remarks: Capture age (e.g., '1-6 yr'), fuel, notes (e.g., 'New/Rollover'), logic.
- Handle pivoted tables: If rows are segments and columns are locations with payins, create an entry per cell.
- Handle multiple entries per row: If a row has multiple segments or policy types, split into separate objects.
- Ignore empty/header rows. Output only valid data rows.
- Be robust: If structure is unclear, reason step-by-step in your mind to map to fields.

Example Input:
row1: Segment,Location,Payin
row2: Pvt Car Package (New),Gujarat,0.28
row3: TW TP,Pan India,0.20

Example Output:
[
  {{"company_name": "Zuno", "segment": "PVT CAR COMP + SAOD", "policy_type": "Comprehensive", "location": "Gujarat", "payin": "28", "other_info": "New"}},
  {{"company_name": "Zuno", "segment": "TW TP", "policy_type": "TP", "location": "Pan India", "payin": "20", "other_info": null}}
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
        print(f"Raw API response: {response_text}")
        result = json.loads(response_text)
        if not isinstance(result, list):
            result = [result]
    except json.JSONDecodeError as e:
        print(f"\n⚠️ JSON Parse Error: {e} - Using fallback empty entry")
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
            output_file = "z_grid_output.xlsx"
            output_df.to_excel(output_file, index=False)
            
            print(f"\n📊 Excel file created: {output_file}")
            print(f"   Total entries: {len(entries)}")
            print("\n✅ Summary:")
            print(f"   - Company: {output_data['COMPANY NAME'][0] if output_data['COMPANY NAME'] else 'N/A'}")
            print(f"   - Unique Segments: {len(set(s for s in output_data['SEGMENT'] if s))}")
            print(f"   - Unique Locations: {len(set(l for l in output_data['LOCATION'] if l))}")
            
            if any(output_data['PAYIN']):
                valid_payins = [p for p in output_data['PAYIN'] if p]
                if valid_payins:
                    print(f"   - Payin Range: {min(valid_payins):.1f}% - {max(valid_payins):.1f}%")
            
            print("\n📋 Sample Entries:")
            for i in range(min(5, len(entries))):
                entry = entries[i]
                print(f"   {i+1}. {entry['SEGMENT']} | {entry.get('POLICY TYPE', 'N/A')} | {entry['PAYIN']}% → {entry['PAYOUT']} | Explanation: {entry['CALCULATION EXPLANATION']}")
        else:
            print("\n⚠️  No entries extracted from file")

if __name__ == "__main__":
    import sys
    
    # Get file path from command line or prompt user
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("📁 Enter file path (Excel/PDF/Image): ").strip()
    
    # Parse the file
    result = parse_file(file_path)
    
    if "error" in result:
        print(f"\n❌ Error: {result['error']}")
        sys.exit(1)
    
    # Extract entries and create output
    entries = result.get('structured_data', [])
    
    if not entries:
        print("\n⚠️  No entries extracted from file")
        sys.exit(0)
    
    # Create DataFrame
    output_data = {
        "COMPANY NAME": [e.get("COMPANY NAME", "") for e in entries],
        "SEGMENT": [e.get("SEGMENT", "") for e in entries],
        "POLICY TYPE": [e.get("POLICY TYPE", "") for e in entries],
        "LOCATION": [e.get("LOCATION", "") for e in entries],
        "PAYIN": [e.get("PAYIN", "") for e in entries],
        "PAYOUT": [e.get("PAYOUT", "") for e in entries],
        "REMARK": ["; ".join(e.get("REMARK", [])) for e in entries],
        "CALCULATION EXPLANATION": [e.get("CALCULATION EXPLANATION", "") for e in entries]
    }
    
    output_df = pd.DataFrame(output_data)
    output_file = "z_grid_output.xlsx"
    output_df.to_excel(output_file, index=False)
    
    print(f"\n✅ Done! Created: {output_file} ({len(entries)} entries)")