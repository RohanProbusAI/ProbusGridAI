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
    {"LOB": "TW", "SEGMENT": "1+5", "INSURER": "ROYAL", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "INSURER": "ROYAL", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "INSURER": "ROYAL", "PO": "90% of Payin", "REMARKS": "All Fuel"},
    {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "ROYAL", "PO": "90% of Payin", "REMARKS": "NIL"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "Upto 2.5 GVW", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "PCV 3W", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "GCV 3W", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW, PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "CV", "SEGMENT": "All GVW & PCV 3W, GCV 3W", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "INSURER": "ROYAL", "PO": "Less 2% of Payin", "REMARKS": "NIL"},
    {"LOB": "BUS", "SEGMENT": "STAFF BUS", "INSURER": "ROYAL", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "ROYAL", "PO": "-2%", "REMARKS": "Payin Below 20%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "ROYAL", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "ROYAL", "PO": "-4%", "REMARKS": "Payin 31% to 50%"},
    {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "ROYAL", "PO": "-5%", "REMARKS": "Payin Above 50%"},
    {"LOB": "MISD", "SEGMENT": "MISD", "INSURER": "ROYAL", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "TRACTOR", "INSURER": "ROYAL", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "MISC", "INSURER": "ROYAL", "PO": "88% of Payin", "REMARKS": "NIL"},
    {"LOB": "MISD", "SEGMENT": "Misc, Misd, Tractor", "INSURER": "ROYAL", "PO": "88% of Payin", "REMARKS": "NIL"}
]

COMPANY_MAPPING = {
    'icici': 'ICICI', 'reliance': 'Reliance', 'chola': 'Chola', 'sompo': 'Sompo',
    'kotak': 'Kotak', 'magma': 'Magma', 'bajaj': 'Bajaj', 'digit': 'Digit', 
    'liberty': 'Liberty', 'future': 'Future', 'tata': 'Tata', 'iffco': 'IFFCO', 
    'royal': 'Royal', 'sbi': 'SBI', 'zuno': 'Zuno', 'hdfc': 'HDFC', 'shriram': 'Shriram'
}

KNOWN_LOCATIONS = [
    'ahmedabad', 'surat', 'vadodara', 'rajkot', 'mumbai', 'pune', 'nagpur', 'thane',
    'delhi', 'bangalore', 'bengaluru', 'chennai', 'hyderabad', 'kolkata', 'jaipur',
    'lucknow', 'kanpur', 'patna', 'indore', 'bhopal', 'chandigarh', 'coimbatore',
    'kochi', 'guwahati', 'bhubaneswar', 'visakhapatnam', 'goa', 'mysore', 'nashik',
    'guj', 'gujarat', 'mh', 'maharashtra', 'apts', 'andhra pradesh', 'karnataka',
    'tamil nadu', 'tn', 'assam', 'tripura', 'meghalaya', 'mizoram', 'west bengal',
    'wb', 'up', 'uttar pradesh', 'mp', 'madhya pradesh', 'raj', 'rajasthan',
    'bihar', 'jharkhand', 'kerala', 'odisha', 'chhattisgarh', 'uttarakhand', 
    'srinagar', 'cochin', 'jammu', 'hubli', 'bhubaneshwar', 'coimbatore', 'madurai',
    'tiruchirappalli', 'salem', 'telangana', 'jammu & kashmir', 'rood', 'roe', 
    'haryana', 'vijaywada', 'telengana', 'bhagalpur', 'gaya', 'muzaffarpur', 'patna',
    'purnia', 'kolkata', 'rest of bihar', 'rest of west bengal'
]

RTO_CODE_PATTERN = r'^[A-Z]{2}\d{1,2}$|^AS\d{1,2}$|^[A-Z]{2,4}\d{1,2}$'


def extract_from_filename(file_path):
    """Extract company name and location from filename"""
    filename = str(file_path).lower()
    base_name = Path(file_path).stem.lower()

    for key, value in COMPANY_MAPPING.items():
        if key in filename:
            return value, _extract_location(filename)

    words = re.findall(r'\b[a-zA-Z]{3,}\b', base_name)
    potential_companies = [w.title() for w in words if w.title() not in 
                          {'Insurance', 'General', 'Grid', 'Payout', 'Payin', 'Excel', 'Pdf', 'Img'}]
    if potential_companies:
        company = potential_companies[0]
        return company, _extract_location(filename)

    return None, _extract_location(filename)


def _extract_location(filename):
    """Helper to extract location from filename"""
    location_patterns = [
        r'\b(guj|gujarat)\b', r'\b(mh|maharashtra|mumbai)\b', 
        r'\b(apts)\b', r'\b(dl|delhi)\b', 
        r'\b(kar|karnataka|bangalore|bengaluru)\b',
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
    
    if re.match(RTO_CODE_PATTERN, str(column_name).strip(), re.IGNORECASE):
        return True
    
    if any(loc in col_lower for loc in KNOWN_LOCATIONS):
        return True
    
    location_terms = ['location', 'city', 'state', 'region', 'rto', 'zone', 'area', 
                     'pradesh', 'division']
    if any(term in col_lower for term in location_terms):
        return True
    
    return False


def detect_segment_from_sheet_name(sheet_name: str) -> Optional[str]:
    """Detect segment from sheet name"""
    sheet_lower = sheet_name.lower()
    
    # CV segments
    if 'gcv' in sheet_lower or 'cv' in sheet_lower:
        if '4 wheeler' in sheet_lower or '4w' in sheet_lower:
            return "All GVW & PCV 3W, GCV 3W"
        return "All GVW & PCV 3W, GCV 3W"
    
    # PCV segments
    if 'pcv' in sheet_lower:
        if '3w' in sheet_lower or '3 wheeler' in sheet_lower:
            return "PCV 3W"
        return "All GVW & PCV 3W, GCV 3W"
    
    # Two wheeler
    if 'tw' in sheet_lower or '2w' in sheet_lower or 'two wheeler' in sheet_lower:
        if 'tp' in sheet_lower:
            return "TW TP"
        elif 'comp' in sheet_lower or 'saod' in sheet_lower:
            return "TW SAOD + COMP"
        return "TW"
    
    # Private car
    if 'pvt car' in sheet_lower or 'private car' in sheet_lower or 'car' in sheet_lower:
        if 'tp' in sheet_lower:
            return "PVT CAR TP"
        return "PVT CAR COMP + SAOD"
    
    # Bus
    if 'bus' in sheet_lower:
        if 'school' in sheet_lower:
            return "SCHOOL BUS"
        elif 'staff' in sheet_lower:
            return "STAFF BUS"
        return "SCHOOL BUS"
    
    # Taxi
    if 'taxi' in sheet_lower:
        return "TAXI"
    
    # 3 Wheeler
    if '3w' in sheet_lower or '3 wheeler' in sheet_lower:
        if 'gcv' in sheet_lower:
            return "GCV 3W"
        return "PCV 3W"
    
    # Misc
    if 'misc' in sheet_lower or 'misd' in sheet_lower or 'tractor' in sheet_lower:
        return "MISD"
    
    return None


def parse_royal_column_header(header_text: str) -> Dict:
    """
    Parse Royal company column headers to extract segment details
    Enhanced to better detect weight ranges and policy types
    """
    header_lower = str(header_text).lower()
    
    result = {
        'weight_range': None,
        'conditions': header_text,
        'policy_type': 'Comprehensive',
        'is_segment_column': False
    }
    
    # Enhanced weight/tonnage range patterns
    weight_patterns = [
        r'(\d+\.?\d*\s*to\s*\d+\.?\d*)',  # Matches: 12 to 20, 3.5 to 7.5, etc.
        r'(\d+\.?\d*\s*-\s*\d+\.?\d*)',    # Matches: 12-20, 3.5-7.5, etc.
        r'(upto\s*\d+\.?\d*)',             # Matches: upto 2.5, upto 85%, etc.
        r'(above\s*\d+\.?\d*)',            # Matches: above 45
        r'(>\s*\d+\.?\d*)',                # Matches: >45
        r'(\d+\.?\d*t)',                   # Matches: 45t, 12t, etc.
    ]
    
    # Check if this column header contains a weight range
    for pattern in weight_patterns:
        match = re.search(pattern, header_lower)
        if match:
            result['weight_range'] = match.group(1).strip()
            result['is_segment_column'] = True
            break
    
    # Additional check: if header contains keywords like "disc", "dep", "od", "tp"
    # along with numbers, it's likely a segment column
    if not result['is_segment_column']:
        segment_keywords = ['disc', 'dep', 'od', 'tp', 'satp', 'comp', 'saod', 'nil dep', 
                           'package', 'tata', 'al', 'wagon', 'other than']
        has_keywords = any(kw in header_lower for kw in segment_keywords)
        has_numbers = bool(re.search(r'\d+', header_lower))
        
        if has_keywords and has_numbers:
            result['is_segment_column'] = True
    
    # Detect policy type from conditions
    if 'tp' in header_lower and 'satp' not in header_lower:
        result['policy_type'] = 'TP'
    elif 'saod' in header_lower or 'od disc' in header_lower:
        result['policy_type'] = 'SAOD'
    elif 'comp' in header_lower or 'comprehensive' in header_lower or 'package' in header_lower:
        result['policy_type'] = 'Comprehensive'
    else:
        result['policy_type'] = 'Comprehensive + saod'
    
    return result


def detect_royal_structure(sheet_data, sheet_name) -> Dict:
    """
    Enhanced Royal structure detection with better range recognition
    """
    print(f"\n🔍 Detecting Royal structure for sheet: {sheet_name}")
    
    structure = {
        'is_royal_format': False,
        'header_row_index': -1,
        'state_col': None,
        'rto_division_col': None,
        'segment_columns': [],
        'data_start_index': 0,
        'detected_segment': None
    }
    
    # Detect segment from sheet name
    detected_segment = detect_segment_from_sheet_name(sheet_name)
    if detected_segment:
        structure['detected_segment'] = detected_segment
        print(f"   ✓ Detected segment from sheet name: {detected_segment}")
    
    # Find header row
    for i, row in enumerate(sheet_data):
        row_vals = {k: str(v).lower().strip() if v else "" for k, v in row.items()}
        
        has_state = any('state' in val for val in row_vals.values())
        has_rto = any('rto' in val and 'division' in val for val in row_vals.values())
        
        if has_state or has_rto:
            structure['header_row_index'] = i
            structure['is_royal_format'] = True
            
            for col_key, val in row.items():
                if not val:
                    continue
                val_lower = str(val).lower().strip()
                
                if val_lower == 'state':
                    structure['state_col'] = col_key
                elif 'rto' in val_lower and 'division' in val_lower:
                    structure['rto_division_col'] = col_key
                elif val_lower not in ['state', '']:
                    # Parse header to check if it's a segment column
                    parsed_header = parse_royal_column_header(str(val))
                    
                    # Include column if it's identified as a segment column OR
                    # if it's not a location column and not state/rto
                    if parsed_header['is_segment_column'] or not is_location_column(val):
                        structure['segment_columns'].append({
                            'col_key': col_key,
                            'header_text': str(val),
                            'parsed': parsed_header
                        })
            
            structure['data_start_index'] = i + 1
            break
    
    # Fallback detection by column patterns
    if not structure['is_royal_format'] and sheet_data:
        first_row = sheet_data[0]
        cols = list(first_row.keys())
        
        if len(cols) >= 3:
            sample_col0 = [sheet_data[i].get(cols[0]) for i in range(min(5, len(sheet_data)))]
            sample_col1 = [sheet_data[i].get(cols[1]) for i in range(min(5, len(sheet_data)))]
            
            col0_looks_like_state = any(is_location_column(str(v)) for v in sample_col0 if v)
            col1_looks_like_rto = any(re.match(r'^[A-Z\s]+$', str(v)) for v in sample_col1 if v)
            
            if col0_looks_like_state and col1_looks_like_rto:
                structure['is_royal_format'] = True
                structure['header_row_index'] = 0
                structure['state_col'] = cols[0]
                structure['rto_division_col'] = cols[1]
                structure['data_start_index'] = 1
                
                # Process remaining columns as potential segment columns
                for col_key in cols[2:]:
                    header_text = first_row.get(col_key, col_key)
                    parsed_header = parse_royal_column_header(str(header_text))
                    
                    # Only add if it looks like a segment column
                    if parsed_header['is_segment_column']:
                        structure['segment_columns'].append({
                            'col_key': col_key,
                            'header_text': str(header_text),
                            'parsed': parsed_header
                        })
    
    print(f"   Royal format detected: {structure['is_royal_format']}")
    if structure['is_royal_format']:
        print(f"   State column: {structure['state_col']}")
        print(f"   RTO Division column: {structure['rto_division_col']}")
        print(f"   Segment columns found: {len(structure['segment_columns'])}")
        if structure['segment_columns']:
            print(f"   Segment column headers:")
            for seg in structure['segment_columns'][:5]:  # Show first 5
                print(f"     - {seg['header_text']}")
            if len(structure['segment_columns']) > 5:
                print(f"     ... and {len(structure['segment_columns']) - 5} more")
    
    return structure


def is_header_row(row, state_col, rto_col):
    """
    Check if a row is a header row (column names repeating in the middle of data)
    """
    if not row:
        return False
    
    # Check if state column contains "state" keyword
    state_val = row.get(state_col) if state_col else None
    if state_val and str(state_val).lower().strip() in ['state', 'states']:
        return True
    
    # Check if RTO column contains "rto" or "division" keywords
    rto_val = row.get(rto_col) if rto_col else None
    if rto_val:
        rto_val_lower = str(rto_val).lower().strip()
        if 'rto' in rto_val_lower and 'division' in rto_val_lower:
            return True
    
    # Check if multiple cells contain header-like terms
    header_keywords = ['state', 'rto', 'division', 'location', 'segment', 'payin', 'payout']
    header_count = 0
    for val in row.values():
        if val:
            val_lower = str(val).lower().strip()
            if any(keyword in val_lower for keyword in header_keywords):
                header_count += 1
    
    # If 3 or more cells contain header keywords, it's likely a header row
    if header_count >= 3:
        return True
    
    return False


def is_header_row_generic(row, segment_col, policy_col):
    """
    Check if a row is a header row for generic pivoted tables
    """
    if not row:
        return False
    
    # Check if segment column contains "segment" keyword
    segment_val = row.get(segment_col) if segment_col else None
    if segment_val:
        segment_val_lower = str(segment_val).lower().strip()
        if segment_val_lower in ['segment', 'segments', 'product', 'product line']:
            return True
    
    # Check if policy column contains "policy" keyword
    policy_val = row.get(policy_col) if policy_col else None
    if policy_val:
        policy_val_lower = str(policy_val).lower().strip()
        if 'policy' in policy_val_lower and 'type' in policy_val_lower:
            return True
    
    # Check if multiple cells contain header-like terms
    header_keywords = ['segment', 'policy', 'location', 'payin', 'payout', 'fuel', 'age']
    header_count = 0
    for val in row.values():
        if val:
            val_lower = str(val).lower().strip()
            if any(keyword in val_lower for keyword in header_keywords):
                header_count += 1
    
    # If 3 or more cells contain header keywords, it's likely a header row
    if header_count >= 3:
        return True
    
    return False


def restructure_royal_data(sheet_data, structure, sheet_name):
    """
    Enhanced restructuring with better range detection and header row detection
    """
    print(f"\n📊 Restructuring Royal data for sheet: {sheet_name}")
    
    restructured_data = []
    state_col = structure['state_col']
    rto_col = structure['rto_division_col']
    segment_columns = structure['segment_columns']
    data_start_index = structure['data_start_index']
    base_segment = structure['detected_segment']
    
    # Validate essential columns
    if not state_col and not rto_col:
        print(f"   ⚠️  Warning: No state or RTO column found. Skipping sheet.")
        return []
    
    if not segment_columns:
        print(f"   ⚠️  Warning: No segment columns found. Skipping sheet.")
        return []
    
    if not base_segment:
        print(f"\n⚠️  Could not auto-detect segment from sheet name: '{sheet_name}'")
        print("   Please enter the segment type for this sheet:")
        print("   Examples: GCV, PCV, TW, PVT CAR, TAXI, BUS, MISD, etc.")
        user_segment = input("   Segment: ").strip()
        
        if user_segment:
            base_segment = normalize_segment(user_segment)
            print(f"   ✓ Using segment: {base_segment}")
        else:
            print("   ⚠️  No segment entered. Using 'All GVW & PCV 3W, GCV 3W' as default")
            base_segment = "All GVW & PCV 3W, GCV 3W"
    
    last_state = None
    skipped_header_rows = 0
    
    for i in range(data_start_index, len(sheet_data)):
        row = sheet_data[i]
        
        # Skip empty rows
        if all(v is None or str(v).strip() == "" for v in row.values()):
            continue
        
        # Skip header rows that appear in the middle of data (continuation headers)
        if is_header_row(row, state_col, rto_col):
            skipped_header_rows += 1
            print(f"   ⚠️  Skipping continuation header at row {i + 1}")
            continue
        
        current_state = row.get(state_col) if state_col else None
        if current_state and str(current_state).strip():
            last_state = str(current_state).strip()
        elif last_state:
            current_state = last_state
        else:
            continue
        
        rto_division = row.get(rto_col) if rto_col else None
        rto_division_str = str(rto_division).strip() if rto_division else ""
        
        location = current_state
        if rto_division_str:
            location = f"{current_state} - {rto_division_str}"
        
        for seg_col_info in segment_columns:
            col_key = seg_col_info['col_key']
            header_info = seg_col_info['parsed']
            header_text = seg_col_info['header_text']
            
            payin_raw = row.get(col_key)
            payin_str = str(payin_raw).strip() if payin_raw is not None else ""
            
            # Try to extract weight range from header first
            segment = extract_weight_range_from_header(header_text)
            
            # If no range found in header, use base segment with range from parsed info
            if not segment:
                segment = base_segment
                if header_info.get('weight_range'):
                    weight_range = header_info['weight_range'].upper()
                    
                    if '0 TO 2' in weight_range or '0 TO 2.5' in weight_range or 'UPTO 2.5' in weight_range:
                        segment = "Upto 2.5 GVW"
                    elif '2 TO 3.5' in weight_range or '2.5 TO 3.5' in weight_range:
                        segment = "2.5-3.5 GVW"
                    elif '3.5 TO 7.5' in weight_range:
                        segment = "3.5-7.5 GVW"
                    elif '7.5 TO 12' in weight_range:
                        segment = "7.5-12 GVW"
                    elif '12 TO 20' in weight_range:
                        segment = "12-20 GVW"
                    elif '20 TO 40' in weight_range:
                        segment = "20-40 GVW"
                    elif '40 TO 45' in weight_range:
                        segment = "40-45 GVW"
                    elif 'ABOVE 45' in weight_range or '>45' in weight_range or '45T' in weight_range:
                        segment = ">45 GVW"
            
            entry = {
                'segment'     : segment,
                'policy_type' : header_info['policy_type'],
                'location'    : location,
                'payin'       : payin_str,
                'other_info'  : header_info['conditions']
            }
            
            restructured_data.append(entry)
    
    if skipped_header_rows > 0:
        print(f"   ⚠️  Skipped {skipped_header_rows} continuation header row(s)")
    print(f"   ✓ Created {len(restructured_data)} entries")
    return restructured_data


def normalize_segment(segment_str):
    """Normalize segment names to standard format"""
    if not segment_str or pd.isna(segment_str):
        return "Unknown"
    
    segment_lower = re.sub(r'\s+', ' ', str(segment_str).lower().strip())

    if segment_lower == "apts":
        return segment_str
    
    if segment_lower == "misc":
        return "MISD"

    # Bus
    if 'school buses' in segment_lower or 'school bus' in segment_lower:
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
        elif '1+5' in segment_lower or 'grid' in segment_lower or 'new' in segment_lower:
            return "1+5"
        return "TW"
    
    # Private Car
    if any(x in segment_lower for x in ['pvt car', 'private car', 'car', 'pvt car tp']):
        if 'pvt car satp' in segment_lower:
            return 'PVT CAR TP'
        elif 'comp' in segment_lower or 'saod' in segment_lower or 'package' in segment_lower:
            return "PVT CAR COMP + SAOD"
        elif 'tp' in segment_lower:
            return "PVT CAR TP"
        return "PVT CAR COMP + SAOD"

    # Commercial Vehicle with weight ranges - THIS IS THE KEY CHANGE
    if any(x in segment_lower for x in ['cv', 'gcv', 'pcv', 'lcv', 'gvw', 'tn', 'tonnage', 'pcv bus', 'pcv school', 'pcv 4w']):
        # Check for "Upto 2.5 GVW" - this is the ONLY one we keep as-is
        if 'upto 2.5' in segment_lower or '0-2.5' in segment_lower or '0 to 2.5' in segment_lower or '0 to 2' in segment_lower:
            return "Upto 2.5 GVW"
        
        # ALL other weight ranges (2.5+) map to "All GVW & PCV 3W, GCV 3W"
        # This includes: 2.5-3.5, 3.5-7.5, 7.5-12, 12-20, 20-40, 40-45, >45, etc.
        if any(x in segment_lower for x in ['2.5-3.5', '3.5-7.5', '7.5-12', '12-20', '20-40', '40-45', '>45', 'above 45', '3.5-12', '12-45', 'gvw', 'tn', 'tonnage']):
            return "All GVW & PCV 3W, GCV 3W"
        
        # Default for CV-related segments
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
    """
    Calculate payout only when a numeric PAYIN exists.
    If PAYIN is None (non-numeric), return empty payout and explanation while preserving PAYIN_RAW.
    """
    company_name = (normalized_data.get("COMPANY NAME") or "Unknown").strip()
    segment = (normalized_data.get("SEGMENT") or "Unknown").upper().strip()
    original_segment = normalized_data.get("ORIGINAL_SEGMENT", segment)
    policy_type = ((normalized_data.get("POLICY TYPE") or "")).upper().strip()
    location = normalized_data.get("LOCATION", "N/A")

    # Numeric payin (float) if parsed, and raw original token
    payin_numeric = normalized_data.get("PAYIN")  # float or None
    payin_raw = normalized_data.get("PAYIN_RAW")  # string or None

    # If no numeric payin available, DO NOT calculate — return empty payout but explanation.
    if payin_numeric is None:
        display_payin = payin_raw if payin_raw is not None else ""
        explanation = f"Non-numeric payin preserved ('{display_payin}') → No numeric payout calculated"
        return "", explanation

    # ensure numeric
    try:
        payin_value = float(payin_numeric)
    except Exception:
        display_payin = payin_raw if payin_raw is not None else str(payin_numeric)
        explanation = f"Payin parsing failed ('{display_payin}') → No numeric payout calculated"
        return "", explanation

    # If numeric payin is zero, return 0% payout
    if payin_value == 0:
        return "0.00%", "Payin is 0% → Payout 0% (no calculation applied)"

    # Determine category
    payin_category = (
        "Payin Below 20%" if payin_value <= 20 else
        "Payin 21% to 30%" if 21 <= payin_value <= 30 else
        "Payin 31% to 50%" if 31 <= payin_value <= 50 else
        "Payin Above 50%"
    )

    # LOB detection
    segment_norm = (segment.replace(" ", "") or "").upper()
    lob = "TW"
    if "BUS" in segment_norm or "SCHOOL" in segment_norm or "STAFF" in segment_norm:
        lob = "BUS"
    elif "PVT" in segment_norm or "CAR" in segment_norm:
        lob = "PVT CAR"
    elif "CV" in segment_norm or "GVW" in segment_norm or "PCV" in segment_norm or "GCV" in segment_norm:
        lob = "CV"
    elif "TAXI" in segment_norm:
        lob = "TAXI"
    elif "MISD" in segment_norm or "TRACTOR" in segment_norm or "MISC" in segment_norm:
        lob = "MISD"

    # Map segment -> formula segment
    formula_segment = get_formula_segment_for_weight_range(segment)
    display_segment = segment
    if formula_segment != segment:
        segment = formula_segment

    specific_insurers = get_all_specific_insurers_for_lob_segment(lob, segment)
    candidate_rules = []
    company_upper = company_name.upper()

    for rule in FORMULA_DATA:
        if rule["LOB"].upper() != lob.upper() or rule["SEGMENT"].upper() != segment.upper():
            continue

        insurers = [i.strip().upper() for i in str(rule.get("INSURER", "")).split(',')]
        insurer_type = None
        specificity = 0

        # specific company match
        for ins in insurers:
            if ins in ["ALL COMPANIES", "REST OF COMPANIES"] or not ins:
                continue
            if (ins in company_upper or company_upper in ins or any(word in company_upper for word in ins.split())):
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

        remarks = (rule.get("REMARKS") or "").upper().strip()
        remarks_score = 0

        if remarks == "NIL":
            remarks_score = 10
        elif payin_value > 0 and remarks == payin_category.upper():
            remarks_score = 8
        elif payin_value > 20 and "PAYIN ABOVE 20%" in remarks:
            remarks_score = 6
        elif "ALL COMPANIES" in insurers and remarks in ["ALL FUEL", "NIL", ""]:
            remarks_score = 5
        else:
            continue

        total_score = (specificity * 100) + remarks_score
        candidate_rules.append((rule, total_score, specificity, insurer_type))

    if candidate_rules:
        candidate_rules.sort(key=lambda x: x[1], reverse=True)
        matched_rule = candidate_rules[0][0]
        explanation = f"Match: LOB={lob}, Segment={segment}, Insurer={matched_rule.get('INSURER','')}, Remarks={matched_rule.get('REMARKS','NIL')}, Formula={matched_rule.get('PO')}"
    else:
        matched_rule = {"PO": "90% of Payin", "INSURER": "Default", "REMARKS": "NIL"}
        explanation = f"No Match: Default Formula=90% of Payin"

    if display_segment != segment:
        explanation += f"; Weight Range Segment: {display_segment} → Formula Segment: {segment}"
    if original_segment and original_segment.upper() != display_segment.upper():
        explanation += f"; Input Segment: {original_segment} → Matched: {display_segment}"

    po = (matched_rule.get("PO") or "").strip()
    payout = payin_value
    op_str = ""

    try:
        if "flat" in po.lower():
            m = re.search(r'(-?\d+(?:\.\d+)?)\s*%', po)
            if m:
                flat_val = float(m.group(1))
            else:
                flat_val = float(po.replace('%', '').split()[0])
            payout = flat_val
            op_str = f"= {payout:.2f}% (flat)"
        elif "% of Payin" in po or "% of payin" in po.lower():
            if "90" in po:
                factor = 0.9
            elif "88" in po:
                factor = 0.88
            elif "Less 2" in po or "less 2" in po.lower():
                factor = 0.98
            else:
                m = re.search(r'(\d+(?:\.\d+)?)\s*%', po)
                if m:
                    factor = float(m.group(1)) / 100.0
                else:
                    factor = 1.0
            payout = payin_value * factor
            op_str = f"* {factor} = {payout:.2f}%"
        elif po.startswith("-"):
            try:
                deduction = float(po.replace('%', ''))
            except:
                m = re.search(r'-?(\d+(?:\.\d+)?)', po)
                deduction = float(m.group(1)) if m else 0.0
            payout = payin_value + deduction
            op_str = f"{po} = {payout:.2f}%"
        else:
            m = re.search(r'(-?\d+(?:\.\d+)?)\s*%', po)
            if m:
                val = float(m.group(1))
                payout = val
                op_str = f"= {payout:.2f}%"
            else:
                payout = payin_value
                op_str = ""
    except Exception as e:
        payout = payin_value
        op_str = f"(calculation error: {e})"

    payout = max(0, payout)
    explanation += f"; Calculated: {payin_value:.2f}% {op_str}"

    return f"{payout:.2f}%", explanation


def extract_weight_range_from_header(header_text: str) -> Optional[str]:
    """
    Extract standardized weight range from column header
    Returns standard segment name based on detected range
    """
    header_lower = str(header_text).lower()
    
    # Define range mappings with multiple pattern variations
    range_mappings = [
        # Pattern: (regex_patterns, standard_segment_name)
        ([r'\b0\s*to\s*2\.?5?\b', r'\bupto\s*2\.?5?\b'], "Upto 2.5 GVW"),
        ([r'\b2\.?5?\s*to\s*3\.5\b'], "2.5-3.5 GVW"),
        ([r'\b3\.5\s*to\s*7\.5\b'], "3.5-7.5 GVW"),
        ([r'\b7\.5\s*to\s*12\b'], "7.5-12 GVW"),
        ([r'\b12\s*to\s*20\b'], "12-20 GVW"),
        ([r'\b20\s*to\s*40\b'], "20-40 GVW"),
        ([r'\b40\s*to\s*45\b'], "40-45 GVW"),
        ([r'\babove\s*45\b', r'\b>?\s*45t?\b', r'\b45\s*\+'], ">45 GVW"),
        ([r'\b3\.5\s*to\s*12\b'], "3.5-12 GVW"),
        ([r'\b12\s*to\s*45\b'], "12-45 GVW"),
    ]
    
    for patterns, segment_name in range_mappings:
        for pattern in patterns:
            if re.search(pattern, header_lower):
                return segment_name
    
    return None


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


def detect_table_structure(sheet_data) -> Dict:
    """
    Intelligently detect the structure of the Excel table.
    """
    header_row_index = -1
    header_map = {}
    location_cols = []
    
    # Step 1: Find header row
    for i, row in enumerate(sheet_data):
        row_vals_lower = {k: str(v).lower().strip() for k, v in row.items() if v}
        
        header_keywords = ['segment', 'product', 'policy', 'fuel', 'vehicle age', 'age', 
                          'location', 'region', 'rto', 'zone']
        
        if any(any(keyword in val for keyword in header_keywords) for val in row_vals_lower.values()):
            header_row_index = i
            
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
        first_row = sheet_data[0]
        first_row_vals = list(first_row.values())
        
        text_cells = sum(1 for v in first_row_vals[:5] if v and not str(v).replace('.','').isdigit())
        if text_cells >= 2:
            header_row_index = 0
            
            for col_key, val in first_row.items():
                if not val:
                    continue
                val_str = str(val)
                
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
        header_map['segment'] = 'col_0'
    
    # Step 4: Detect location columns if not found in headers
    if not location_cols and header_row_index >= 0:
        for col_key in sheet_data[0].keys():
            if col_key in header_map.values():
                continue
            
            sample_values = []
            for i in range(header_row_index + 1, min(header_row_index + 5, len(sheet_data))):
                val = sheet_data[i].get(col_key)
                if val and str(val).strip():
                    sample_values.append(str(val))
            
            if sample_values:
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


def restructure_pivoted_data(sheet_data, structure_info, sheet_name):
    """
    Restructure pivoted table data where locations are columns.
    """
    segment_col = structure_info['segment_col']
    policy_col = structure_info['policy_col']
    fuel_col = structure_info['fuel_col']
    vehicle_age_col = structure_info['vehicle_age_col']
    location_cols = structure_info['location_cols']
    data_start_index = structure_info['data_start_index']
    
    restructured_data = []
    
    # Validate essential columns
    if not location_cols:
        print(f"   ⚠️  Warning: No location columns found. Skipping restructure.")
        return []
    
    if not segment_col:
        print(f"   ⚠️  Warning: No segment column found. Skipping restructure.")
        return []
    
    location_names = {}
    if structure_info['header_row_index'] >= 0:
        header_row = sheet_data[structure_info['header_row_index']]
        for col_key in location_cols:
            location_names[col_key] = str(header_row.get(col_key, col_key))
    else:
        for col_key in location_cols:
            location_names[col_key] = col_key
    
    last_segment = None
    last_policy = None
    last_fuel = None
    last_vehicle_age = None
    skipped_header_rows = 0
    
    for i in range(data_start_index, len(sheet_data)):
        row = sheet_data[i]
        
        if all(v is None or str(v).strip() == "" for v in row.values()):
            continue
        
        # Skip continuation header rows
        if is_header_row_generic(row, segment_col, policy_col):
            skipped_header_rows += 1
            continue
        
        current_segment = row.get(segment_col) if segment_col else None
        if current_segment and str(current_segment).strip():
            last_segment = str(current_segment).strip()
        elif last_segment:
            current_segment = last_segment
        else:
            continue
        
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
        
        for col_key in location_cols:
            payin_value = row.get(col_key)
            
            if payin_value is None or str(payin_value).strip() == "":
                continue
            
            entry = {
                'segment': current_segment,
                'policy_type': current_policy,
                'location': location_names.get(col_key, col_key),
                'payin': str(payin_value).strip(),
                'other_info': []
            }
            
            if current_fuel:
                entry['other_info'].append(current_fuel)
            if current_vehicle_age:
                entry['other_info'].append(current_vehicle_age)
            
            entry['other_info'] = ', '.join(entry['other_info']) if entry['other_info'] else None
            
            restructured_data.append(entry)
    
    if skipped_header_rows > 0:
        print(f"   ⚠️  Skipped {skipped_header_rows} continuation header row(s)")
    
    return restructured_data


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


def parse_converted_json(json_path_or_data, original_excel_path=None):
    """
    Parse converted JSON with Royal format detection
    """
    print("\n" + "="*60)
    print("PARSING CONVERTED JSON (with Royal Format Support)")
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
    
    print(f"\n📁 Source: {source_file}")
    print(f"📁 Detected Company: {filename_company}")
    print(f"📍 Detected Location: {filename_location}")
    
    all_entries = []
    
    if "sheets" in data:
        print(f"\n📊 Processing {data['total_sheets']} sheet(s)...")
        
        for sheet_name, sheet_data in data["sheets"].items():
            print(f"\n🔄 Processing sheet: '{sheet_name}'")
            print(f"   Records in sheet: {len(sheet_data)}")
            
            # FIRST: Check if this is Royal format
            royal_structure = detect_royal_structure(sheet_data, sheet_name)
            
            if royal_structure['is_royal_format']:
                print(f"   ✅ Royal format detected!")
                extracted_list = restructure_royal_data(sheet_data, royal_structure, sheet_name)
                print(f"   ✓ Restructured into {len(extracted_list)} entries")
            else:
                # Not Royal format - check for standard pivoted table
                print(f"   📝 Standard format detection...")
                structure_info = detect_table_structure(sheet_data)
                
                print(f"   Structure detected:")
                print(f"     - Header row: {structure_info['header_row_index']}")
                print(f"     - Segment column: {structure_info['segment_col']}")
                print(f"     - Location columns: {structure_info['location_cols']}")
                
                if structure_info['location_cols']:
                    print(f"   📊 Pivoted table detected! Restructuring data...")
                    extracted_list = restructure_pivoted_data(sheet_data, structure_info, sheet_name)
                    print(f"   ✓ Restructured into {len(extracted_list)} entries")
                else:
                    print(f"   📝 Using AI extraction...")
                    
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
                            if current_segment is not None and str(current_segment).strip():
                                last_segment = current_segment
                            elif last_segment is not None:
                                row[segment_col_key] = last_segment
                            
                            processed_sheet_data.append(row)
                    else:
                        processed_sheet_data = sheet_data[data_start_index:]
                    
                    sheet_text = extract_text_from_sheet(processed_sheet_data, sheet_name)
                    extracted_list = extract_structured_data_single(sheet_text, "Excel")
            
            # Normalize and calculate payout for each entry
            for extracted in extracted_list:
                normalized = normalize_extracted_data(extracted, source_file)
                if filename_company:
                    normalized["COMPANY NAME"] = filename_company
                if filename_location and not normalized["LOCATION"]:
                    normalized["LOCATION"] = filename_location
                
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
    """
    Normalize extracted row:
    - PAYIN_RAW: exact original token (trimmed)
    - PAYIN: numeric percent (float) if numeric token found, else None
    Rule: if any numeric token found (digits or %), parse it. If parsed number is between 0 and 1 -> treat as fraction and *100.
    """
    normalized = {
        "COMPANY NAME": None,
        "ORIGINAL_SEGMENT": None,
        "SEGMENT": None,
        "POLICY TYPE": None,
        "LOCATION": None,
        "PAYIN": None,         # numeric percent (float) if parsed
        "PAYIN_RAW": None,     # original exact token (string)
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

    # NEW LOGIC: Default to Comprehensive when policy type is not specified
    if extracted.get("policy_type"):
        normalized["POLICY TYPE"] = normalize_policy_type(extracted["policy_type"])
    else:
        # Default to Comprehensive (COMP + SAOD) when not specified
        normalized["POLICY TYPE"] = "Comprehensive + SAOD"

    if extracted.get("location"):
        normalized["LOCATION"] = extracted["location"]
    elif filename_location:
        normalized["LOCATION"] = filename_location

    # Preserve raw payin exactly (trim whitespace)
    raw_payin_val = extracted.get("payin")
    if raw_payin_val is not None:
        raw_str = str(raw_payin_val).strip()
        normalized["PAYIN_RAW"] = raw_str
        # Decide numeric or non-numeric: if any digit present, treat as numeric candidate
        if re.search(r'\d', raw_str):
            # extract first numeric token (handles 64, 64%, 0.64, .64, 64.0 etc.)
            m = re.search(r'(-?\d+(?:\.\d+)?)', raw_str)
            if m:
                try:
                    num = float(m.group(1))
                    # If fraction-like (0 < abs(num) < 1) treat as percent fraction
                    if 0 < abs(num) < 1:
                        num = num * 100.0
                    normalized["PAYIN"] = num
                except:
                    normalized["PAYIN"] = None
            else:
                normalized["PAYIN"] = None
        else:
            # no digit → non-numeric token like "NA" -> preserve and don't parse
            normalized["PAYIN"] = None
    else:
        normalized["PAYIN_RAW"] = None
        normalized["PAYIN"] = None

    # other_info -> REMARK
    if extracted.get("other_info"):
        other_info_str = str(extracted["other_info"]).strip()
        if other_info_str:
            normalized["REMARK"] = [x.strip() for x in other_info_str.split(',') if x.strip()]

    return normalized





def extract_structured_data_single(text, file_type):
    """Extract structured data using OpenAI"""
    prompt = f"""You are an expert at extracting structured data from insurance payout grids. Parse the entire text intelligently.

Key Guidelines:
- Identify company: From text or infer as 'Unknown'. Common: ICICI, Reliance, Chola, Sompo, Kotak, Magma, Bajaj, Digit, Liberty, Future, Tata, IFFCO, Royal, SBI, Zuno, HDFC, Shriram.
- Segment normalization: Map to standard names like 'PVT CAR COMP + SAOD', 'TW TP', 'GCV 3W', 'SCHOOL BUS', 'MISD'. If MISC is given, treat as MISD.
- Policy type: if its in this  'Package' or 'Comp' -> 'Comprehensive', 'TP' -> 'TP', 'SAOD' -> 'SAOD' and if policy type is not given then it will be "comprehnsive + saod"
- Location: Recognize city names (Ahmedabad, Surat, Mumbai, Delhi, Bangalore, Chennai, Haryana, Vijaywada, etc.), state names (Gujarat, Maharashtra, Karnataka, Jammu & Kashmir, Telangana, etc.), and RTO codes (AS1, AS2, GJ01, MH01, ROOD, ROE, etc.) as LOCATIONS.
- Payin: Extract percentages or numbers (e.g., 0.25 -> 25%, 65 -> 65). If 0 then put 0 only.
- Other info/Remarks: Capture age, fuel, notes, logic.
- Handle pivoted tables: If rows are segments and columns are locations with payins, create an entry per cell.
- Ignore empty/header rows. Output only valid data rows.

TEXT TO ANALYZE:
{text}

Return ONLY a list of JSON objects. Format:
[
  {{"company_name": "Company", "segment": "Segment", "policy_type": "Type", "location": "Location", "payin": "Value", "other_info": "Info"}},
  ...
]
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
        print(f"\n⚠️ JSON Parse Error: {e}")
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
                        "text": "Extract all text from this image accurately, preserving structure like tables, rows, columns. Output in readable text format."
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
    
    extracted_list = extract_structured_data_single(text, "Image")
    structured_data = []
    
    for extracted in extracted_list:
        normalized = normalize_extracted_data(extracted, path)
        normalized["PAYOUT"], normalized["CALCULATION EXPLANATION"] = calculate_payout(normalized)
        structured_data.append(normalized)
    
    return {
        "file_type": "Image",
        "file_path": path,
        "image_size": img.size,
        "raw_text": text,
        "structured_data": structured_data
    }

def get_formula_segment_for_weight_range(segment: str) -> str:
    """
    Map actual segment to formula segment for matching
    - "Upto 2.5 GVW" uses its own formula
    - All other GVW ranges (2.5+) use "All GVW & PCV 3W, GCV 3W" formula
    """
    segment_upper = segment.upper().strip()
    
    # If it's exactly "Upto 2.5 GVW", keep it
    if "UPTO 2.5" in segment_upper or segment_upper == "UPTO 2.5 GVW" or "0 TO 2.5" in segment_upper or "0-2.5" in segment_upper:
        return "Upto 2.5 GVW"
    
    # If it's any other weight range (2.5+), map to "All GVW & PCV 3W, GCV 3W"
    weight_patterns = [
        "2.5-3.5", "3.5-7.5", "7.5-12", 
        "12-20", "20-40", "40-45", ">45",
        "3.5-12", "12-45", "ABOVE 45", "45T"
    ]
    
    for pattern in weight_patterns:
        if pattern in segment_upper:
            return "All GVW & PCV 3W, GCV 3W"
    
    # For PCV 3W and GCV 3W specifically
    if "PCV 3W" in segment_upper or "GCV 3W" in segment_upper:
        if "ALL GVW" in segment_upper:
            return "All GVW & PCV 3W, GCV 3W"
        elif "UPTO 2.5" not in segment_upper:
            return "All GVW & PCV 3W, GCV 3W"
        elif "PCV 3W" in segment_upper:
            return "PCV 3W"
        elif "GCV 3W" in segment_upper:
            return "GCV 3W"
    
    # Default: return as-is
    return segment


def parse_file(file_path=None):
    """Universal file parser - handles all file types including Royal format"""
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


# if __name__ == "__main__":
#     result = parse_file()
    
#     if "error" in result:
#         print(f"\n❌ Error: {result['error']}")
#     else:
#         print("\n" + "="*60)
#         print("✅ PARSING COMPLETE")
#         print("="*60)
#         print(f"\nFile Type: {result['file_type']}")
        
#         entries = result.get('structured_data', [])
        
#         if entries:
#             rows = []
#             for entry in entries:
#                 # prefer numeric PAYIN for internal value, but if missing show PAYIN_RAW in output
#                 payin_val = entry.get("PAYIN")
                
#                 if payin_val is None or payin_val == "":
#                     # show exactly what was in input (PAYIN_RAW), could be "NA" or "-" etc.
#                     payin_for_output = entry.get("PAYIN_RAW", "")
#                 else:
#                     payin_for_output = payin_val

#                 row = {
#                     "COMPANY NAME": entry.get("COMPANY NAME", "") or "",
#                     "ORIGINAL_SEGMENT": entry.get("ORIGINAL_SEGMENT", "") or "",
#                     "SEGMENT": entry.get("SEGMENT", "") or "",
#                     "POLICY TYPE": entry.get("POLICY TYPE", "") or "",
#                     "LOCATION": entry.get("LOCATION", "") or "",
#                     "PAYIN": payin_for_output,
#                     "PAYOUT": entry.get("PAYOUT", "") or "",
#                     "REMARK": "; ".join(entry.get("REMARK", [])) if entry.get("REMARK") else (entry.get("REMARK", "") or ""),
#                     "CALCULATION EXPLANATION": entry.get("CALCULATION EXPLANATION", "") or ""
#                 }
#                 rows.append(row)

#             output_df = pd.DataFrame(rows)
#             output_file = "grid_output.xlsx"
#             output_df.to_excel(output_file, index=False)

#             print(f"\n📊 Excel file created: {output_file}")
#             print(f"   Total entries: {len(output_df)}")
#             print("\n✅ Summary:")
#             first_company = output_df["COMPANY NAME"].iloc[0] if not output_df["COMPANY NAME"].isna().all() else "N/A"
#             print(f"   - Company: {first_company}")
#             unique_segments = output_df["SEGMENT"].dropna().unique()
#             print(f"   - Unique Segments: {len([s for s in unique_segments if str(s).strip()])}")
#             unique_locations = output_df["LOCATION"].dropna().unique()
#             print(f"   - Unique Locations: {len([l for l in unique_locations if str(l).strip()])}")

#             # Print sample entries
#             print("\n📋 Sample Entries:")
#             for i in range(min(5, len(rows))):
#                 r = rows[i]
#                 print(f"   {i+1}. {r['ORIGINAL_SEGMENT']} | {r.get('POLICY TYPE','N/A')} | {r.get('LOCATION','N/A')} | {r.get('PAYIN')} → {r.get('PAYOUT')}")
#         else:
#             print("\n⚠️  No entries extracted from file")

