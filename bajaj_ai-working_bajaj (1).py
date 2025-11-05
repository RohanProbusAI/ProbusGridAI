
# import os
# from pathlib import Path
# from PyPDF2 import PdfReader
# import pandas as pd
# from PIL import Image
# import base64
# from openai import OpenAI
# from dotenv import load_dotenv
# import json
# import re
# from typing import List, Dict, Optional, Tuple

# load_dotenv()
# client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# # Embedded Formula Data (kept as fallback)
# FORMULA_DATA = [
#     {"LOB": "TW", "SEGMENT": "1+5", "INSURER": "BAJAJ", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TW", "SEGMENT": "TW SAOD + COMP", "INSURER": "BAJAJ", "PO": "90% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "TW", "SEGMENT": "TW TP", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR COMP + SAOD", "INSURER": "BAJAJ", "PO": "90% of Payin", "REMARKS": "All Fuel"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "PVT CAR", "SEGMENT": "PVT CAR TP", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin Above 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "CV", "SEGMENT": "All GVW", "INSURER": "BAJAJ", "PO": "-3%", "REMARKS": "Payin 21% to 30%"},
#     {"LOB": "BUS", "SEGMENT": "SCHOOL BUS", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"},
#     {"LOB": "TAXI", "SEGMENT": "TAXI", "INSURER": "BAJAJ", "PO": "-2%", "REMARKS": "Payin Below 20%"},
#     {"LOB": "MISD", "SEGMENT": "MISD", "INSURER": "BAJAJ", "PO": "88% of Payin", "REMARKS": "NIL"},
# ]

# COMPANY_MAPPING = {
#     'icici': 'ICICI', 'reliance': 'Reliance', 'chola': 'Chola', 'sompo': 'Sompo',
#     'kotak': 'Kotak', 'magma': 'Magma', 'bajaj': 'Bajaj', 'digit': 'Digit', 
#     'liberty': 'Liberty', 'future': 'Future', 'tata': 'Tata', 'iffco': 'IFFCO', 
#     'royal': 'Royal', 'sbi': 'SBI', 'zuno': 'Zuno', 'hdfc': 'HDFC', 'shriram': 'Shriram'
# }

# def extract_from_filename(file_path):
#     """Extract company name and location from filename intelligently"""
#     filename = str(file_path).lower()
#     base_name = Path(file_path).stem.lower()

#     for key, value in COMPANY_MAPPING.items():
#         if key in filename:
#             return value, _extract_location(filename)

#     words = re.findall(r'\b[a-zA-Z]{3,}\b', base_name)
#     potential_companies = [w.title() for w in words if w.title() not in {'Insurance', 'General', 'Grid', 'Payout', 'Payin', 'Excel', 'Pdf', 'Img'}]
#     if potential_companies:
#         company = potential_companies[0]
#         return company, _extract_location(filename)

#     return None, _extract_location(filename)


# def _extract_location(filename):
#     """Helper to extract location from filename"""
#     location_patterns = [
#         r'\b(guj|gujarat)\b', r'\b(mh|maharashtra|mumbai)\b', 
#         r'\b(apts)\b', 
#         r'\b(dl|delhi)\b', r'\b(kar|karnataka|bangalore|bengaluru)\b',
#         r'\b(tn|tamilnadu|tamil nadu|chennai)\b', r'\b(assam)\b',
#         r'\b(tripura)\b', r'\b(meghalaya)\b', r'\b(mizoram)\b',
#         r'\b(wb|west bengal)\b', r'\b(up|uttar pradesh)\b',
#         r'\b(mp|madhya pradesh)\b', r'\b(raj|rajasthan)\b',
#         r'\b(bihar)\b', r'\b(jharkhand)\b'
#     ]
#     for pattern in location_patterns:
#         match = re.search(pattern, filename, re.IGNORECASE)
#         if match:
#             return match.group(0).upper()
#     return None


# def excel_to_json_with_formatting(excel_path):
#     """
#     🆕 Convert Excel to JSON while preserving formatting info (merged cells, headers, etc.)
#     """
#     print("\n" + "="*60)
#     print("EXCEL TO JSON CONVERTER (with Structure Preservation)")
#     print("="*60)
    
#     excel_file = pd.ExcelFile(excel_path)
#     sheet_names = excel_file.sheet_names
    
#     print(f"\n📊 Found {len(sheet_names)} worksheet(s):")
#     for i, sheet in enumerate(sheet_names, 1):
#         print(f"   {i}. {sheet}")
    
#     all_sheets_data = {}
    
#     for sheet_name in sheet_names:
#         print(f"\n🔄 Processing sheet: '{sheet_name}'")
        
#         # Read with header=None to preserve all rows
#         df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
#         df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
#         print(f"   Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        
#         # Convert to records with row/column indices
#         json_data = []
#         for row_idx, row in df.iterrows():
#             row_data = {}
#             for col_idx, value in enumerate(row):
#                 if pd.notna(value) and str(value).strip():
#                     row_data[f"col_{col_idx}"] = str(value).strip()
#             if row_data:  # Only add non-empty rows
#                 json_data.append({
#                     "row_index": int(row_idx),
#                     "data": row_data
#                 })
        
#         all_sheets_data[sheet_name] = {
#             "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
#             "data": json_data
#         }
        
#         print(f"   ✓ Converted to {len(json_data)} JSON records")
    
#     output_data = {
#         "source_file": str(excel_path),
#         "total_sheets": len(sheet_names),
#         "sheets": all_sheets_data
#     }
    
#     return output_data


# def call_openai_for_extraction(sheet_text: str, sheet_name: str, company_name: str, chunk_size: int = 50) -> List[Dict]:
#     """
#     🆕 Use OpenAI GPT-4 to intelligently extract insurance rate data from Excel sheet text
#     Uses chunking to handle large sheets
#     """
#     # Split sheet into manageable chunks
#     lines = sheet_text.split('\n')
#     header_lines = lines[:10]  # Keep header context
#     data_lines = [l for l in lines[10:] if l.strip() and l.startswith('Row')]
    
#     all_extracted = []
    
#     # Process in chunks
#     num_chunks = (len(data_lines) + chunk_size - 1) // chunk_size
    
#     if num_chunks > 1:
#         print(f"   📦 Large sheet detected - processing in {num_chunks} chunks...")
    
#     for chunk_idx in range(num_chunks):
#         start_idx = chunk_idx * chunk_size
#         end_idx = min((chunk_idx + 1) * chunk_size, len(data_lines))
#         chunk_lines = data_lines[start_idx:end_idx]
        
#         chunk_text = '\n'.join(header_lines) + '\n\n' + '\n'.join(chunk_lines)
        
#         if num_chunks > 1:
#             print(f"   🔄 Processing chunk {chunk_idx + 1}/{num_chunks} (rows {start_idx}-{end_idx})...")
        
#         prompt = f"""You are an expert at extracting insurance rate data from Excel spreadsheets.

# **TASK**: Extract ALL insurance rate entries from this Excel worksheet chunk.

# **WORKSHEET NAME**: {sheet_name}
# **COMPANY**: {company_name}

# **EXCEL DATA**:
# {chunk_text}

# **IMPORTANT INSTRUCTIONS**:
# 1. This Excel sheet contains insurance rates for different locations/RTOs
# 2. There may be MULTIPLE rate columns (e.g., Petrol Rate, Diesel Rate, CNG Rate, Electric Rate)
# 3. There are also COMMENT/REMARK columns for each fuel type (e.g., "Petrol Comments for all Channels", "Diesel Comments", "CNG Comments")
# 4. Extract EACH location-fuel combination as a SEPARATE entry

# 5. **CRITICAL - FUEL-SPECIFIC DATA EXTRACTION**:
#    Look for patterns like:
#    - "Rate for RSME/W eb-Sales Aggregators" columns for Petrol, Diesel, CNG
#    - "Petrol Comment s for all Channels" or "Petrol Comments"
#    - "Diesel Comments for all channels" or "Diesel Comments"
#    - "CNG Comments for all channels" or "CNG Comments"
   
#    Example Row:
#    - Location: GUJARAT
#    - Petrol Rate: 50.50% | Petrol Comments: "0% for Zen"
#    - Diesel Rate: IRDA | Diesel Comments: ""
#    - CNG Rate: 0% | CNG Comments: "Blocked @ no payout in system"
   
#    Create 3 SEPARATE entries:
#    * Entry 1: Location=GUJARAT, Fuel=Petrol, Payin=50.50%, Remarks="0% for Zen"
#    * Entry 2: Location=GUJARAT, Fuel=Diesel, Payin=IRDA, Remarks=""
#    * Entry 3: Location=GUJARAT, Fuel=CNG, Payin=0%, Remarks="Blocked @ no payout in system"

# 6. **REMARKS EXTRACTION RULES**:
#    - Each fuel type has its OWN comment/remark column
#    - Extract the remark that corresponds to THAT specific fuel type
#    - DO NOT mix remarks from different fuel types
#    - If Petrol has "0% for Zen" comment, only add it to Petrol entry
#    - If CNG has "Blocked" comment, only add it to CNG entry
#    - Common patterns: "0% for Zen", "Blocked @ no payout", "IRDA guidelines", etc.

# 7. Look for:
#    - Location/RTO/State names (e.g., ASSAM, BIHAR, GUJARAT, CHENNAI, etc.)
#    - District/City names if present
#    - Rate percentages (e.g., 45.50%, 55.00%, IRDA)
#    - Policy types (TP, SATP, Comprehensive, COMP, SAOD)
#    - Segment/Vehicle type (Two Wheeler, Private Car, Commercial Vehicle, etc.)

# 8. **VALUE NORMALIZATION**:
#    - If a rate is "IRDA", keep it as "IRDA"
#    - If a rate says "Blocked @ no payout" or "Blocked", mark payin as "BLOCKED"
#    - Convert percentage values: "0.455" → "45.50", "45.50%" → "45.50"
#    - Extract segment from worksheet name if possible (e.g., "PC- SATP" → "PVT CAR TP")

# **OUTPUT FORMAT** (JSON array):
# [
#   {{
#     "location": "LOCATION_NAME",
#     "district": null,
#     "fuel_type": "Petrol",
#     "segment": "PVT CAR TP",
#     "policy_type": "SATP",
#     "payin": "45.50",
#     "remarks": "0% for Zen"
#   }},
#   {{
#     "location": "LOCATION_NAME",
#     "district": null,
#     "fuel_type": "Diesel",
#     "segment": "PVT CAR TP",
#     "policy_type": "SATP",
#     "payin": "IRDA",
#     "remarks": ""
#   }},
#   {{
#     "location": "LOCATION_NAME",
#     "district": null,
#     "fuel_type": "CNG",
#     "segment": "PVT CAR TP",
#     "policy_type": "SATP",
#     "payin": "BLOCKED",
#     "remarks": "Blocked @ no payout in system"
#   }}
# ]

# **RESPOND ONLY WITH THE JSON ARRAY, NO OTHER TEXT. MAXIMUM 100 ENTRIES PER RESPONSE.**
# """

#         try:
#             response = client.chat.completions.create(
#                 model="gpt-4o-mini",
#                 messages=[
#                     {"role": "system", "content": "You are an expert at extracting structured data from insurance rate grids. Always respond with valid JSON only."},
#                     {"role": "user", "content": prompt}
#                 ],
#                 temperature=0.1,
#                 max_tokens=8000  # Increased token limit
#             )
            
#             ai_response = response.choices[0].message.content.strip()
            
#             # Remove markdown code blocks if present
#             if ai_response.startswith("```json"):
#                 ai_response = ai_response[7:]
#             if ai_response.startswith("```"):
#                 ai_response = ai_response[3:]
#             if ai_response.endswith("```"):
#                 ai_response = ai_response[:-3]
            
#             ai_response = ai_response.strip()
            
#             # Try to fix incomplete JSON
#             if not ai_response.endswith(']'):
#                 # Find last complete entry
#                 last_complete = ai_response.rfind('}')
#                 if last_complete > 0:
#                     ai_response = ai_response[:last_complete + 1] + '\n]'
#                     print(f"   ⚠️ Fixed truncated JSON in chunk {chunk_idx + 1}")
            
#             # Parse JSON
#             chunk_data = json.loads(ai_response)
#             all_extracted.extend(chunk_data)
            
#             if num_chunks > 1:
#                 print(f"   ✓ Chunk {chunk_idx + 1}: Extracted {len(chunk_data)} entries")
            
#         except json.JSONDecodeError as e:
#             print(f"   ❌ JSON decode error in chunk {chunk_idx + 1}: {e}")
#             print(f"   Response preview: {ai_response[:300]}...")
#             # Continue with next chunk instead of failing
#             continue
#         except Exception as e:
#             print(f"   ❌ API error in chunk {chunk_idx + 1}: {e}")
#             continue
    
#     print(f"   🤖 AI extracted {len(all_extracted)} total entries from {num_chunks} chunks")
#     return all_extracted


# def normalize_segment(segment_str):
#     """Normalize segment names to standard format"""
#     if not segment_str or pd.isna(segment_str):
#         return "Unknown"
    
#     segment_lower = re.sub(r'\s+', ' ', str(segment_str).lower().strip())

#     if segment_lower == "apts":
#         return segment_str
    
#     if 'school bus' in segment_lower:
#         return "SCHOOL BUS"
#     if 'staff bus' in segment_lower:
#         return "STAFF BUS"
#     if 'bus' in segment_lower:
#         return "SCHOOL BUS"
#     if 'taxi' in segment_lower:
#         return "TAXI"
    
#     if any(x in segment_lower for x in ['tw', '2w', 'two wheeler', 'bike', 'scooter']):
#         if 'saod' in segment_lower or 'comp' in segment_lower:
#             return "TW SAOD + COMP"
#         elif 'tp' in segment_lower or 'satp' in segment_lower:
#             return "TW TP"
#         elif '1+5' in segment_lower:
#             return "1+5"
#         return "TW"
    
#     if any(x in segment_lower for x in ['pvt car', 'private car', 'car', 'pc']):
#         if 'comp' in segment_lower or 'saod' in segment_lower:
#             return "PVT CAR COMP + SAOD"
#         elif 'tp' in segment_lower or 'satp' in segment_lower:
#             return "PVT CAR TP"
#         return "PVT CAR TP"
    
#     if any(x in segment_lower for x in ['cv', 'gcv', 'pcv', 'commercial']):
#         return "All GVW & PCV 3W, GCV 3W"
    
#     if 'misd' in segment_lower or 'tractor' in segment_lower:
#         return "MISD"
    
#     return segment_str


# def normalize_policy_type(policy_str):
#     """Normalize policy type"""
#     if not policy_str or pd.isna(policy_str):
#         return None
    
#     policy_lower = str(policy_str).lower().strip()
    
#     if 'package' in policy_lower or 'comp' in policy_lower:
#         return "Comprehensive"
#     elif 'saod' in policy_lower or 'od' in policy_lower:
#         return "SAOD"
#     elif 'tp' in policy_lower or 'satp' in policy_lower:
#         return "TP"
    
#     return policy_str.upper()


# def extract_percentage(value) -> Optional[float]:
#     """Extract numeric percentage from various formats"""
#     if pd.isna(value) or value is None:
#         return None
    
#     try:
#         num_val = float(value)
#         if 0 < num_val < 1:
#             return num_val * 100
#         elif 0 <= num_val <= 100:
#             return num_val
#     except:
#         pass
    
#     value_str = str(value)
#     match = re.search(r'(\d+(?:\.\d+)?)\s*%?', value_str)
#     if match:
#         num_val = float(match.group(1))
#         if 0 <= num_val <= 100:
#             return num_val
    
#     return None


# def calculate_payout(normalized_data):
#     """Calculate payout using formula matching"""
#     company_name = normalized_data.get("COMPANY NAME", "").strip()
#     segment = normalized_data.get("SEGMENT", "").upper().strip()
#     policy_type = (normalized_data.get("POLICY TYPE", "") or "").upper().strip()
#     payin = normalized_data.get("PAYIN", 0)
    
#     if payin is None or payin == "" or str(payin).strip() == "" or str(payin).strip().upper() == "IRDA":
#         return "As per IRDA", "Payin is IRDA regulated"
    
#     if str(payin).upper() == "BLOCKED" or "BLOCKED" in str(payin).upper():
#         return "0.00%", "Blocked - No payout"
    
#     try:
#         payin_value = float(str(payin).replace('%', '').strip())
#     except (ValueError, AttributeError):
#         return str(payin), "Non-numeric payin value"
    
#     if payin_value == 0:
#         return "0.00%", "Payin is 0%"
    
#     # Default formula: 90% of payin
#     payout = payin_value * 0.9
#     explanation = f"Default: 90% of {payin_value:.2f}% = {payout:.2f}%"
    
#     return f"{payout:.2f}%", explanation


# def parse_file_with_ai(file_path=None, user_company=None):
#     """
#     🆕 AI-POWERED PARSER
#     Parse Excel file using OpenAI to understand complex structures
#     """
#     if file_path is None:
#         file_path = input("Enter file path: ")
    
#     if not os.path.exists(file_path):
#         return {"error": f"File not found: {file_path}"}
    
#     ext = Path(file_path).suffix.lower()
    
#     if ext not in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
#         return {"error": f"Unsupported file type: {ext}. Only Excel files supported."}
    
#     try:
#         print("\n" + "="*60)
#         print("🤖 AI-POWERED EXCEL PARSER")
#         print("="*60)
        
#         # Get company name
#         if user_company is None:
#             filename_company, _ = extract_from_filename(file_path)
#             if filename_company:
#                 print(f"\n📁 Detected company from filename: {filename_company}")
#                 use_detected = input("Use this company name? (y/n): ").strip().lower()
#                 if use_detected == 'y':
#                     user_company = filename_company
            
#             if user_company is None:
#                 user_company = input("📝 Enter COMPANY NAME (e.g., BAJAJ, ICICI): ").strip()
        
#         # Convert Excel to structured JSON
#         json_data = excel_to_json_with_formatting(file_path)
        
#         # Process each worksheet with AI
#         all_worksheets_data = {}
        
#         for sheet_name, sheet_info in json_data["sheets"].items():
#             print(f"\n{'='*60}")
#             print(f"📑 PROCESSING: '{sheet_name}' with AI")
#             print(f"{'='*60}")
            
#             # Convert sheet data to text for AI
#             sheet_text = f"Sheet: {sheet_name}\n"
#             sheet_text += f"Dimensions: {sheet_info['shape']['rows']} rows × {sheet_info['shape']['cols']} columns\n\n"
            
#             for entry in sheet_info["data"]:
#                 row_idx = entry["row_index"]
#                 row_data = entry["data"]
#                 sheet_text += f"Row {row_idx}: " + " | ".join([f"{k}={v}" for k, v in row_data.items()]) + "\n"
            
#             # Call OpenAI for extraction
#             print(f"   🤖 Sending to OpenAI for intelligent extraction...")
#             ai_extracted = call_openai_for_extraction(sheet_text, sheet_name, user_company)
            
#             if not ai_extracted:
#                 print(f"   ⚠️ No data extracted by AI for '{sheet_name}'")
#                 continue
            
#             # Process AI extracted data
#             processed_entries = []
#             for entry in ai_extracted:
#                 # Normalize payin value
#                 payin_raw = entry.get("payin", "")
#                 payin_normalized = None
                
#                 if str(payin_raw).upper() in ["IRDA", "BLOCKED"]:
#                     payin_normalized = str(payin_raw).upper()
#                 else:
#                     # Try to convert to percentage
#                     try:
#                         payin_float = float(str(payin_raw).replace('%', '').strip())
#                         # If value is between 0-1, convert to percentage
#                         if 0 < payin_float < 1:
#                             payin_normalized = payin_float * 100
#                         elif 0 <= payin_float <= 100:
#                             payin_normalized = payin_float
#                         else:
#                             payin_normalized = payin_raw
#                     except (ValueError, AttributeError):
#                         payin_normalized = payin_raw
                
#                 normalized = {
#                     "COMPANY NAME": user_company,
#                     "SEGMENT": normalize_segment(entry.get("segment", "Unknown")),
#                     "POLICY TYPE": normalize_policy_type(entry.get("policy_type")),
#                     "LOCATION": entry.get("location", ""),
#                     "DISTRICT": entry.get("district"),
#                     "FUEL TYPE": entry.get("fuel_type"),
#                     "PAYIN": payin_normalized,
#                     "PAYOUT": "",
#                     "REMARK": [],
#                     "CALCULATION EXPLANATION": ""
#                 }
                
#                 # Combine location and district
#                 if normalized["DISTRICT"]:
#                     normalized["LOCATION"] = f"{normalized['LOCATION']}-{normalized['DISTRICT']}"
                
#                 # Add fuel-specific remarks from AI
#                 fuel_remark = entry.get("remarks", "")
#                 if fuel_remark and str(fuel_remark).strip() and str(fuel_remark).strip().lower() not in ["none", "null", "", "-"]:
#                     normalized["REMARK"].append(str(fuel_remark).strip())
                
#                 # Add fuel type label only if no specific remarks
#                 if normalized["FUEL TYPE"] and not normalized["REMARK"]:
#                     normalized["REMARK"].append(f"Fuel: {normalized['FUEL TYPE']}")
#                 elif normalized["FUEL TYPE"]:
#                     # Add fuel type at the beginning
#                     normalized["REMARK"].insert(0, f"[{normalized['FUEL TYPE']}]")
                
#                 # Calculate payout
#                 payout, explanation = calculate_payout(normalized)
#                 normalized["PAYOUT"] = payout
#                 normalized["CALCULATION EXPLANATION"] = explanation
                
#                 processed_entries.append(normalized)
            
#             all_worksheets_data[sheet_name] = processed_entries
#             print(f"   ✅ Processed {len(processed_entries)} entries")
        
#         return {
#             "file_type": "Excel (AI-Powered Extraction)",
#             "file_path": file_path,
#             "worksheets": all_worksheets_data
#         }
        
#     except Exception as e:
#         import traceback
#         return {"error": f"Error: {str(e)}\n{traceback.format_exc()}"}


# if __name__ == "__main__":
#     result = parse_file_with_ai(user_company=None)
    
#     if "error" in result:
#         print(f"\n❌ Error: {result['error']}")
#     else:
#         print("\n" + "="*60)
#         print("✅ AI PARSING COMPLETE")
#         print("="*60)
        
#         worksheets_data = result.get('worksheets', {})
        
#         if worksheets_data:
#             # Create Excel output
#             output_file = "ai_grid_output.xlsx"
            
#             with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#                 total_entries = 0
                
#                 for sheet_name, entries in worksheets_data.items():
#                     if not entries:
#                         continue
                    
#                     output_data = {
#                         "COMPANY NAME": [],
#                         "SEGMENT": [],
#                         "POLICY TYPE": [],
#                         "LOCATION": [],
#                         "FUEL TYPE": [],
#                         "PAYIN": [],
#                         "PAYOUT": [],
#                         "REMARK": [],
#                         "CALCULATION EXPLANATION": []
#                     }
                    
#                     for entry in entries:
#                         output_data["COMPANY NAME"].append(entry.get("COMPANY NAME", ""))
#                         output_data["SEGMENT"].append(entry.get("SEGMENT", ""))
#                         output_data["POLICY TYPE"].append(entry.get("POLICY TYPE", ""))
#                         output_data["LOCATION"].append(entry.get("LOCATION", ""))
#                         output_data["FUEL TYPE"].append(entry.get("FUEL TYPE", ""))
#                         output_data["PAYIN"].append(entry.get("PAYIN", ""))
#                         output_data["PAYOUT"].append(entry.get("PAYOUT", ""))
#                         output_data["REMARK"].append("; ".join(entry.get("REMARK", [])))
#                         output_data["CALCULATION EXPLANATION"].append(entry.get("CALCULATION EXPLANATION", ""))
                    
#                     output_df = pd.DataFrame(output_data)
#                     safe_sheet_name = sheet_name[:31]
#                     output_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    
#                     print(f"\n   ✓ '{sheet_name}': {len(entries)} entries")
#                     total_entries += len(entries)
            
#             print(f"\n📊 Excel output created: {output_file}")
#             print(f"   Total worksheets: {len(worksheets_data)}")
#             print(f"   Total entries: {total_entries}")
            
#             # Print summary with sample entries
#             print(f"\n📋 SUMMARY:")
#             for sheet_name, entries in worksheets_data.items():
#                 if entries:
#                     print(f"\n   📑 {sheet_name}:")
#                     print(f"      - Total entries: {len(entries)}")
                    
#                     # Count fuel types
#                     fuel_types = {}
#                     for e in entries:
#                         ft = e.get('FUEL TYPE', 'Not specified')
#                         fuel_types[ft] = fuel_types.get(ft, 0) + 1
                    
#                     if len(fuel_types) > 1:
#                         print(f"      - Fuel type breakdown:")
#                         for fuel, count in fuel_types.items():
#                             print(f"         * {fuel}: {count} entries")
                    
#                     # Show sample entries with remarks
#                     print(f"\n      📝 Sample entries with fuel-specific remarks:")
#                     sample_with_remarks = [e for e in entries[:20] if e.get('REMARK') and any(e.get('REMARK'))]
#                     for i, sample in enumerate(sample_with_remarks[:3], 1):
#                         remarks_str = "; ".join(sample.get('REMARK', []))
#                         print(f"         {i}. {sample.get('LOCATION', 'N/A')} - {sample.get('FUEL TYPE', 'N/A')}")
#                         print(f"            Payin: {sample.get('PAYIN', 'N/A')} | Payout: {sample.get('PAYOUT', 'N/A')}")
#                         if remarks_str:
#                             print(f"            Remarks: {remarks_str}")
#         else:
#             print("\n⚠️ No entries extracted")