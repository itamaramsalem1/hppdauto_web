import pandas as pd
import openpyxl
import xlrd
from datetime import datetime
import os
import re
from difflib import get_close_matches
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, numbers
import concurrent.futures
from functools import lru_cache

def normalize_name(name):
    if not name:
        return ""
    name = str(name).lower()
    name = re.sub(r"[^a-z0-9\s]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


@lru_cache(maxsize=1000)
def extract_core_from_report(report_name):
    if not report_name:
        return ""
    report_name = str(report_name).lower()

    # Remove prefix like "Total Nursing Wrkd - " if present
    if report_name.startswith("total nursing wrkd - "):
        core = report_name[21:].strip()
    else:
        core = report_name.strip()

    # Normalize
    core = re.sub(r"[^a-z0-9\s]", "", core)
    core = re.sub(r"\s+", " ", core).strip()

    # Apply overrides
    overrides = {
        "dallastown": "inners creek",
        "lancaster": "abbeyville",
        "montgomeryville": "montgomery",
        "west reading": "lebanon",
        "sunbury": "sunbury"  # just to be safe
    }
    return overrides.get(core, core)

def build_template_name_map(template_entries):
    return {entry["cleaned_name"]: entry["facility"] for entry in template_entries}

def match_report_to_template_cached(report_name, template_keys_tuple, cutoff=0.6):
    """Cached version of matching function"""
    core_name = extract_core_from_report(report_name)
    if not core_name:
        return None
    
    template_keys = list(template_keys_tuple)
    
    # Try exact match first
    if core_name in template_keys:
        return core_name
    
    # Try fuzzy matching with higher cutoff
    match = get_close_matches(core_name, template_keys, n=1, cutoff=cutoff)
    if match:
        return match[0]
    
    # Try with lower cutoff as fallback
    match = get_close_matches(core_name, template_keys, n=1, cutoff=0.3)
    return match[0] if match else None

def match_report_to_template(report_name, template_name_map, cutoff=0.6):
    core_name = extract_core_from_report(report_name)
    template_keys = list(template_name_map.keys())

    # 1. Try override or exact match
    if core_name in template_name_map:
        return template_name_map[core_name]

    # 2. Try fuzzy match
    match = get_close_matches(core_name, template_keys, n=1, cutoff=cutoff)
    if match:
        return template_name_map[match[0]]

    # 3. Try low-confidence match as fallback
    match = get_close_matches(core_name, template_keys, n=1, cutoff=0.3)
    if match:
        return template_name_map[match[0]]

    return None


def safe_float_conversion(value, default=0.0):
    """Safely convert a value to float"""
    try:
        if value is None or value == "":
            return default
        return float(value)
    except (ValueError, TypeError):
        return default

def safe_cell_value(ws, cell_ref):
    """Safely get cell value"""
    try:
        return ws[cell_ref].value
    except:
        return None

def safe_xlrd_cell_value(ws, row, col):
    """Safely get cell value from xlrd worksheet"""
    try:
        if ws.nrows > row and ws.ncols > col:
            return ws.cell_value(row, col)
        return None
    except:
        return None

def is_valid_file(filename, extension):
    """Check if file is valid (not a Mac OS hidden file or corrupt)"""
    if filename.startswith('._'):
        return False
    if not filename.lower().endswith(extension):
        return False
    return True

def extract_agency_cna_rnlpn_from_sheet2(ws2):
    """
    Extract agency staffing hours for CNAs and RN+LPNs from Sheet2.
    
    Args:
        ws2: xlrd worksheet object for Sheet2
        
    Returns:
        dict: {
            'agency_cna_hours': float,
            'agency_rnlpn_hours': float,
            'agency_total_hours': float
        }
    """
    agency_cna_hours = 0.0
    agency_rnlpn_hours = 0.0
    
    current_block_type = None  # 'agency_cna', 'agency_rn', 'agency_lpn', or None
    
    # Start scanning from row 11 (index 10) downward
    for row_idx in range(10, ws2.nrows):
        try:
            # Get value from column A (index 0)
            cell_value = ws2.cell_value(row_idx, 0)
            
            # Skip empty cells
            if not cell_value or str(cell_value).strip() == "":
                continue
                
            cell_str = str(cell_value).strip().upper()
            
            # Check if this is a header row (contains forward slashes)
            if '/' in cell_str:
                # Reset current block type
                current_block_type = None
                
                # Parse the header pattern: e.g., "806/AGY/.../CNA"
                parts = cell_str.split('/')
                
                # Check if this is an agency block (contains 'AGY')
                is_agency = any('AGY' in part for part in parts)
                
                if is_agency and len(parts) > 0:
                    # Get the last part to determine staff type
                    last_part = parts[-1].strip()
                    
                    if 'CNA' in last_part:
                        current_block_type = 'agency_cna'
                    elif 'RN' in last_part:
                        current_block_type = 'agency_rn'
                    elif 'LPN' in last_part:
                        current_block_type = 'agency_lpn'
            
            else:
                # This is a data row - extract hours from column M (index 12)
                if current_block_type and ws2.ncols > 12:  # Make sure column M exists
                    try:
                        hours_value = ws2.cell_value(row_idx, 12)  # Column M
                        hours = safe_float_conversion(hours_value)
                        
                        if current_block_type == 'agency_cna':
                            agency_cna_hours += hours
                        elif current_block_type in ['agency_rn', 'agency_lpn']:
                            agency_rnlpn_hours += hours
                            
                    except (ValueError, TypeError):
                        # Skip rows with invalid hour values
                        continue
                        
        except Exception:
            # Skip problematic rows
            continue
    
    return {
        'agency_cna_hours': agency_cna_hours,
        'agency_rnlpn_hours': agency_rnlpn_hours,
        'agency_total_hours': agency_cna_hours + agency_rnlpn_hours
    }

def compute_agency_percentages(ws3, agency_data):
    """
    Compute agency staffing percentages using Sheet3 total hours data.
    
    Args:
        ws3: xlrd worksheet object for Sheet3
        agency_data: dict from extract_agency_cna_rnlpn_from_sheet2()
        
    Returns:
        dict: {
            'actual_agency_cna_pct': float,
            'actual_agency_nurse_pct': float, 
            'actual_agency_total_pct': float,
            'actual_cna_hours': float,
            'actual_rn_hours': float,
            'actual_lpn_hours': float
        }
    """
    # Extract total hours from Sheet3
    # H13 (row 12, col 7) → actual total CNA hours
    # H12 (row 11, col 7) → LPN hours  
    # H11 (row 10, col 7) → RN hours
    actual_cna_hours = safe_float_conversion(safe_xlrd_cell_value(ws3, 12, 7))
    actual_lpn_hours = safe_float_conversion(safe_xlrd_cell_value(ws3, 11, 7))
    actual_rn_hours = safe_float_conversion(safe_xlrd_cell_value(ws3, 10, 7))
    
    # Calculate total nurse hours (RN + LPN)
    actual_rnlpn_hours = actual_rn_hours + actual_lpn_hours
    actual_total_hours = actual_cna_hours + actual_rnlpn_hours
    
    # Get agency hours from the input data
    agency_cna_hours = agency_data['agency_cna_hours']
    agency_rnlpn_hours = agency_data['agency_rnlpn_hours']
    agency_total_hours = agency_data['agency_total_hours']
    
    # Calculate percentages (handle divide-by-zero safely)
    actual_agency_cna_pct = (agency_cna_hours / actual_cna_hours * 100) if actual_cna_hours > 0 else 0.0
    actual_agency_nurse_pct = (agency_rnlpn_hours / actual_rnlpn_hours * 100) if actual_rnlpn_hours > 0 else 0.0
    actual_agency_total_pct = (agency_total_hours / actual_total_hours * 100) if actual_total_hours > 0 else 0.0
    
    return {
        'actual_agency_cna_pct': round(actual_agency_cna_pct, 2),
        'actual_agency_nurse_pct': round(actual_agency_nurse_pct, 2),
        'actual_agency_total_pct': round(actual_agency_total_pct, 2),
        'actual_cna_hours': actual_cna_hours,
        'actual_rn_hours': actual_rn_hours,
        'actual_lpn_hours': actual_lpn_hours
    }

def process_template_file(args):
    """Process a single template file - for parallel processing"""
    filepath, filename, target_date = args
    
    if not is_valid_file(filename, ".xlsx"):
        if filename.startswith('._'):
            return None, (filename, "Mac OS hidden file, skipped")
        else:
            return None, (filename, "Not .xlsx, skipped")
    
    try:
        # Use read_only=True for speed and memory efficiency
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    except Exception as e:
        return None, (filename, f"Openpyxl error: {str(e)[:100]}")

    try:
        sheet_day = str(datetime.strptime(target_date, "%Y-%m-%d").day)
        if sheet_day not in wb.sheetnames:
            wb.close()
            return None, (filename, f"No sheet named '{sheet_day}'")
            
        ws = wb[sheet_day]

        # Batch read all needed cells at once
        cell_values = {}
        cells_to_read = ["D3", "E62", "B11", "E27", "G58", "E58", "F58", "L37", "L34", "O34"]
        for cell_ref in cells_to_read:
            cell_values[cell_ref] = safe_cell_value(ws, cell_ref)

        facility_full = cell_values["D3"]
        if not facility_full:
            wb.close()
            return None, (filename, "Missing facility name in D3")
            
        cleaned_facility = normalize_name(facility_full)

        # Add back the necessary mappings (but NOT the conflicting ones)
        if "sunbury skilled nursing and rehabilitation" in cleaned_facility:
            cleaned_facility = "sunbury" 
        elif "lebanon skilled nursing and rehabilitation" in cleaned_facility:
            cleaned_facility = "lebanon"
        elif "chambersburg skilled nursing and rehabilitation" in cleaned_facility:
            cleaned_facility = "chambersburg"
        elif "pottstown skilled nursing and rehabilitation" in cleaned_facility:
            cleaned_facility = "pottstown"
        # NOTE: Do NOT add back abbeyville, inners creek, or montgomery mappings

        date_cell = cell_values["B11"]
        if not date_cell:
            wb.close()
            return None, (filename, "Missing date in B11")

        try:
            if isinstance(date_cell, datetime):
                sheet_date = date_cell.date()
            else:
                sheet_date = pd.to_datetime(date_cell).date()
        except:
            wb.close()
            return None, (filename, "Invalid date format in B11")
        
        if target_date and sheet_date != datetime.strptime(target_date, "%Y-%m-%d").date():
            wb.close()
            return None, (filename, f"Date mismatch: sheet has {sheet_date}, looking for {target_date}")

        census = safe_float_conversion(cell_values["E27"])
        if census <= 0:
            wb.close()
            return None, (filename, f"Invalid census value: {census} (census must be > 0)")

        # Calculate all values at once
        cna_hours = safe_float_conversion(cell_values["G58"])
        nurse_e_hours = safe_float_conversion(cell_values["E58"])
        nurse_f_hours = safe_float_conversion(cell_values["F58"])
        
        projected_cna_hppd = cna_hours / census
        projected_nurse_hppd = (nurse_e_hours + nurse_f_hours) / census
        projected_total_hppd = projected_cna_hppd + projected_nurse_hppd

        proj_agency_total = safe_float_conversion(cell_values["L37"]) * 100
        proj_agency_nurse = safe_float_conversion(cell_values["L34"]) * 100
        proj_agency_cna = safe_float_conversion(cell_values["O34"]) * 100

        template_entry = {
            "facility": str(facility_full),
            "cleaned_name": cleaned_facility,
            "date": sheet_date,
            "note": str(cell_values["E62"]) if cell_values["E62"] else "",
            "census": census,
            "proj_total": projected_total_hppd,
            "proj_cna": projected_cna_hppd,
            "proj_nurse": projected_nurse_hppd,
            "proj_agency_total": proj_agency_total,
            "proj_agency_cna": proj_agency_cna,
            "proj_agency_nurse": proj_agency_nurse
        }
        
        wb.close()
        return template_entry, None
        
    except Exception as e:
        try:
            wb.close()
        except:
            pass
        return None, (filename, f"Data parsing error: {str(e)[:100]}")

def process_report_file(args):
    """Process a single report file - for parallel processing"""
    filepath, filename, target_date, template_map = args
    
    if not is_valid_file(filename, ".xls"):
        if filename.startswith('._'):
            return None, (filename, "Mac OS hidden file, skipped")
        else:
            return None, (filename, "Not .xls, skipped")
    
    try:
        wb = xlrd.open_workbook(filepath)
        
        # Check for required sheets
        if "Sheet3" not in wb.sheet_names():
            return None, (filename, "No Sheet3 found")
        if "Sheet2" not in wb.sheet_names():
            return None, (filename, "No Sheet2 found")
            
        ws3 = wb.sheet_by_name("Sheet3")
        ws2 = wb.sheet_by_name("Sheet2")
        
        # Get date
        try:
            raw_date = ws3.cell_value(3, 1)
            if isinstance(raw_date, float):
                report_date = datetime(*xlrd.xldate_as_tuple(raw_date, wb.datemode)).date()
            else:
                report_date = pd.to_datetime(raw_date).date()
        except:
            return None, (filename, "Invalid date format")
        
        if target_date and report_date != datetime.strptime(target_date, "%Y-%m-%d").date():
            return None, (filename, f"Date mismatch: report has {report_date}, looking for {target_date}")

        report_facility = ws3.cell_value(4, 1)
        print("DEBUG REPORT FILE:", filename)
        print("Facility Name:", report_facility)
        print("Date from Sheet3:", ws3.cell_value(3, 1))
        print("H11:", ws3.cell_value(10, 7))  # RN
        print("H12:", ws3.cell_value(11, 7))  # LPN
        print("H13:", ws3.cell_value(12, 7))  # CNA

        if not report_facility:
            return None, (filename, "Missing facility name")

        # Get hours data from Sheet3
        try:
            actual_hours = safe_float_conversion(ws3.cell_value(13, 7))
            actual_cna_hours = safe_float_conversion(ws3.cell_value(12, 7))
            actual_rn_hours = safe_float_conversion(ws3.cell_value(11, 7))
            actual_lpn_hours = safe_float_conversion(ws3.cell_value(10, 7))
            actual_rn_lpn_hours = actual_rn_hours + actual_lpn_hours
        except Exception as e:
            return None, (filename, f"Failed to extract hours data: {str(e)[:50]}")

        # Extract agency data from Sheet2
        try:
            agency_data = extract_agency_cna_rnlpn_from_sheet2(ws2)
            agency_percentages = compute_agency_percentages(ws3, agency_data)
        except Exception as e:
            return None, (filename, f"Failed to extract agency data: {str(e)[:50]}")

        matched_template_name = match_report_to_template(report_facility, template_map)
        if not matched_template_name:
            core_name = extract_core_from_report(report_facility)
            return None, (filename, f"No matched facility name. Report: '{core_name}'")

        report_data = {
            "filename": filename,
            "report_facility": report_facility,
            "matched_template_name": matched_template_name,
            "report_date": report_date,
            "actual_hours": actual_hours,
            "actual_cna_hours": actual_cna_hours,
            "actual_rn_lpn_hours": actual_rn_lpn_hours,
            
            # New agency data
            "actual_agency_cna_pct": agency_percentages['actual_agency_cna_pct'],
            "actual_agency_nurse_pct": agency_percentages['actual_agency_nurse_pct'],
            "actual_agency_total_pct": agency_percentages['actual_agency_total_pct']
        }
        
        return report_data, None

    except Exception as e:
        return None, (filename, f"Failed to parse report: {str(e)[:100]}")

def run_hppd_comparison_for_date(templates_folder, reports_folder, target_date, output_path, progress_callback=None):
    print("Starting HPPD comparison...")

    def progress(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"Progress {pct}%: {msg}")

    progress(5, "Collecting template files...")
    template_files = []
    for root, _, files in os.walk(templates_folder):
        print("\n--- TEMPLATE FILES BEING SEEN ---")
        for filename in files:
            print("FILE:", filename)
            filepath = os.path.join(root, filename)
            template_files.append((filepath, filename, target_date))

    print(f"Processing {len(template_files)} template files...")

    print(f"\n=== TEMPLATE FILES DEBUG ===")
    print(f"Total template files found: {len(template_files)}")

    missing_facilities = ["Chambersburg", "Pottstown"]
    found_missing = []

    for filepath, filename, target_date_param in template_files:
        print(f"Found template: {filename}")
        for missing in missing_facilities:
            if missing.lower() in filename.lower():
                found_missing.append(f"✅ FOUND {missing}: {filename}")
                print(f"  --> This is {missing}!")

    print(f"\nMissing facilities tracking:")
    for item in found_missing:
        print(item)

    not_found = [facility for facility in missing_facilities
                 if not any(facility.lower() in filename.lower() for filepath, filename, target_date_param in template_files)]
    if not_found:
        print(f"❌ NOT FOUND: {not_found}")

    print(f"=== END TEMPLATE FILES DEBUG ===\n")

    template_entries = []
    skipped_templates = []

    progress(15, "Processing template files...")

    print(f"\n=== PROCESSING TEMPLATES DEBUG ===")
    processed_count = 0
    skipped_count = 0

    for template_file_args in template_files:
        filepath, filename, target_date_param = template_file_args

        is_missing_facility = any(missing.lower() in filename.lower() for missing in missing_facilities)
        if is_missing_facility:
            print(f"\n🔍 PROCESSING MISSING FACILITY: {filename}")
            print(f"   Full path: {filepath}")
            print(f"   Target date: {target_date_param}")

        entry, skip_info = process_template_file(template_file_args)

        if entry:
            template_entries.append(entry)
            processed_count += 1
            if is_missing_facility:
                print(f"   ✅ SUCCESS: Added to template_entries")
                print(f"   Facility name: {entry['facility']}")
                print(f"   Cleaned name: {entry['cleaned_name']}")
        elif skip_info:
            skipped_templates.append(skip_info)
            skipped_count += 1
            if is_missing_facility:
                print(f"   ❌ SKIPPED: {skip_info}")
        else:
            if is_missing_facility:
                print(f"   ⚠️ RETURNED NONE,NONE - THIS IS THE PROBLEM!")
                print(f"   Attempting detailed processing...")
                try:
                    result = process_template_file(template_file_args)
                    print(f"   Retry result: {result}")
                except Exception as e:
                    print(f"   Exception during retry: {e}")
                    import traceback
                    traceback.print_exc()

    print(f"\nProcessing summary:")
    print(f"Successfully processed: {processed_count}")
    print(f"Skipped: {skipped_count}")
    print(f"Total template_entries: {len(template_entries)}")
    print(f"=== END PROCESSING DEBUG ===\n")

    print(f"Successfully processed {len(template_entries)} templates, skipped {len(skipped_templates)}")

    progress(30, "Building template map...")
    template_map = build_template_name_map(template_entries)

    progress(40, "Collecting report files...")
    report_files = []
    for root, _, files in os.walk(reports_folder):
        for filename in files:
            filepath = os.path.join(root, filename)
            report_files.append((filepath, filename, target_date, template_map))

    print(f"Processing {len(report_files)} report files...")

    report_data_list = []
    skipped_reports = []

    progress(50, "Processing report files...")
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        results = executor.map(process_report_file, report_files)

        for report_data, skip_info in results:
            if report_data:
                report_data_list.append(report_data)
            elif skip_info:
                skipped_reports.append(skip_info)

    print(f"Successfully processed {len(report_data_list)} reports, skipped {len(skipped_reports)}")

    progress(65, "Matching reports to templates...")

    print("\n=== REPORT-TEMPLATE MATCH DEBUG ===")
    for report_data in report_data_list:
        matched_template_name = report_data["matched_template_name"]
        report_date = report_data["report_date"]

        candidates = [entry for entry in template_entries
                      if entry["facility"] == matched_template_name
                      and entry["date"] == report_date]

        if candidates:
            print(f"✅ MATCH: '{matched_template_name}' on {report_date} — Census: {candidates[0]['census']}")
        else:
            print(f"❌ MATCH FAILED: '{matched_template_name}' on {report_date}")
            alt_dates = [entry['date'] for entry in template_entries if entry["facility"] == matched_template_name]
            if alt_dates:
                print(f"   ➜ Found template(s) for this facility, but on different date(s): {alt_dates}")
            else:
                print("   ➜ No templates found at all for this facility")
