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
import logging

# Set up logging for Render
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@lru_cache(maxsize=1000)
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
    report_name = str(report_name)
    if report_name.lower().startswith("total nursing wrkd - "):
        return normalize_name(report_name[21:])
    return normalize_name(re.sub(r"\s+PA\d+_\d+", "", report_name))

def build_template_name_map(template_entries):
    return {entry["cleaned_name"]: entry["facility"] for entry in template_entries}

@lru_cache(maxsize=1000)
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
    template_keys_tuple = tuple(template_name_map.keys())
    matched_key = match_report_to_template_cached(report_name, template_keys_tuple, cutoff)
    return template_name_map.get(matched_key) if matched_key else None

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
    """
    agency_cna_hours = 0.0
    agency_rnlpn_hours = 0.0
    
    current_block_type = None
    logger.info(f"🔍 Starting Sheet2 extraction - {ws2.nrows} rows, {ws2.ncols} columns")
    
    # Start scanning from row 11 (index 10) downward
    for row_idx in range(10, ws2.nrows):
        try:
            cell_value = ws2.cell_value(row_idx, 0)
            
            if not cell_value or str(cell_value).strip() == "":
                continue
                
            cell_str = str(cell_value).strip().upper()
            
            # Check if this is a header row (contains forward slashes)
            if '/' in cell_str:
                current_block_type = None
                parts = cell_str.split('/')
                is_agency = any('AGY' in part for part in parts)
                
                if is_agency and len(parts) > 0:
                    last_part = parts[-1].strip()
                    
                    if 'CNA' in last_part:
                        current_block_type = 'agency_cna'
                        logger.info(f"📌 Found Agency CNA block at row {row_idx}: {cell_str}")
                    elif 'RN' in last_part:
                        current_block_type = 'agency_rn'
                        logger.info(f"📌 Found Agency RN block at row {row_idx}: {cell_str}")
                    elif 'LPN' in last_part:
                        current_block_type = 'agency_lpn'
                        logger.info(f"📌 Found Agency LPN block at row {row_idx}: {cell_str}")
            
            else:
                # This is a data row
                if current_block_type and ws2.ncols > 12:
                    try:
                        hours_value = ws2.cell_value(row_idx, 12)  # Column M
                        hours = safe_float_conversion(hours_value)
                        
                        if hours > 0:
                            if current_block_type == 'agency_cna':
                                agency_cna_hours += hours
                                logger.info(f"   ✅ Added {hours} CNA agency hours")
                            elif current_block_type in ['agency_rn', 'agency_lpn']:
                                agency_rnlpn_hours += hours
                                logger.info(f"   ✅ Added {hours} RN/LPN agency hours")
                            
                    except (ValueError, TypeError):
                        continue
                        
        except Exception as e:
            logger.warning(f"Error processing row {row_idx}: {e}")
            continue
    
    result = {
        'agency_cna_hours': agency_cna_hours,
        'agency_rnlpn_hours': agency_rnlpn_hours,
        'agency_total_hours': agency_cna_hours + agency_rnlpn_hours
    }
    
    logger.info(f"📊 Sheet2 extraction complete: CNA={agency_cna_hours}, RN+LPN={agency_rnlpn_hours}, Total={result['agency_total_hours']}")
    return result

def compute_agency_percentages(ws3, agency_data):
    """
    Compute agency staffing percentages using Sheet3 total hours data.
    """
    # Extract total hours from Sheet3
    actual_cna_hours = safe_float_conversion(safe_xlrd_cell_value(ws3, 12, 7))
    actual_lpn_hours = safe_float_conversion(safe_xlrd_cell_value(ws3, 11, 7))
    actual_rn_hours = safe_float_conversion(safe_xlrd_cell_value(ws3, 10, 7))
    
    logger.info(f"📊 Sheet3 hours: CNA={actual_cna_hours}, RN={actual_rn_hours}, LPN={actual_lpn_hours}")
    
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
    
    result = {
        'actual_agency_cna_pct': round(actual_agency_cna_pct, 2),
        'actual_agency_nurse_pct': round(actual_agency_nurse_pct, 2),
        'actual_agency_total_pct': round(actual_agency_total_pct, 2),
        'actual_cna_hours': actual_cna_hours,
        'actual_rn_hours': actual_rn_hours,
        'actual_lpn_hours': actual_lpn_hours
    }
    
    logger.info(f"🧮 Agency percentages: CNA={result['actual_agency_cna_pct']}%, RN+LPN={result['actual_agency_nurse_pct']}%, Total={result['actual_agency_total_pct']}%")
    return result

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
        logger.info(f"🔄 Processing report: {filename}")
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

        # Extract agency data from Sheet2 - WITH LOGGING
        actual_agency_cna_pct = 0.0
        actual_agency_nurse_pct = 0.0
        actual_agency_total_pct = 0.0
        
        try:
            logger.info(f"🔍 Extracting agency data from {filename}")
            agency_data = extract_agency_cna_rnlpn_from_sheet2(ws2)
            agency_percentages = compute_agency_percentages(ws3, agency_data)
            
            actual_agency_cna_pct = agency_percentages.get('actual_agency_cna_pct', 0.0)
            actual_agency_nurse_pct = agency_percentages.get('actual_agency_nurse_pct', 0.0) 
            actual_agency_total_pct = agency_percentages.get('actual_agency_total_pct', 0.0)
            
            logger.info(f"✅ Agency extraction successful for {filename}: CNA={actual_agency_cna_pct}%, RN+LPN={actual_agency_nurse_pct}%, Total={actual_agency_total_pct}%")
            
        except Exception as e:
            logger.error(f"❌ Failed to extract agency data from {filename}: {str(e)}")

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
            
            # Agency data
            "actual_agency_cna_pct": actual_agency_cna_pct,
            "actual_agency_nurse_pct": actual_agency_nurse_pct,
            "actual_agency_total_pct": actual_agency_total_pct
        }
        
        logger.info(f"✅ Report processed: {filename} - Agency data: CNA={actual_agency_cna_pct}%, RN+LPN={actual_agency_nurse_pct}%, Total={actual_agency_total_pct}%")
        return report_data, None

    except Exception as e:
        logger.error(f"❌ Failed to parse report {filename}: {str(e)}")
        return None, (filename, f"Failed to parse report: {str(e)[:100]}")

def run_hppd_comparison_for_date(templates_folder, reports_folder, target_date, output_path):
    print("Starting HPPD comparison...")
    logger.info(f"🚀 Starting HPPD comparison for date: {target_date}")
    
    # Collect all template files
    template_files = []
    for root, _, files in os.walk(templates_folder):
        for filename in files:
            filepath = os.path.join(root, filename)
            template_files.append((filepath, filename, target_date))
    
    print(f"Processing {len(template_files)} template files...")
    
    # Process template files in parallel
    template_entries = []
    skipped_templates = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        results = executor.map(process_template_file, template_files)
        
        for entry, skip_info in results:
            if entry:
                template_entries.append(entry)
            elif skip_info:
                skipped_templates.append(skip_info)

    print(f"Successfully processed {len(template_entries)} templates, skipped {len(skipped_templates)}")
    
    # Build template map once
    template_map = build_template_name_map(template_entries)
    
    # Collect all report files
    report_files = []
    for root, _, files in os.walk(reports_folder):
        for filename in files:
            filepath = os.path.join(root, filename)
            report_files.append((filepath, filename, target_date, template_map))
    
    print(f"Processing {len(report_files)} report files...")
    logger.info(f"📊 Processing {len(report_files)} report files...")
    
    # Process report files in parallel
    report_data_list = []
    skipped_reports = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        results = executor.map(process_report_file, report_files)
        
        for report_data, skip_info in results:
            if report_data:
                report_data_list.append(report_data)
            elif skip_info:
                skipped_reports.append(skip_info)

    print(f"Successfully processed {len(report_data_list)} reports, skipped {len(skipped_reports)}")
    logger.info(f"✅ Successfully processed {len(report_data_list)} reports, skipped {len(skipped_reports)}")
    
    # Log a sample of processed reports to verify agency data
    for i, report in enumerate(report_data_list[:3]):  # Log first 3 reports
        logger.info(f"📋 Sample Report {i+1}: {report['filename']} - Agency: CNA={report['actual_agency_cna_pct']}%, RN+LPN={report['actual_agency_nurse_pct']}%, Total={report['actual_agency_total_pct']}%")
    
    # Match reports to templates and build results
    results = {}
    template_lookup = {entry["facility"]: entry for entry in template_entries}
    
    for report_data in report_data_list:
        # Find matching template
        candidates = [entry for entry in template_entries 
                     if entry["facility"] == report_data["matched_template_name"] 
                     and entry["date"] == report_data["report_date"]]
        
        if not candidates:
            skipped_reports.append((report_data["filename"], f"No matched date in template. Report date: {report_data['report_date']}"))
            continue

        t = candidates[0]
        key = (t["facility"], report_data["report_date"])
        
        # Calculate actual HPPD values
        actual_hppd = report_data["actual_hours"] / t["census"] if t["census"] > 0 else 0
        actual_cna_hppd = report_data["actual_cna_hours"] / t["census"] if t["census"] > 0 else 0
        actual_rn_lpn_hppd = report_data["actual_rn_lpn_hours"] / t["census"] if t["census"] > 0 else 0

        # LOG THE AGENCY DATA BEING PUT INTO RESULTS
        logger.info(f"🏗️ Building results for {t['facility']}: Agency data - CNA={report_data['actual_agency_cna_pct']}%, RN+LPN={report_data['actual_agency_nurse_pct']}%, Total={report_data['actual_agency_total_pct']}%")

        results[key] = [
            {
                "Facility": t["facility"],
                "Type": "Projected",
                "Total HPPD": round(t["proj_total"], 2),
                "CNA HPPD": round(t["proj_cna"], 2),
                "RN+LPN HPPD": round(t["proj_nurse"], 2),
                "Projected CNA Agency Percentage": round(t["proj_agency_cna"], 2),
                "Projected RN+LPN Agency Percentage": round(t["proj_agency_nurse"], 2),
                "Projected Total Agency Percentage": round(t["proj_agency_total"], 2),
                "Actual CNA Agency Percentage": None,
                "Actual RN+LPN Agency Percentage": None,
                "Actual Total Agency Percentage": None,
                "Notes": t.get("note", ""),
                "Date": report_data["report_date"]
            },
            {
                "Facility": t["facility"],
                "Type": "Actual",
                "Total HPPD": round(actual_hppd, 2),
                "CNA HPPD": round(actual_cna_hppd, 2),
                "RN+LPN HPPD": round(actual_rn_lpn_hppd, 2),
                "Projected CNA Agency Percentage": None,
                "Projected RN+LPN Agency Percentage": None,
                "Projected Total Agency Percentage": None,
                "Actual CNA Agency Percentage": report_data["actual_agency_cna_pct"],
                "Actual RN+LPN Agency Percentage": report_data["actual_agency_nurse_pct"],
                "Actual Total Agency Percentage": report_data["actual_agency_total_pct"],
                "Notes": t.get("note", ""),
                "Date": report_data["report_date"]
            }
        ]

    print(f"Generated results for {len(results)} facilities")
    print("Creating Excel output...")

    # Pre-calculate all difference rows and column widths
    all_difference_rows = {}
    column_headers = [
        "Facility", "Type", "Total HPPD", "CNA HPPD", "RN+LPN HPPD",
        "Projected CNA Agency Percentage", "Projected RN+LPN Agency Percentage",
        "Projected Total Agency Percentage", "Actual CNA Agency Percentage",
        "Actual RN+LPN Agency Percentage", "Actual Total Agency Percentage", 
        "Notes", "Date"
    ]
    
    # Initialize column widths with header lengths
    column_widths = {header: len(header) for header in column_headers}
    
    for key in results.keys():
        projected_row = results[key][0]
        actual_row = results[key][1]
        
        # LOG THE ACTUAL VALUES GOING INTO THE EXCEL
        logger.info(f"📊 Excel data for {projected_row['Facility']}: Actual CNA Agency = {actual_row['Actual CNA Agency Percentage']}")
        
        difference_row = {"Type": "Difference", "Facility": "", "Date": projected_row["Date"]}
        for col_name in column_headers:
            if col_name in ("Facility", "Type", "Date"):
                continue
            proj_val = projected_row.get(col_name)
            act_val = actual_row.get(col_name)
            if isinstance(proj_val, (int, float)) and isinstance(act_val, (int, float)):
                difference_row[col_name] = round(proj_val - act_val, 2)
            else:
                difference_row[col_name] = None
        
        all_difference_rows[key] = difference_row
        
        # Calculate column widths - compare header length vs content length
        for row_data in [projected_row, actual_row, difference_row]:
            for header in column_headers:
                if header == "Facility" and row_data["Type"] in ["Actual", "Difference"]:
                    content = ""
                else:
                    content = str(row_data.get(header, ""))
                
                # Make sure we use the maximum of header length and content length
                content_width = len(content)
                column_widths[header] = max(column_widths[header], content_width)

    # Create output Excel file
    wb = Workbook()
    ws = wb.active
    ws.title = "HPPD Comparison"
    current_row = 1
    header_written = False

    def write_section(title, keys):
        nonlocal current_row, header_written
        ws.cell(row=current_row, column=1, value=title).font = Font(bold=True, size=20)
        current_row += 1

        if not keys:
            ws.cell(row=current_row, column=1, value="No data available for this category.")
            current_row += 2
            return

        for col_idx, col_name in enumerate(column_headers, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=col_name)
            cell.font = Font(bold=True, size=16)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        if not header_written:
            ws.freeze_panes = ws.cell(row=current_row + 1, column=1)
            header_written = True

        current_row += 1

        for key in keys:
            projected_row = results[key][0]
            actual_row = results[key][1]
            difference_row = all_difference_rows[key]

            for row_data in [projected_row, actual_row, difference_row]:
                for col_idx, col_name in enumerate(column_headers, 1):
                    # Don't show facility name for Actual and Difference rows
                    if col_name == "Facility" and row_data["Type"] in ["Actual", "Difference"]:
                        val = ""
                    else:
                        val = row_data.get(col_name, "")
                    
                    cell = ws.cell(row=current_row, column=col_idx, value=val)
                    cell.font = Font(size=14, italic=(row_data["Type"] == "Difference"))
                    
                    # Color coding for rows
                    if row_data["Type"] == "Projected":
                        cell.fill = PatternFill("solid", fgColor="D1CFCF")  # Light grey
                    elif row_data["Type"] == "Actual":
                        cell.fill = PatternFill("solid", fgColor="FFFFFF")  # White
                    elif row_data["Type"] == "Difference":
                        # Only apply red/green/yellow fill to specific columns
                        if col_name in ("Total HPPD", "CNA HPPD", "RN+LPN HPPD"):
                            diff_val = difference_row.get(col_name)
                            if isinstance(diff_val, (int, float)):
                                if diff_val < 0:
                                    cell.fill = PatternFill("solid", fgColor="C8E6C9")  # Light green
                                else:
                                    cell.fill = PatternFill("solid", fgColor="FFCDD2")  # Light red
                            else:
                                cell.fill = PatternFill("solid", fgColor="FFFACD")  # Light yellow
                        else:
                            cell.fill = PatternFill("solid", fgColor="FFFFFF")  # No fill for other columns
                    if col_name == "Date":
                        cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
                current_row += 1

        current_row += 2

    # Categorize results
    group1, group2, group3 = [], [], []
    for key, rows in results.items():
        actual = rows[1]
        hppd = actual["Total HPPD"]
        cna = actual["CNA HPPD"]
        rn = actual["RN+LPN HPPD"]
        if 3.0 <= hppd <= 3.3 and 2.00 <= cna <= 2.06 and rn <= 1.2:
            group1.append(key)
        elif 3.0 <= hppd <= 3.3 and (cna < 2.0 or rn > 1.2):
            group2.append(key)
        elif (hppd < 3.0 or hppd > 3.3) and (cna < 2.0 or rn > 1.2):
            group3.append(key)

    write_section("Good HPPD & Good Split (3.0<HPPD<3.3, 2.00<CNA<2.06, RN+LPN<=1.20)", group1)
    write_section("Good HPPD & Bad Split (3.0<HPPD<3.3, CNA<2.00, RN+LPN>1.20)", group2)
    write_section("Bad HPPD & Bad Split (HPPD>3.3 | HPPD<3.0, CNA<2.00, RN+LPN>1.20)", group3)

    # Set column widths based on pre-calculated widths with extra padding for long headers
    for col_idx, header in enumerate(column_headers, 1):
        col_letter = get_column_letter(col_idx)
        # Add extra padding (4 characters) to ensure headers never get cut off
        ws.column_dimensions[col_letter].width = column_widths[header] + 4

    # Add skipped templates sheet
    ws_skipped = wb.create_sheet(title="Skipped Templates")
    ws_skipped.append(["File Name", "Reason", "Category"])
    for filename, reason in skipped_templates:
        category = "Mac OS Hidden File" if "Mac OS hidden" in reason else "Invalid Data" if "Invalid" in reason else "File Error"
        ws_skipped.append([filename, reason, category])
    if not skipped_templates:
        ws_skipped.append(["✅ No skipped templates", "", ""])
    ws_skipped.column_dimensions["A"].width = 40
    ws_skipped.column_dimensions["B"].width = 50
    ws_skipped.column_dimensions["C"].width = 20

    # Add skipped reports sheet
    ws_skipped_reports = wb.create_sheet(title="Skipped Reports")
    ws_skipped_reports.append(["File Name", "Reason", "Category"])
    for filename, reason in skipped_reports:
        category = "Mac OS Hidden File" if "Mac OS hidden" in reason else "Name Matching Issue" if "No matched facility" in reason else "File Error"
        ws_skipped_reports.append([filename, reason, category])
    if not skipped_reports:
        ws_skipped_reports.append(["✅ No skipped reports", "", ""])
    ws_skipped_reports.column_dimensions["A"].width = 40
    ws_skipped_reports.column_dimensions["B"].width = 50
    ws_skipped_reports.column_dimensions["C"].width = 20

    # Save the file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_output_path = os.path.join(output_path, f"HPPD_Comparison_{timestamp}.xlsx")
    wb.save(final_output_path)
    
    print("Excel file created successfully!")
    logger.info(f"✅ Excel file created successfully: {final_output_path}")
    return final_output_path