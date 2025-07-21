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

def normalize_name(name):
    if not name:
        return ""
    name = str(name).lower()
    name = re.sub(r"[^a-z0-9\s]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name

def extract_core_from_report(report_name):
    if not report_name:
        return ""
    report_name = str(report_name)
    if report_name.lower().startswith("total nursing wrkd - "):
        return normalize_name(report_name[21:])
    return normalize_name(re.sub(r"\s+PA\d+_\d+", "", report_name))

def build_template_name_map(template_entries):
    return {entry["cleaned_name"]: entry["facility"] for entry in template_entries}

def match_report_to_template(report_name, template_name_map, cutoff=0.6):  # Increased cutoff
    core_name = extract_core_from_report(report_name)
    if not core_name:
        return None
    
    # Try exact match first
    if core_name in template_name_map:
        return template_name_map[core_name]
    
    # Try fuzzy matching with higher cutoff
    match = get_close_matches(core_name, list(template_name_map.keys()), n=1, cutoff=cutoff)
    if match:
        return template_name_map[match[0]]
    
    # Try with lower cutoff as fallback
    match = get_close_matches(core_name, list(template_name_map.keys()), n=1, cutoff=0.3)
    return template_name_map[match[0]] if match else None

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

def is_valid_file(filename, extension):
    """Check if file is valid (not a Mac OS hidden file or corrupt)"""
    # Skip Mac OS hidden files
    if filename.startswith('._'):
        return False
    # Check extension
    if not filename.lower().endswith(extension):
        return False
    return True

def run_hppd_comparison_for_date(templates_folder, reports_folder, target_date, output_path):
    skipped_templates = []
    skipped_reports = []
    template_entries = []

    # Process template files
    for root, _, files in os.walk(templates_folder):
        for filename in files:
            filepath = os.path.join(root, filename)
            
            if not is_valid_file(filename, ".xlsx"):
                if filename.startswith('._'):
                    skipped_templates.append((filename, "Mac OS hidden file, skipped"))
                else:
                    skipped_templates.append((filename, "Not .xlsx, skipped"))
                continue
                
            try:
                # Use read_only=True to avoid memory issues and header/footer problems
                wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            except Exception as e:
                skipped_templates.append((filename, f"Openpyxl error: {str(e)[:100]}"))
                continue

            try:
                sheet_day = str(datetime.strptime(target_date, "%Y-%m-%d").day)
                if sheet_day not in wb.sheetnames:
                    skipped_templates.append((filename, f"No sheet named '{sheet_day}'"))
                    wb.close()
                    continue
                    
                ws = wb[sheet_day]

                # Safely extract facility name
                facility_full = safe_cell_value(ws, "D3")
                if not facility_full:
                    skipped_templates.append((filename, f"Missing facility name in D3"))
                    wb.close()
                    continue
                    
                cleaned_facility = normalize_name(facility_full)

                # Safely extract other values
                note = safe_cell_value(ws, "E62")
                date_cell = safe_cell_value(ws, "B11")
                
                if not date_cell:
                    skipped_templates.append((filename, f"Missing date in B11"))
                    wb.close()
                    continue

                try:
                    if isinstance(date_cell, datetime):
                        sheet_date = date_cell.date()
                    else:
                        sheet_date = pd.to_datetime(date_cell).date()
                except:
                    skipped_templates.append((filename, f"Invalid date format in B11"))
                    wb.close()
                    continue
                
                if target_date and sheet_date != datetime.strptime(target_date, "%Y-%m-%d").date():
                    skipped_templates.append((filename, f"Date mismatch: sheet has {sheet_date}, looking for {target_date}"))
                    wb.close()
                    continue

                # Safely extract numeric values
                census = safe_float_conversion(safe_cell_value(ws, "E27"))
                if census <= 0:
                    skipped_templates.append((filename, f"Invalid census value: {census} (census must be > 0)"))
                    wb.close()
                    continue

                cna_hours = safe_float_conversion(safe_cell_value(ws, "G58"))
                nurse_e_hours = safe_float_conversion(safe_cell_value(ws, "E58"))
                nurse_f_hours = safe_float_conversion(safe_cell_value(ws, "F58"))
                
                projected_cna_hppd = cna_hours / census if census > 0 else 0
                projected_nurse_hppd = (nurse_e_hours + nurse_f_hours) / census if census > 0 else 0
                projected_total_hppd = projected_cna_hppd + projected_nurse_hppd

                # Agency percentages
                proj_agency_total = safe_float_conversion(safe_cell_value(ws, "L37")) * 100
                proj_agency_nurse = safe_float_conversion(safe_cell_value(ws, "L34")) * 100
                proj_agency_cna = safe_float_conversion(safe_cell_value(ws, "O34")) * 100

                template_entries.append({
                    "facility": str(facility_full),
                    "cleaned_name": cleaned_facility,
                    "date": sheet_date,
                    "note": str(note) if note else "",
                    "census": census,
                    "proj_total": projected_total_hppd,
                    "proj_cna": projected_cna_hppd,
                    "proj_nurse": projected_nurse_hppd,
                    "proj_agency_total": proj_agency_total,
                    "proj_agency_cna": proj_agency_cna,
                    "proj_agency_nurse": proj_agency_nurse
                })
                
                wb.close()
                
            except Exception as e:
                skipped_templates.append((filename, f"Data parsing error: {str(e)[:100]}"))
                try:
                    wb.close()
                except:
                    pass
                continue

    results = {}
    template_map = build_template_name_map(template_entries)
    
    print(f"Template facilities found: {list(template_map.values())}")

    # Process report files
    for root, _, files in os.walk(reports_folder):
        for filename in files:
            if not is_valid_file(filename, ".xls"):
                if filename.startswith('._'):
                    skipped_reports.append((filename, "Mac OS hidden file, skipped"))
                else:
                    skipped_reports.append((filename, "Not .xls, skipped"))
                continue
                
            filepath = os.path.join(root, filename)
            
            try:
                wb = xlrd.open_workbook(filepath)
                if "Sheet3" not in wb.sheet_names():
                    skipped_reports.append((filename, "No Sheet3 found"))
                    continue
                    
                ws = wb.sheet_by_name("Sheet3")
                
                # Safely get date
                try:
                    raw_date = ws.cell_value(3, 1)
                    if isinstance(raw_date, float):
                        report_date = datetime(*xlrd.xldate_as_tuple(raw_date, wb.datemode)).date()
                    else:
                        report_date = pd.to_datetime(raw_date).date()
                except:
                    skipped_reports.append((filename, "Invalid date format"))
                    continue
                
                if target_date and report_date != datetime.strptime(target_date, "%Y-%m-%d").date():
                    skipped_reports.append((filename, f"Date mismatch: report has {report_date}, looking for {target_date}"))
                    continue

                report_facility = ws.cell_value(4, 1)
                if not report_facility:
                    skipped_reports.append((filename, "Missing facility name"))
                    continue

                print(f"Trying to match report facility: '{report_facility}'")

                # Safely get hours
                try:
                    actual_hours = safe_float_conversion(ws.cell_value(13, 7))
                    actual_cna_hours = safe_float_conversion(ws.cell_value(12, 7))
                    actual_rn_hours = safe_float_conversion(ws.cell_value(11, 7))
                    actual_lpn_hours = safe_float_conversion(ws.cell_value(10, 7))
                    actual_rn_lpn_hours = actual_rn_hours + actual_lpn_hours
                except Exception as e:
                    skipped_reports.append((filename, f"Failed to extract hours data: {str(e)[:50]}"))
                    continue

                matched_template_name = match_report_to_template(report_facility, template_map)
                if not matched_template_name:
                    # Show what we tried to match for debugging
                    core_name = extract_core_from_report(report_facility)
                    skipped_reports.append((filename, f"No matched facility name. Report: '{core_name}' vs Templates: {list(template_map.keys())[:3]}..."))
                    continue

                print(f"Matched '{report_facility}' to '{matched_template_name}'")

                candidates = [entry for entry in template_entries 
                            if entry["facility"] == matched_template_name and entry["date"] == report_date]
                if not candidates:
                    skipped_reports.append((filename, f"No matched date in template. Report date: {report_date}"))
                    continue

                t = candidates[0]
                key = (t["facility"], report_date)
                actual_hppd = actual_hours / t["census"] if t["census"] > 0 else 0
                actual_cna_hppd = actual_cna_hours / t["census"] if t["census"] > 0 else 0
                actual_rn_lpn_hppd = actual_rn_lpn_hours / t["census"] if t["census"] > 0 else 0

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
                        "Notes": t.get("note", ""),
                        "Date": report_date
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
                        "Notes": t.get("note", ""),
                        "Date": report_date
                    }
                ]

            except Exception as e:
                skipped_reports.append((filename, f"Failed to parse report: {str(e)[:100]}"))
                continue

    # Create output Excel file
    wb = Workbook()
    ws = wb.active
    ws.title = "HPPD Comparison"
    current_row = 1
    header_written = False

    column_headers = [
        "Facility", "Type", "Total HPPD", "CNA HPPD", "RN+LPN HPPD",
        "Projected CNA Agency Percentage",
        "Projected RN+LPN Agency Percentage",
        "Projected Total Agency Percentage",
        "Notes",
        "Date"
    ]

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
                        cell.fill = PatternFill("solid", fgColor="E6F4EA")  # Light green
                    elif row_data["Type"] == "Actual":
                        cell.fill = PatternFill("solid", fgColor="FFFFFF")  # White
                    elif row_data["Type"] == "Difference":
                        # Determine color based on Total HPPD difference (negative = good = green, positive = bad = red)
                        total_hppd_diff = difference_row.get("Total HPPD")
                        if isinstance(total_hppd_diff, (int, float)):
                            if total_hppd_diff < 0:
                                cell.fill = PatternFill("solid", fgColor="C8E6C9")  # Light green for negative (under budget)
                            else:
                                cell.fill = PatternFill("solid", fgColor="FFCDD2")  # Light red for positive (over budget)
                        else:
                            cell.fill = PatternFill("solid", fgColor="FFFACD")  # Light yellow for no data
                    
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

    # Set column widths
    for idx, header in enumerate(column_headers, 1):
        col_letter = get_column_letter(idx)
        ws.column_dimensions[col_letter].width = 28 if "Percentage" in header else 15 if "Date" in header else len(header) + 6

    # Add skipped templates sheet with better categorization
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

    # Add skipped reports sheet with better categorization
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
    return final_output_path