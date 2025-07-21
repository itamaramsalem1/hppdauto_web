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
    name = name.lower()
    name = re.sub(r"[^a-z0-9\s]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name

def extract_core_from_report(report_name):
    if report_name.lower().startswith("total nursing wrkd - "):
        return normalize_name(report_name[21:])
    return normalize_name(re.sub(r"\s+PA\d+_\d+", "", report_name))

def build_template_name_map(template_entries):
    return {entry["cleaned_name"]: entry["facility"] for entry in template_entries}

def match_report_to_template(report_name, template_name_map, cutoff=0.3):
    core_name = extract_core_from_report(report_name)
    match = get_close_matches(core_name, list(template_name_map.keys()), n=1, cutoff=cutoff)
    return template_name_map[match[0]] if match else None

def run_hppd_comparison_for_date(templates_folder, reports_folder, target_date, output_path):
    skipped_templates = []
    skipped_reports = []
    template_entries = []

    for filename in os.listdir(templates_folder):
        if not filename.endswith(".xlsx"):
            continue
        filepath = os.path.join(templates_folder, filename)
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            facility_full = wb["1"]["D3"].value
            cleaned_facility = normalize_name(facility_full)
        except Exception:
            skipped_templates.append((filename, "Error opening file or missing D3"))
            continue

        for ws in wb.worksheets:
            try:
                note = ws["E62"].value
                date_cell = ws["B11"].value
                sheet_date = date_cell.date() if isinstance(date_cell, datetime) else pd.to_datetime(date_cell).date()
                if target_date and sheet_date != datetime.strptime(target_date, "%Y-%m-%d").date():
                    continue
                census = float(ws["E27"].value)
                projected_cna_hppd = float(ws["G58"].value) / census if census else None
                projected_nurse_hppd = (float(ws["E58"].value) + float(ws["F58"].value)) / census if census else None
                projected_total_hppd = projected_cna_hppd + projected_nurse_hppd if census else None
                proj_agency_total = float(ws["L37"].value * 100)
                proj_agency_nurse = float(ws["L34"].value * 100)
                proj_agency_cna = float(ws["O34"].value * 100)
                template_entries.append({
                    "facility": facility_full,
                    "cleaned_name": cleaned_facility,
                    "date": sheet_date,
                    "note": note,
                    "census": census,
                    "proj_total": projected_total_hppd,
                    "proj_cna": projected_cna_hppd,
                    "proj_nurse": projected_nurse_hppd,
                    "proj_agency_total": proj_agency_total,
                    "proj_agency_cna": proj_agency_cna,
                    "proj_agency_nurse": proj_agency_nurse
                })
            except:
                continue

    results = {}
    template_map = build_template_name_map(template_entries)

    for filename in os.listdir(reports_folder):
        if not filename.endswith(".xls"):
            continue
        filepath = os.path.join(reports_folder, filename)
        try:
            wb = xlrd.open_workbook(filepath)
            ws = wb.sheet_by_name("Sheet3")
            raw_date = ws.cell_value(3, 1)
            report_date = datetime(*xlrd.xldate_as_tuple(raw_date, wb.datemode)).date() if isinstance(raw_date, float) else pd.to_datetime(raw_date).date()
            if target_date and report_date != datetime.strptime(target_date, "%Y-%m-%d").date():
                continue
            report_facility = ws.cell_value(4, 1)
            actual_hours = float(ws.cell_value(13, 7))
            actual_cna_hours = float(ws.cell_value(12, 7))
            actual_rn_lpn_hours = float(ws.cell_value(11, 7)) + float(ws.cell_value(10, 7))

            matched_template_name = match_report_to_template(report_facility, template_map)
            if not matched_template_name:
                skipped_reports.append((filename, "No matched facility name"))
                continue

            candidates = [entry for entry in template_entries if entry["facility"] == matched_template_name and entry["date"] == report_date]
            if not candidates:
                skipped_reports.append((filename, "No matched date in template"))
                continue

            t = candidates[0]
            key = (t["facility"], report_date)
            actual_hppd = actual_hours / t["census"]
            actual_cna_hppd = actual_cna_hours / t["census"]
            actual_rn_lpn_hppd = actual_rn_lpn_hours / t["census"]

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

        except:
            continue

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

            # Calculate difference row
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

            # Write projected row
            for col_idx, col_name in enumerate(column_headers, 1):
                val = projected_row.get(col_name, "")
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = Font(size=14)
                cell.fill = PatternFill("solid", fgColor="E6F4EA")
                if col_name == "Date":
                    cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
            current_row += 1

            # Write actual row (no Facility name)
            for col_idx, col_name in enumerate(column_headers, 1):
                val = actual_row.get(col_name, "") if col_name != "Facility" else ""
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = Font(size=14)
                cell.fill = PatternFill("solid", fgColor="FFFFFF")
                if col_name == "Date":
                    cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
            current_row += 1

            # Write difference row
            for col_idx, col_name in enumerate(column_headers, 1):
                val = difference_row.get(col_name, "")
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = Font(size=14, italic=True)
                cell.fill = PatternFill("solid", fgColor="FFFACD")  # Light yellow
                if col_name == "Date":
                    cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
            current_row += 1

        current_row += 2


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
    
    for idx, header in enumerate(column_headers, 1):
        col_letter = get_column_letter(idx)
        if header == "Date":
            ws.column_dimensions[col_letter].width = 15
        elif "Percentage" in header or "Agency" in header:
            ws.column_dimensions[col_letter].width = 28
        else:
            ws.column_dimensions[col_letter].width = len(header) + 6

    wb.save(output_path)
