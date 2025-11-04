import os
import pandas as pd
import xlrd

input_dir = "/home/pritom/Desktop/C&A Packing List Extractor/input/packing_lists_extractor_data"
TARGET_SHEET = "summary sheet"

# -------- Helpers -------- #

def find_value_after_cell(rows, label):
    """Return the first non-empty cell immediately after the cell containing `label`."""
    label = label.lower()
    for row in rows:
        for i, cell in enumerate(row):
            if isinstance(cell, str) and label in cell.strip().lower():
                # scan to the right for first non-empty value
                for val in row[i + 1:]:
                    if val is not None and str(val).strip() != "":
                        return str(val).strip()
    return ""

def find_total_carton(rows):
    """Return the value in the TOTAL CARTON column and Total: row."""
    total_row = None
    total_carton_col = None

    # find 'Total:' row
    for row in rows:
        if any(isinstance(cell, str) and "total:" in cell.strip().lower() for cell in row):
            total_row = row
            break

    if total_row is None:
        return ""

    # find 'TOTAL CARTON' column (search entire rows for header)
    for row in rows:
        for idx, cell in enumerate(row):
            if isinstance(cell, str) and "total carton" in cell.strip().lower():
                total_carton_col = idx
                break
        if total_carton_col is not None:
            break

    if total_carton_col is None:
        return ""

    # return value at intersection
    val = total_row[total_carton_col]
    return int(val) if val is not None and str(val).strip().isdigit() else ""

# -------- Main Extraction -------- #

def extract_summary_sheet(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    filename = os.path.basename(file_path)

    try:
        # Get sheet names
        if ext == ".xls":
            workbook = xlrd.open_workbook(file_path)
            sheet_names = workbook.sheet_names()
        else:
            workbook = pd.ExcelFile(file_path)
            sheet_names = workbook.sheet_names

        # find target summary sheet
        target = next((s for s in sheet_names if s.strip().lower() == TARGET_SHEET), None)
        if not target:
            print(f"⚠️ Summary Sheet NOT found in {filename}")
            return None

        # Read rows
        df = pd.read_excel(file_path, sheet_name=target, header=None)
        rows = df.fillna("").values.tolist()  # convert to list of lists

        # Extract values
        order_no = find_value_after_cell(rows, "order number")
        short_po = find_value_after_cell(rows, "short po")
        total_carton = find_total_carton(rows)

        missing_fields = []
        if not order_no:
            missing_fields.append("Order Number")
        if not short_po:
            missing_fields.append("Short PO")
        if not total_carton:
            missing_fields.append("Total Carton")

        if missing_fields:
            print(f"⚠️ Could not add row for {filename}. Empty fields: {', '.join(missing_fields)}")
            return None

        return {
            "order_no": order_no,
            "short_po": short_po,
            "total_carton": total_carton
        }

    except Exception as e:
        print(f"❌ Error processing {filename}: {e}")
        return None

# -------- Main Driver -------- #

def extract_packing_lists():
    all_rows = []

    for file in os.listdir(input_dir):
        if file.lower().endswith((".xls", ".xlsx", ".xlsm")):
            row = extract_summary_sheet(os.path.join(input_dir, file))
            if row:
                all_rows.append(row)

    if all_rows:
        df = pd.DataFrame(all_rows)
        print("\n✅ Combined DataFrame:\n")
        print(df)
    else:
        print("\n⚠️ No valid rows to create DataFrame.")

if __name__ == "__main__":
    extract_packing_lists()
