import os
import pandas as pd
import xlrd

def find_value_by_label(rows, label):
    """Helper function to find a value by searching for a label and returning the next non-empty cell.
    
    Args:
        rows: List of rows from the Excel sheet
        label: The label text to search for (e.g., "Port of Loading :")
    
    Returns:
        The value found next to the label, or None if not found
    """
    for row in rows:
        row_values = [str(cell).strip() for cell in row]
        if label in row_values:
            col_idx = row_values.index(label)
            # Look for the next non-empty cell to the right
            value = next((row_values[i] for i in range(col_idx + 1, len(row_values))
                        if row_values[i] and row_values[i].upper() != "NONE"), None)
            return value
    return None

def find_value_below_label(rows, label):
    """Helper function to find a value by searching for a label and returning the non-empty cell below it.
    
    Args:
        rows: List of rows from the Excel sheet
        label: The label text to search for (e.g., "Description of Goods:")
    
    Returns:
        The value found below the label, or None if not found
    """
    for row_idx, row in enumerate(rows):
        row_values = [str(cell).strip() for cell in row]
        if label in row_values:
            col_idx = row_values.index(label)
            # Look for the next non-empty cell in the same column below
            for next_row_idx in range(row_idx + 1, len(rows)):
                next_row = rows[next_row_idx]
                if col_idx < len(next_row):
                    cell_value = str(next_row[col_idx]).strip()
                    if cell_value and cell_value.upper() != "NONE":
                        return cell_value
            return None
    return None

def find_all_values_below_label(rows, label):
    """Find values below all the instance of a label.
    
    Args:
        rows: List of rows from the Excel sheet
        label: The label text to search for (e.g., "ORDER NO:")
    
    Returns:
        List of all values found below each instance of the label
    """
    values = []
    for row_idx, row in enumerate(rows):
        row_values = [str(cell).strip() for cell in row]
        if label in row_values:
            col_idx = row_values.index(label)
            # Look for the next non-empty cell in the same column below
            for next_row_idx in range(row_idx + 1, len(rows)):
                next_row = rows[next_row_idx]
                if col_idx < len(next_row):
                    cell_value = str(next_row[col_idx]).strip()
                    if cell_value and cell_value.upper() != "NONE":
                        if cell_value not in values:  # Avoid duplicates
                            values.append(cell_value)
                        break  # Move to next instance of label after finding one value
    return values if values else None

def find_invoice_no_and_date(rows):
    """Find invoice number and date."""
    for row in rows:
        row_values = [str(cell).strip() for cell in row]
        if "Inv. No. & Date:" in row_values:
            col_idx = row_values.index("Inv. No. & Date:")
            invoice_no = next((row_values[i] for i in range(col_idx + 1, len(row_values))
                               if row_values[i] and row_values[i].upper() != "NONE"), None)
            try:
                dt_idx = next(i for i in range(col_idx + 1, len(row_values)) if "DT." in row_values[i])
                dt_cell = row_values[dt_idx]
                if "DT." in dt_cell:
                    invoice_date = dt_cell.split("DT.")[-1].strip()
                else:
                    invoice_date = None
            except StopIteration:
                invoice_date = None
            return invoice_no, invoice_date
    return None, None

def find_exp_no_and_date(rows):
    """Find export number and date."""
    for row in rows:
        row_values = [str(cell).strip() for cell in row]
        if "EXP No. & Date:" in row_values:
            col_idx = row_values.index("EXP No. & Date:")
            exp_no = next((row_values[i] for i in range(col_idx + 1, len(row_values))
                           if row_values[i] and row_values[i].upper() != "NONE"), None)
            try:
                dt_idx = next(i for i in range(col_idx + 1, len(row_values)) if "DT." in row_values[i])
                dt_cell = row_values[dt_idx]
                if "DT." in dt_cell:
                    exp_date = dt_cell.split("DT.")[-1].strip()
                else:
                    exp_date = None
            except StopIteration:
                exp_date = None
            return exp_no, exp_date
    return None, None

def find_contract_no_and_date(rows):
    """ Find Contract number and Contract date """
    for row_idx, row in enumerate(rows):
        row_values = [str(cell).strip() for cell in row]

        # Check for both marker variations
        marker = None
        if "Contract No.& Date:" in row_values:
            marker = "Contract No.& Date:"

        if marker:
            col_idx = row_values.index(marker)

            # Contract number: first non-empty cell after marker
            contract_no = next(
                (row_values[i] for i in range(col_idx + 1, len(row_values))
                 if row_values[i] and row_values[i].upper() != "NONE"),
                None
            )
            contract_date = None
            try:
                dt_idx = next(i for i in range(col_idx + 1, len(row_values)) if "DT." in row_values[i])
                dt_cell = row_values[dt_idx]
                if "DT." in dt_cell:
                    # Get everything after "DT." in the same cell
                    contract_date = dt_cell.split("DT.")[-1].strip()
            except StopIteration:
                contract_date = None

            return contract_no, contract_date

    return None, None

def find_hs_code(rows):
    """Find HS Code from the rows in the format 'H.S CODE: XXXXX'."""
    for row in rows:
        row_values = [str(cell).strip() for cell in row]
        for cell in row_values:
            if "H.S CODE:" in cell.upper():
                parts = cell.split(":")
                if len(parts) > 1:
                    hs_code = parts[1].strip()
                    if hs_code and hs_code.upper() != "NONE":
                        return hs_code
    return None

def read_excel_rows(file_path):
    """Read rows from Excel file, supports xls, xlsx, xlsm."""
    ext = os.path.splitext(file_path)[1].lower()
    rows = []
    
    if ext in [".xls", ".xlsx", ".xlsm"]:
        engine = "xlrd" if ext == ".xls" else None
        try:
            # Get all sheet names and find one with "INV"
            all_sheets = pd.ExcelFile(file_path, engine=engine).sheet_names
            inv_sheets = [name for name in all_sheets if "INV" in name.upper()]
            if not inv_sheets:
                return []
            df_temp = pd.read_excel(file_path, sheet_name=inv_sheets[0], engine=engine)
            rows = df_temp.fillna("").values.tolist()
        except Exception:
            return []
    return rows

def extract_invoice_data(directory):
    """
    Extract invoice, export, and Contract numbers/dates from all Excel files in a directory.
    Returns a combined DataFrame or prints message if nothing found.
    """
    excel_files = [
        f for f in os.listdir(directory)
        if f.lower().endswith((".xls", ".xlsx", ".xlsm")) and "inv" in f.lower()
    ]

    all_dfs = []

    for filename in excel_files:
        file_path = os.path.join(directory, filename)
        rows = read_excel_rows(file_path)
        if not rows:
            print(f"❌ 'Invoice' sheet not found in {filename}. Stopping execution.")
            return pd.DataFrame()


        invoice_no, invoice_date = find_invoice_no_and_date(rows)
        exp_no, exp_date = find_exp_no_and_date(rows)
        contract_no, contract_date = find_contract_no_and_date(rows)
        goods_descption = find_value_below_label(rows, "Description Of Goods:")
        hs_code = find_hs_code(rows)
        # bank = find_value_below_label(rows, "Negotiating Bank :")
        order_nos = find_all_values_below_label(rows, "ORDER NO:")

        file_has_valid_data = False
        
        # Create one row for each order_no
        if order_nos:
            for order_no in order_nos:
                # Create the data dictionary for each order_no
                data = {
                    "invoice_no": invoice_no,
                    "invoice_date": invoice_date,
                    "exp_no": exp_no,
                    "exp_date": exp_date,
                    "contract_no": contract_no,
                    "contract_date": contract_date,
                    "goods_descption": goods_descption,
                    "hs_code": hs_code,
                    "order_no": order_no,  # Single order_no per row
                }
                
                # Check if any field is missing
                missing = [key for key, value in data.items() if not value]
                
                if not missing:
                    all_dfs.append(pd.DataFrame([data]))
                    file_has_valid_data = True
                else:
                    print(f"❌ Could not extract row from {filename} - Missing: {', '.join(missing)}")
                    # Return empty DataFrame immediately if any row fails
                    return pd.DataFrame()
        else:
            print(f"❌ No order numbers found in {filename}")
            # Return empty DataFrame immediately if no order numbers
            return pd.DataFrame()
        
        # Print success message after processing each file
        if file_has_valid_data:
            print(f"✅ Successfully extracted data from {filename}")

    if all_dfs:
        master_df = pd.concat(all_dfs, ignore_index=True)
        return master_df
    else:
        print("No valid data extracted from any Excel files.")
        return pd.DataFrame() 
# Example usage
directory = "/home/pritom/Desktop/C&A Packing List Extractor/Packing List Extractor/Upload/All"
df = extract_invoice_data(directory)
# if df is not None:
#     print(df)