import os
import pandas as pd
import xlrd

def extract_order_no(rows):
    """Extract order number using the same pattern as in the new script"""
    label = "order number"
    for row in rows:
        for i, cell in enumerate(row):
            if isinstance(cell, str) and label in cell.strip().lower():
                # scan to the right for first non-empty value
                for val in row[i + 1:]:
                    if val is not None and str(val).strip() != "":
                        return str(val).strip()
    return ""

def extract_column_values(rows, target_header_variations, default_value="0", row_offset=2):
    """Extract values under specified header (with variations) until 'Total:' row"""
    target_col = None
    
    # locate header row
    header_row_idx = None
    for i, row in enumerate(rows):
        for target_header in target_header_variations:
            for cell in row:
                if isinstance(cell, str) and cell.strip().lower() == target_header.lower():
                    header_row_idx = i
                    break
            if header_row_idx is not None:
                break
        if header_row_idx is not None:
            break

    if header_row_idx is None:
        return []

    # find the exact column index
    for target_header in target_header_variations:
        for col_idx, cell in enumerate(rows[header_row_idx]):
            if isinstance(cell, str) and cell.strip().lower() == target_header.lower():
                target_col = col_idx
                break
        if target_col is not None:
            break

    if target_col is None:
        return []

    # start reading from the row after the header with specified offset
    values = []
    for row in rows[header_row_idx + row_offset :]:
        # stop when first column begins with "Total:"
        first_col = str(row[0]).strip().lower()
        if first_col.startswith("total"):
            break

        value = row[target_col]

        if isinstance(value, (int, float)):
            values.append(str(value))
        elif value == "" or pd.isna(value):
            values.append(default_value)
        elif isinstance(value, str) and value.strip() == "":
            values.append(default_value)
        else:
            values.append(str(value))

    return values

def validate_and_create_dataframe(result):
    """Validate the extracted data and create DataFrame if valid"""
    # Check if all lists have the same size
    list_fields = ['gross_weight', 'ctn', 'delivery_qty', 'cbm', 'net_weight', 'country_iso']
    list_sizes = {field: len(result[field]) for field in list_fields}
    
    if len(set(list_sizes.values())) != 1:
        return None, f"FAILED: List sizes don't match - {list_sizes}"
    
    num_rows = list_sizes['country_iso']
    
    # Format weight and cbm values to 2 decimal places before creating DataFrame
    formatted_gross_weight = []
    formatted_net_weight = []
    formatted_cbm = []
    
    for gw, nw, cb in zip(result['gross_weight'], result['net_weight'], result['cbm']):
        try:
            formatted_gross_weight.append(f"{float(gw):.2f}")
        except:
            formatted_gross_weight.append("0.00")
        
        try:
            formatted_net_weight.append(f"{float(nw):.2f}")
        except:
            formatted_net_weight.append("0.00")
            
        try:
            formatted_cbm.append(f"{float(cb):.2f}")
        except:
            formatted_cbm.append("0.00")
    
    # Create DataFrame with formatted values
    df_data = {
        'order_no': [result['order_no']] * num_rows,
        'country_iso': result['country_iso'],
        'ctn': result['ctn'],
        'delivery_qty': result['delivery_qty'],
        'gross_weight': formatted_gross_weight,
        'net_weight': formatted_net_weight,
        'cbm': formatted_cbm,
    }
    
    df = pd.DataFrame(df_data)
    
    # Check for partially incomplete rows (some zeros but not all)
    zero_check_columns = ['gross_weight', 'ctn', 'delivery_qty', 'cbm', 'net_weight']
    
    # Create masks for different conditions - checking for string "0" or "0.00"
    all_zero_mask = (df[zero_check_columns].isin(["0", "0.00", "0.0"])).all(axis=1)
    any_zero_mask = (df[zero_check_columns].isin(["0", "0.00", "0.0"])).any(axis=1)
    some_but_not_all_zero_mask = any_zero_mask & ~all_zero_mask

    # FAIL if any row has partial zeros (some zeros but not all)
    if some_but_not_all_zero_mask.any():
        return None, f"FAILED: Found {some_but_not_all_zero_mask.sum()} row(s) with partial zero values"
    
    # Remove rows where all numeric values are "0" (this is acceptable)
    df_cleaned = df[~all_zero_mask].reset_index(drop=True)
    
    # Check if we have any rows left after cleaning
    if len(df_cleaned) == 0:
        return None, "FAILED: All rows were removed (all zeros)"
    
    return df_cleaned, "SUCCESS"

def process_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    filename = os.path.basename(file_path)

    try:
        # read sheet names
        if ext == ".xls":
            workbook = xlrd.open_workbook(file_path)
            sheet_names = workbook.sheet_names()
        else:
            workbook = pd.ExcelFile(file_path)
            sheet_names = workbook.sheet_names

        # find summary sheet
        sheet_name = next((s for s in sheet_names if s.strip().lower() == "summary sheet"), None)
        if not sheet_name:
            return None

        # read sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        rows = df.fillna("").values.tolist()

        # extract data with appropriate row offsets
        gross_weight = extract_column_values(rows, ["Total (Gross Weight)"], default_value="0", row_offset=2)
        ctn = extract_column_values(rows, ["TOTAL CARTON"], default_value="0", row_offset=1) 
        delivery_qty = extract_column_values(rows, ["Delivery   Quantity (PCS)"], default_value="0", row_offset=2)
        cbm = extract_column_values(rows, ["Total (CBM)"], default_value="0", row_offset=2)
        net_weight = extract_column_values(rows, ["Total (Net Weight)"], default_value="0", row_offset=2)
        country_iso = extract_column_values(rows, ["Country"], default_value="N/A", row_offset=2)
        order_no = extract_order_no(rows)

        result = {
            'filename': filename,
            'order_no': order_no,
            'gross_weight': gross_weight,
            'ctn': ctn,
            'delivery_qty': delivery_qty,
            'cbm': cbm,
            'net_weight': net_weight,
            'country_iso': country_iso,
        }
        
        # Validate and create DataFrame
        df_validated, status = validate_and_create_dataframe(result)
        result['status'] = status
        result['dataframe'] = df_validated
        
        return result

    except Exception as e:
        return {'filename': filename, 'status': f'FAILED: {str(e)}', 'dataframe': None}

def process_all_files(input_dir):
    """Process all files and return consolidated results"""
    if not os.path.isdir(input_dir):
        return []

    files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xls", ".xlsx", ".xlsm"))]

    all_results = []
    for f in files:
        result = process_file(os.path.join(input_dir, f))
        if result:
            all_results.append(result)

    return all_results

def extract_myshipment_pl(directory):
    results = process_all_files(directory)
    
    successful_dfs = []
    
    for result in results:
        if result['dataframe'] is not None:
            successful_dfs.append(result['dataframe'])
    
    # Combine all successful DataFrames
    if successful_dfs:
        final_df = pd.concat(successful_dfs, ignore_index=True)
        return final_df
    else:
        print("No valid DataFrames were created from any file.")
        return None

if __name__ == "__main__":
    # Example usage
    directory = "/home/pritom/Desktop/C&A Packing List Extractor/Uploads/sample"
    df = extract_myshipment_pl(directory)
    if df is not None:
        print(df)