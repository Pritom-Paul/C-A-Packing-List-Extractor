import os
import pandas as pd
import xlrd
from extract_pl_pdf_data import extract_packing_list_data

input_dir = "/home/pritom/Desktop/C&A Packing List Extractor/input/Upload"
output_dir = "/home/pritom/Desktop/C&A Packing List Extractor/output"
TARGET_SHEET = "summary sheet"

df = extract_packing_list_data(input_dir)
print(df)

# Create output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# -------- Helpers -------- #

def find_value_after_cell(rows, label):
    label = label.lower()
    for row in rows:
        for i, cell in enumerate(row):
            if isinstance(cell, str) and label in cell.strip().lower():
                for val in row[i + 1:]:
                    if val is not None and str(val).strip() != "":
                        return str(val).strip()
    return ""

def get_sheet_names(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".xls":
        return xlrd.open_workbook(file_path).sheet_names()
    else:
        return pd.ExcelFile(file_path).sheet_names

def extract_order_from_excel(file_path):
    try:
        sheet_names = get_sheet_names(file_path)
        target_sheet = next((s for s in sheet_names if s.strip().lower() == TARGET_SHEET.lower()), None)
        
        if not target_sheet:
            print(f"    ⚠️ Summary sheet not found in: {os.path.basename(file_path)}")
            return None
        
        df_sheet = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
        rows = df_sheet.fillna("").values.tolist()
        order_no = find_value_after_cell(rows, "order number")
        
        if not order_no:
            print(f"    ⚠️ Order number not found in summary sheet: {os.path.basename(file_path)}")
            return None
            
        return order_no
        
    except Exception as e:
        print(f"    ❌ Error extracting order number from {os.path.basename(file_path)}: {e}")
        return None

# -------- Validation Function -------- #

def validate_excel_file(file_path, pdf_orders):
    filename = os.path.basename(file_path)
    
    try:
        excel_order_no = extract_order_from_excel(file_path)
        if not excel_order_no:
            print(f"❌ Order number extraction failed: {filename}")
            return None
        
        if excel_order_no not in pdf_orders:
            print(f"❌ Order {excel_order_no} not in PDF data: {filename}")
            return None
        
        sheet_names = get_sheet_names(file_path)
        expected_countries = pdf_orders[excel_order_no]
        available_sheets = [s.upper() for s in sheet_names]
        missing_sheets = [c for c in expected_countries if c not in available_sheets]
        
        if missing_sheets:
            print(f"❌ Missing sheets for {excel_order_no}: {missing_sheets} in {filename}")
            return None
        
        print(f"✅ Validated {excel_order_no}: {filename}")
        return excel_order_no, file_path
            
    except Exception as e:
        print(f"❌ Error processing {filename}: {e}")
        return None

# -------- Main Driver -------- #

def update_packing_lists():
    if df.empty:
        print("❌ No PDF data available - terminating")
        return
    
    pdf_orders = {}
    for order_no, group in df.groupby('order_no'):
        pdf_orders[order_no] = group['country_iso'].unique().tolist()
    
    print(f"🔍 Validating {len(pdf_orders)} orders from PDF data...")
    
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith((".xls", ".xlsx", ".xlsm"))]
    
    if not excel_files:
        print("❌ No Excel files found - terminating")
        return
    
    print(f"📁 Found {len(excel_files)} Excel files")
    
    found_orders = set()
    validated_files = []
    
    for order_no in pdf_orders:
        matching_files = [f for f in excel_files if order_no in f]
        if matching_files:
            found_orders.add(order_no)
            print(f"⚡ Quick match found for order {order_no}")
    
    for file in excel_files:
        file_path = os.path.join(input_dir, file)
        result = validate_excel_file(file_path, pdf_orders)
        if result:
            order_no, file_path = result
            found_orders.add(order_no)
            validated_files.append((order_no, file_path))
    
    # After matching and running validate_excel_file()
    missing_orders = set(pdf_orders.keys()) - found_orders

    # Orders that have files but failed validation
    failed_validations = [o for o, f in validated_files if f is None]

    if missing_orders or len(validated_files) != len(pdf_orders):
        print("\n❌ Validation failed!")
        
        if missing_orders:
            print(f"❌ Missing Excel files for orders: {missing_orders}")
        
        # Find orders that failed sheet validation
        validated_order_nos = {o for o, p in validated_files}
        failed_orders = set(pdf_orders.keys()) - validated_order_nos
        
        if failed_orders:
            print(f"❌ Sheet validation failed for orders: {failed_orders}")
        
        print("❌ Not all orders passed validation.")
        return

    print("\n🎯 Validation completed successfully!")
    print(f"✅ All orders found and all sheets present.")

if __name__ == "__main__":
    update_packing_lists()