import os
import pdfplumber
import pandas as pd
import re

def extract_packing_list_data(directory_path):
    """
    Extract order number, country ISOs, net weight and gross weight from PDF packing lists
    """
    if not os.path.exists(directory_path):
        print(f"Error: Directory '{directory_path}' does not exist.")
        return pd.DataFrame()
    
    # Get all PDF files in the directory
    pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("No PDF files found in the directory.")
        return pd.DataFrame()
    
    print(f"Found {len(pdf_files)} PDF file(s) in the directory.\n")
    
    all_data = []
    failed_files = []
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(directory_path, pdf_file)
        print(f"Processing: {pdf_file}")
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = ""
                
                # Extract all text from PDF
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + "\n"
                
                # Extract order number using regex
                order_no = extract_order_number(full_text)
                
                # Extract country data table using regex
                country_data = extract_country_data(full_text)
                
                # Validate extraction
                if order_no and country_data:
                    for country_iso, net_weight, gross_weight in country_data:
                        all_data.append({
                            'order_no': order_no,
                            'country_iso': country_iso,
                            'net_weight': net_weight,
                            'gross_weight': gross_weight,
                        })
                    print(f"  ✓ Successfully extracted data: {len(country_data)} country entries")
                else:
                    print(f"  ✗ Failed to extract complete data from {pdf_file}")
                    failed_files.append(pdf_file)
                    
        except Exception as e:
            print(f"  ✗ Error processing {pdf_file}: {str(e)}")
            failed_files.append(pdf_file)
    
    # Create DataFrame
    if all_data:
        df = pd.DataFrame(all_data)
        print(f"\n{'='*60}")
        print(f"SUCCESS: Extracted data from {len(pdf_files) - len(failed_files)}/{len(pdf_files)} files")
        print(f"Total entries: {len(df)}")
        print(f"Failed files: {failed_files if failed_files else 'None'}")
        return df
    else:
        print("\nNo valid data extracted from any PDF files.")
        return pd.DataFrame()

def extract_order_number(text):
    """
    Extract order number from tour number line
    Expected format: "Tour number: 30261710 2251645 97239526 84771-030-46-130-001"
    Returns the last number (84771-030-46-130-001) only if it matches the exact pattern
    """
    # More strict pattern that expects exactly 5 number groups separated by 4 hyphens
    pattern = r"Tour number:\s*(?:\d+\s+)*(\d{5}-\d{3}-\d{2}-\d{3}-\d{3})"
    match = re.search(pattern, text)
    
    if match:
        order_no = match.group(1)
        # Additional validation for exactly 4 hyphens
        if order_no.count('-') == 4:
            return order_no
    return None

def extract_country_data(text):
    """
    Extract country data table starting from "Company (if packed) in KG in KG Quantity"
    and ending at "Total"
    """
    # Pattern to find the table section
    table_start = "Company (if packed) in KG in KG Quantity"
    table_end = "Total:"
    
    # Find the table section
    start_idx = text.find(table_start)
    if start_idx == -1:
        return None
    
    # Find the end of the table (Total line)
    end_idx = text.find(table_end, start_idx)
    if end_idx == -1:
        return None
    
    # Extract the table content
    table_content = text[start_idx:end_idx + len(table_end)]
    
    # Pattern to match country data rows
    # Country code (2 letters), numbers for the first two weights, quantity (ignored)
    country_pattern = r"^([A-Z]{2})\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+(\d+)$"
    
    country_data = []
    lines = table_content.split('\n')
    
    for line in lines:
        line = line.strip()
        # Skip header line and empty lines
        if table_start in line or not line:
            continue
        if table_end in line:
            break
            
        # Match country data pattern
        match = re.match(country_pattern, line)
        if match:
            country_iso = match.group(1)
            # The 3rd and 4th values are net_weight and gross_weight respectively
            net_weight = float(match.group(3))
            gross_weight = float(match.group(4))
            country_data.append((country_iso, net_weight, gross_weight))
    
    return country_data if country_data else None

# Main execution
if __name__ == "__main__":
    directory_path = "/home/pritom/Desktop/C&A Packing List Extractor/input/Upload"
    
    df = extract_packing_list_data(directory_path)
    
    if not df.empty:
        print(f"\nFinal DataFrame:")
        print(f"{'='*60}")
        print(df.to_string(index=False))
    
    else:
        print("No data extracted.")