import pandas as pd
from extract_invoice import extract_invoice_data
from extract_myshipment_pl import extract_myshipment_pl

input_dir = "/home/pritom/Desktop/C&A Packing List Extractor/Packing List Extractor/Upload/All"

# Extract data from both sources
print("Extracting invoice data...")
invoice_data_df = extract_invoice_data(input_dir)

print("Extracting packing list data...")
myshipment_pl_df = extract_myshipment_pl(input_dir)
# print(myshipment_pl_df)

# Check if data extraction was successful
if invoice_data_df is not None and myshipment_pl_df is not None:
    # print(f"Invoice data shape: {invoice_data_df.shape}")
    # print(f"Packing list data shape: {myshipment_pl_df.shape}")
    
    # Merge DataFrames on order_no (inner join - only matching records)
    merged_pl_data = pd.merge(
        invoice_data_df, 
        myshipment_pl_df, 
        on='order_no', 
        how='inner'
    )
    
    # print(f"Merged data shape: {merged_pl_data.shape}")
    print("\nMerged DataFrame:")
    print(merged_pl_data)
    
    # Save to Excel
    output_path = "/home/pritom/Desktop/C&A Packing List Extractor/merged_packing_list.xlsx"
    merged_pl_data.to_excel(output_path, index=False)
    print(f"\n✅ Merged data saved to: {output_path}")
    
else:
    print("❌ Error: Could not extract data from one or both sources")