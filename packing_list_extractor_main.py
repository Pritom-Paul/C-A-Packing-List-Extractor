import pandas as pd
from extract_invoice import extract_invoice_data
from extract_myshipment_pl import extract_myshipment_pl

def packing_list_extractor(input_dir):
    """
    Extract invoice and packing list data, then merge them on order_no.
    
    Parameters:
    input_dir (str): Directory path containing the files to process
    
    Returns:
    pandas.DataFrame: Merged DataFrame containing both invoice and packing list data
    None: If data extraction from one or both sources fails
    """
    # Extract data from both sources
    print("Extracting invoice data...")
    invoice_data_df = extract_invoice_data(input_dir)

    print("Extracting packing list data...")
    myshipment_pl_df = extract_myshipment_pl(input_dir)
    
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
        
        return merged_pl_data
        
    else:
        print("❌ Error: Could not extract data from one or both sources")
        return None

# Main execution block - only runs when script is executed directly
if __name__ == "__main__":
    input_dir="/home/pritom/Desktop/C&A Packing List Extractor/Upload/All"
    merged_data = packing_list_extractor(input_dir)