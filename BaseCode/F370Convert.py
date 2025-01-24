import os
import pandas as pd

def find_matching_column(columns, keywords):
    """
    Finds the first column that matches any of the given keywords.
    Matching is case-insensitive and ignores leading/trailing spaces.
    """
    for col in columns:
        for keyword in keywords:
            if keyword.lower() in col.lower().strip():
                return col
    return None

def process_f370_files(input_folder, output_file):
    """
    Processes F370 print job files, detects relevant columns dynamically, splits filament types (ABS and TPU),
    and combines the data into a single CSV.
    
    Args:
        input_folder (str): Path to the folder containing the input files.
        output_file (str): Path to save the combined output CSV.
    
    Returns:
        pd.DataFrame: Combined dataset.
    """
    combined_df = pd.DataFrame()
    
    # Iterate over all files in the folder
    for file_name in os.listdir(input_folder):
        if "_Print_F370" in file_name and file_name.endswith(".csv"):
            file_path = os.path.join(input_folder, file_name)
            print(f"Processing file: {file_name}")
            
            # Read the file
            try:
                df = pd.read_csv(file_path, header=0)
            except Exception as e:
                print(f"Error reading file {file_name}: {e}")
                continue
            
            # Dynamically find the columns
            company_name_col = find_matching_column(df.columns, ["Company", "CompanyName"])
            print_time_col = find_matching_column(df.columns, ["Print Time", "h:mm"])
            support_col = find_matching_column(df.columns, ["Support", "QSR"])
            abs_col = find_matching_column(df.columns, ["ABS"])
            tpu_col = find_matching_column(df.columns, ["TPU"])
            
            if not company_name_col:
                print(f"Warning: Could not find 'Company Name' column in {file_name}. Skipping file.")
                continue
            if not all([print_time_col, support_col]):
                print(f"Warning: Missing one or more critical columns in {file_name}. Skipping file.")
                print(f"Detected columns: {df.columns.tolist()}")
                continue
            
            # Initialize ABS and TPU columns
            df["PC-ABS BLK"] = 0.0
            df["TPU 92A - Black"] = 0.0
            
            # Fill ABS or TPU columns based on detected material
            if abs_col:
                df["PC-ABS BLK"] = df[abs_col].fillna(0)
            if tpu_col:
                df["TPU 92A - Black"] = df[tpu_col].fillna(0)
            
            # Rename other columns for consistency
            df = df.rename(columns={
                company_name_col: "Company Name",
                print_time_col: "Print Time",
                support_col: "QSR support"
            })
            
            # Select and standardize required columns
            required_columns = ["Print Time", "PC-ABS BLK", "TPU 92A - Black", "QSR support", "Company Name"]
            df = df[required_columns]
            
            # Combine the data
            combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    # Save the combined data
    if not combined_df.empty:
        combined_df.to_csv(output_file, index=False)
        print(f"Combined dataset saved to: {output_file}")
    else:
        print("No valid files processed. No output generated.")
    
    return combined_df

# File paths
input_folder = os.path.dirname(os.path.abspath(__file__))  # The current script's directory
final_output_path = os.path.join(input_folder, "Appended_Print_F370.csv")

# Generate the final combined file
try:
    combined_data = process_f370_files(input_folder, final_output_path)
except Exception as e:
    print(f"Error during processing: {e}")
