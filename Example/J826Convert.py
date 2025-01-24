import pandas as pd
import os

def convert_j826_file(input_file, output_file, is_first_file):
    """
    Converts a J826 file into the desired format with consistent columns and appends the output.
    
    Args:
        input_file (str): Path to the input J826 file.
        output_file (str): Path to the output CSV file.
        is_first_file (bool): Whether this is the first file being processed (to include headers).
    """
    # Read the CSV file
    df = pd.read_csv(input_file, header=None, names=['Key', 'Value'])

    # Extract the company name from row 15 in the first column
    try:
        company_name = df.iloc[14, 0]  # Row 15 is 14 (0-based index)
        company_name = str(company_name).strip() if pd.notna(company_name) else "Unknown Company"
    except IndexError:
        print(f"Warning: Row 15 is missing in {input_file}. Assigning 'Unknown Company'.")
        company_name = "Unknown Company"

    # Restructure alternating rows into key-value pairs
    keys = df['Key'][::2].reset_index(drop=True)  # Odd rows (keys)
    values = df['Key'][1::2].reset_index(drop=True)  # Even rows (values)
    data = dict(zip(keys, values))

    # Extract and process necessary values
    print_time = data.get('Print Time', '0h 0m').strip()  # Example: '3h 55m'
    draft_grey = data.get('DraftGrey (g)', '0').strip()
    vero_ultra_white = data.get('VeroUltraWhite (g)', '0').strip()
    vero_black_plus = data.get('VeroBlackPlus (g)', '0').strip()
    sup706 = data.get('SUP706 (g)', '0').strip()

    # Safely handle numeric conversions
    def safe_float(value):
        try:
            return float(value)
        except ValueError:
            return 0

    draft_grey = safe_float(draft_grey)
    vero_ultra_white = safe_float(vero_ultra_white)
    vero_black_plus = safe_float(vero_black_plus)
    sup706 = safe_float(sup706)

    # Format the print time to h:mm
    if 'h' in print_time and 'm' in print_time:
        print_time_split = print_time.replace('h', ':').replace('m', '').split(':')
        formatted_time = f"{print_time_split[0]}:{print_time_split[1].zfill(2)}"
    else:
        formatted_time = "0:00"

    # Create a new DataFrame in the desired format
    output_df = pd.DataFrame({
        'Print Time (h:mm)': [formatted_time],
        'DraftGrey (g)': [draft_grey],
        'VeroUltraWhite (g)': [vero_ultra_white],
        'VeroBlackPlus (g)': [vero_black_plus],
        'SUP706 (g)': [sup706],
        'Company Name': [company_name]
    })

    # Save the new CSV file
    if os.path.exists(output_file) and not is_first_file:
        output_df.to_csv(output_file, mode='a', header=False, index=False)
    else:
        output_df.to_csv(output_file, index=False)
    print(f"Output file saved: {os.path.abspath(output_file)}")
    return output_df


def process_all_j826_files(folder_path):
    """
    Processes all J826 print job files in the specified folder and combines them into a single CSV.
    
    Args:
        folder_path (str): Path to the folder containing the J826 files.
    """
    found_files = False  # Track if any files are found
    print(f"Scanning directory: {folder_path}")

    # Define the appended output file
    appended_file = os.path.join(folder_path, "Appended_Print_J826.csv")
    is_first_file = not os.path.exists(appended_file)

    for file_name in os.listdir(folder_path):
        if "Print_J826" in file_name and file_name.endswith(".csv"):
            input_file = os.path.join(folder_path, file_name)
            print(f"Processing file: {file_name}")
            convert_j826_file(input_file, appended_file, is_first_file)
            found_files = True
            is_first_file = False

    if not found_files:
        print("No new files containing 'Print_J826' found in the directory.")

# Specify folder path
folder_path = os.getcwd()  # Current working directory

# Process all J826 files in the folder
process_all_j826_files(folder_path)

print("Processing completed.")
