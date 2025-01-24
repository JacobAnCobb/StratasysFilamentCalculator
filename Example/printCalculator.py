import pandas as pd
import os
import re

# Paths to input files
f370_file = "Appended_Print_F370.csv"
j826_file = "Appended_Print_J826.csv"
material_cost_file = "Material_Cost.txt"
output_file = "Print_Cost_Analysis.csv"

def parse_material_cost(file_path):
    """
    Parses the Material_Cost.txt file to extract costs and volumes for each material.
    """
    costs = {}
    with open(file_path, "r") as file:
        for line in file:
            if "**" in line or line.strip() == "":
                continue  # Ignore lines with ** or empty lines
            
            match = re.match(r"(.*?); (.*?); (.*?); \$(.*)$", line.strip())
            if match:
                printer, material, volume, cost = match.groups()
                volume = re.sub(r"[^\d.]", "", volume)  # Remove non-numeric characters
                volume = float(volume) if volume else None
                cost = float(cost.replace(",", ""))  # Convert cost to float
                costs[(printer.strip(), material.strip())] = {
                    "volume": volume,
                    "cost": cost
                }
    return costs

def calculate_costs(f370_file, j826_file, material_cost_file, output_file):
    """
    Reads the input files, calculates costs, and generates the output CSV.
    """
    # Parse the material costs
    material_costs = parse_material_cost(material_cost_file)

    # Read the input CSV files
    f370_data = pd.read_csv(f370_file)
    j826_data = pd.read_csv(j826_file)

    results = []

    # Process F370 prints
    for idx, row in f370_data.iterrows():
        printer = "F370"
        company_name = row["Company Name"]
        flat_fee = 16.25  # Flat fee for F370
        filament_fee = 0

        # Calculate costs for ABS, TPU, and Support
        for material, col_name in [("PC-ABS BLK", "ABS in cm3"), ("TPU 92A - Black", "TPU 92A - Black"), ("QSR support", "QSR support")]:
            if col_name in row:
                volume = row[col_name]
                if pd.notna(volume) and volume > 0:
                    key = (printer, material)
                    if key in material_costs and material_costs[key]["volume"]:
                        cost_per_unit = material_costs[key]["cost"] / material_costs[key]["volume"]
                        filament_fee += volume * cost_per_unit

        total_fee = filament_fee + flat_fee
        results.append({
            "Print Job": f"Print Job {idx + 1}",
            "Printer Type": printer,
            "Filament Fee": round(filament_fee, 2),
            "Flat Fee": flat_fee,
            "Total Fee": round(total_fee, 2),
            "Company Name": company_name
        })

    # Process J826 prints
    for idx, row in j826_data.iterrows():
        printer = "J826"
        company_name = row["Company Name"]
        flat_fee = 7.22  # Flat fee for J826
        filament_fee = 0

        # Calculate costs for materials
        for material, col_name in [("DraftGrey", "DraftGrey (g)"), 
                                   ("VeroUltraWhite", "VeroUltraWhite (g)"), 
                                   ("VeroBlackPlus", "VeroBlackPlus (g)"), 
                                   ("SUP706", "SUP706 (g)")]:
            if col_name in row:
                volume = row[col_name]
                if pd.notna(volume) and volume > 0:
                    key = (printer, material)
                    if key in material_costs and material_costs[key]["volume"]:
                        cost_per_unit = material_costs[key]["cost"] / material_costs[key]["volume"]
                        filament_fee += volume * cost_per_unit

        total_fee = filament_fee + flat_fee
        results.append({
            "Print Job": f"Print Job {len(results) + 1}",
            "Printer Type": printer,
            "Filament Fee": round(filament_fee, 2),
            "Flat Fee": flat_fee,
            "Total Fee": round(total_fee, 2),
            "Company Name": company_name
        })

    # Create a DataFrame for the results
    results_df = pd.DataFrame(results)

    # Save the results to a CSV file
    results_df.to_csv(output_file, index=False)
    print(f"Cost analysis saved to: {output_file}")

# Run the calculation
calculate_costs(f370_file, j826_file, material_cost_file, output_file)
