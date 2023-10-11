import pandas as pd
import os

def extract_data_from_file(file_path):
    # Extracting identifier from filename
    identifier = os.path.basename(file_path).split('_')[0]
    
    # Reading the file using pandas and skipping the first 9 lines
    df = pd.read_csv(file_path, skiprows=9)
    
    # Extracting values from column 'Total Cashflow' and ignore empty cells
    values = df['Total Cashflow'].dropna().values[1:]  # Skipping the "Total Cashflow" header
    
    return identifier, values

def gather_data_from_all_files(directory):
    # Create an empty DataFrame to store all the data
    final_df = pd.DataFrame()
    
    # Iterate over all files in the directory with the specified pattern
    for file_name in os.listdir(directory):
        if file_name.startswith("EPAAG6") and file_name.endswith(".csv"):
            identifier, values = extract_data_from_file(os.path.join(directory, file_name))
            final_df[identifier] = pd.Series(values)
    
    return final_df

# Provide the path to the directory where your files are located
directory_path = "Files/AG38_V6_S130/"  # Replace with your directory path

# Gather data from all matching files in the directory
final_data = gather_data_from_all_files(directory_path)

# Save the gathered data to a new Excel file
output_path = os.path.join(directory_path, "consolidated_data.xlsx")
final_data.to_excel(output_path, index=False)
print(f"Data saved to: {output_path}")
