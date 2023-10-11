import pandas as pd
import os

def extract_data_from_file(file_path):
    identifier = os.path.basename(file_path).split('_')[0]
    
    df = pd.read_csv(file_path, skiprows=9)
    
    values = df['Total Cashflow'].dropna().values[1:]  # Skipping the "Total Cashflow" header
    
    return identifier, values

def gather_data_from_all_files(directory):
    final_df = pd.DataFrame()
    
    for file_name in os.listdir(directory):
        if file_name.startswith("EPAAG6") and file_name.endswith(".csv"):
            identifier, values = extract_data_from_file(os.path.join(directory, file_name))
            final_df[identifier] = pd.Series(values)
    
    final_df['Total'] = final_df.sum(axis=1)
    
    return final_df

input_directory_path = "Files/AG38_V6_S130/"  # Replace with your directory path where the files are located

output_directory_path = "Files"  # Replace with your desired output directory path

final_data = gather_data_from_all_files(input_directory_path)

output_path = os.path.join(output_directory_path, "consolidated_data.xlsx")
final_data.to_excel(output_path, index=False)
print(f"Data saved to: {output_path}")
