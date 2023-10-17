import pandas as pd
import os

def extract_data_from_file(file_path):
    identifier = os.path.basename(file_path).split('_')[0]
    df = pd.read_csv(file_path, skiprows=9)
    values = df['Total Cashflow'].values[1:]  # Skipping the "Total Cashflow" header
    return identifier, values

def gather_data_from_all_files(directory):
    # Initialize an empty list to store all extracted values
    all_values = []
    
    # Extract data from all files
    for file_name in os.listdir(directory):
        if file_name.startswith("EPAFH2") and file_name.endswith(".csv"):
            identifier, values = extract_data_from_file(os.path.join(directory, file_name))
            all_values.append((identifier, values))
            
    # Determine the maximum length of values
    max_length = max(len(values) for _, values in all_values)
    
    # Initialize final_df with the correct number of rows
    final_df = pd.DataFrame(index=range(max_length))
    
    # Assign the extracted values to final_df
    for identifier, values in all_values:
        final_df[identifier] = pd.Series(values)
    
    # Fill NaN values with zeros
    final_df.fillna(0, inplace=True)
    
    # Calculate the total for each row
    final_df['Total'] = final_df.sum(axis=1)
    
    return final_df

input_directory_path = "Files/AG38_V9_S130/"  # Replace with your directory path where the files are located
output_directory_path = "Files"  # Replace with your desired output directory path
final_data = gather_data_from_all_files(input_directory_path)
output_path = os.path.join(output_directory_path, "consolidated_data.xlsx")

final_data.to_excel(output_path, index=False)

print(f"Data saved to: {output_path}")
