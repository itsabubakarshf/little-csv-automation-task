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
    
    # Calculate the sum of non-empty cells horizontally for each row and add to a new "Total" column
    final_df['Total'] = final_df.sum(axis=1)
    
    return final_df

# Provide the path to the directory where your files are located
input_directory_path = "Files/AG38_V6_S130/"  # Replace with your directory path where the files are located

# Specify the location where the consolidated data should be saved
output_directory_path = "Files"  # Replace with your desired output directory path

# Gather data from all matching files in the input directory
final_data = gather_data_from_all_files(input_directory_path)

# Save the gathered data to the output location
output_path = os.path.join(output_directory_path, "consolidated_data.xlsx")
final_data.to_excel(output_path, index=False)
print(f"Data saved to: {output_path}")
