import os
import glob
import openpyxl
import pandas as pd
import re

from pathlib import Path
from openpyxl.styles import Font

# Define the pre-existing folder locations as a dictionary
folder_locations = {
    1: "/CAE/01_Daimler-Truck",
    2: "/CAE/02_Daimler_Van",
    3: "/CAE/03_Daimler-EvoBus",
    4: "/CAE/30_Stihl",
    5: "/CAE"
}

# Define the pre-defined file types to choose from
file_types = {
    1: "*.odb",
    2: "*.bof",
    3: "*.pdf"
}

# Initialize an empty DataFrame to store the results
result_df = pd.DataFrame(columns=["File Name", "Author", "Size", "File Path"])

# Initalizing an empty data frame to store generated file extension
file_ext = ""
# Initalizing an empty data frame to store generated folder key
final_key = ""

while True:
    # Ask the user to choose a folder location
    def get_valid_folder_choice():
        while True:
            # Ask the user to choose a folder location
            print("Choose a folder location:")
            for key, value in folder_locations.items():
                print(f"{key}: {value}")
    
            folder_choice = input("Enter the number of your choice or enter 'Stop' to finsh process: ")
            
            if folder_choice.lower() == 'stop':
                return None, ''  # Return None for both values when the user enters 'stop'
                
            try:
                folder_choice = int(folder_choice)  # Convert the input to an integer
            except ValueError:
                print("Invalid choice. Please enter a valid number.")
                continue
            
            selected_folder = folder_locations.get(folder_choice)
            
            if selected_folder:
                return selected_folder, folder_choice
            else:
                print("Invalid choice. Please try again.")

    selected_folder, folder_choice = get_valid_folder_choice()   
    
    if selected_folder is None:
        break  # User entered 'stop', so exit the loop
              
    # Ask the user to choose a file type
    def get_valid_file_type_choice():
        while True:
            # Ask the user to choose a file type
            print("Choose a file type:")
            for key, value in file_types.items():
                print(f"{key}: {value}")
    
            file_type_choice = input("Enter the number of your choice: ")
            
            try:
                file_type_choice = int(file_type_choice)  # Convert the input to an integer
            except ValueError:
                print("Invalid choice. Please enter a valid number.")
                continue
            
            selected_file_type = file_types.get(file_type_choice)
    
            if selected_file_type:
                return selected_file_type
                
            else:
                print("Invalid choice. Please try again.")
            
    selected_file_type = get_valid_file_type_choice()
    file_ext += str(selected_file_type) + "," # Append the selected file type to the list
    final_key += str(folder_choice) + ","
    # Create a list to store the matching file paths
    matching_files = []
    
    # Use pathlib to search for matching files
    folder_path = Path(selected_folder)
    matching_files = folder_path.glob('**/' + selected_file_type)
    
    # Function to get file author
    def get_file_author(file_path):
        try:
            # Use a platform-specific method to get file author (example for Linux)
            author = os.popen(f"stat -c '%U' '{file_path}'").read().strip()
            return author
        except:
            return "N/A"
    
    # Function to get file size
    def get_file_size(file_path):
        try:
            # Check if the file exists
            if not os.path.exists(file_path):
               return f'File not found- {file_path}' # Zero file size is assigned for invalid file types
            size = os.path.getsize(file_path)
            return size
        except Exception as e:
            print(f"Error getting file size: {e}")
            return "N/A"
    
    # Create a list to store the file information
    file_info = []
    
    # Extract and store file information in the file_info list
    for file_path in matching_files:
        file_name = file_path.name
        file_author = get_file_author(file_path)
        file_size = get_file_size(file_path)
        
        file_info.append((file_name, file_author, file_size, str(file_path)))
     
        # Create a Pandas DataFrame from the file_info list
    df0 = pd.DataFrame(file_info, columns=["File Name", "Author", "Size", "File Path"])
    
    # Concatenate the current DataFrame with the result_df
    result_df = pd.concat([result_df, df0], ignore_index=True)

# Save the Excel file with both selected file type and selected folder in the filename
file_ext = file_ext.replace("*", "")  # Remove asterisk from the file type
 
print(f"Found {len(file_info)} matching files.")

# Function to clean and convert the string in 'Size' column
def process_size_column(df):
    for index, row in result_df.iterrows():
        size_value = row['Size']
        if isinstance(size_value, str):
            print(f"Found a string value in row {index}: {size_value}")
            result_df.at[index, 'Size'] = 0
    return result_df

# Define a dictionary to map unit abbreviations to multipliers (e.g., KB to 1024)
unit_multipliers = {
    'B': 1,
    'KB': 1024,
    'MB': 1024 ** 2,
    'GB': 1024 ** 3,
    'TB': 1024 ** 4,
}

while True:
    user_input = input("Enter the file size limit (e.g., 100B, 1KB, 10MB, 5GB): ").strip().upper()
    
    # Use regular expressions to parse the input for numeric value and unit
    match = re.match(r'^(\d+)([A-Z]+)$', user_input)
    
    if match:
        size_value, size_unit = match.groups()
        try:
            size_value = float(size_value)
            size_multiplier = unit_multipliers.get(size_unit, None)
            if size_multiplier:
                size_limit = size_value * size_multiplier
                break  # Exit the loop if the input is valid
            else:
                print("Invalid unit. Please use B, KB, MB, GB, or TB.")
        except ValueError:
            print("Invalid input. Please enter a valid number.")
    else:
        print("Invalid input format. Please use the format: 100B, 1KB, 10MB, etc.")
        
df2 = process_size_column(result_df)

df2 = df2[df2['Size'] >= size_limit]  # Filter based on the numeric 'Size'

def convert_size_column(dataframe):
    for index, row in dataframe.iterrows():
        size_value = row['Size']
        if isinstance(size_value, (int, float)):
            if size_value >= 1e9:
                dataframe.at[index, 'Size'] = f"{size_value / 1e9:.2f}GB"
            elif size_value >= 1e6:
                dataframe.at[index, 'Size'] = f"{size_value / 1e6:.2f}MB"
            elif size_value >= 1e3:
                dataframe.at[index, 'Size'] = f"{size_value / 1e3:.2f}KB"
            else:
                dataframe.at[index, 'Size'] = f"{size_value:.2f}bytes"
    return dataframe

#Sort the DataFrame by file size in descending order
df2 = df2.sort_values(by='Size', ascending=False)

# Convert the 'Size' column
df3 = convert_size_column(df2)

# Create a Pandas Excel writer object for saving the updated data
updated_excel_filename = f'Evaluated_Files:{file_ext} Search_folders:{final_key} Filter_size:{size_value}{size_unit}.xlsx'

with pd.ExcelWriter(updated_excel_filename, engine='openpyxl') as writer:
    df3.to_excel(writer, sheet_name='All Users', index=False)
   
    # Group the data by author name and save each group to a separate sheet
    grouped = df3.groupby('Author')
    print(grouped)
    for author, group_data in grouped:
        sheet_name = f'{author}'
        group_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
    # Get the openpyxl workbook object and save it
    workbook = writer.book
    workbook.save(updated_excel_filename)
    
print(f"Data has been segregated by author and saved to '{updated_excel_filename}'.")