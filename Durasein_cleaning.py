import numpy as np
import pandas as pd
import re
import os


# Import raw data from the downloaded QuickBooks Excel sheet
df= pd.read_excel(r"C:\Users\18366\Downloads\Book2_WI.xlsx",
                  sheet_name = "Sheet1",
                  header = 1)



# Function to remove NaN values
def remove_nan(my_list):
    return [[item for item in row if pd.notna(item)] for row in my_list]



# Convert DataFrame to a list for easier removal of NaN values
list_df = df.values.tolist()

# Apply the remove_nan() function
cleaned_df = remove_nan(list_df)



# Convert the list back to a DataFrame for further cleaning and analysis
cleaned_df = pd.DataFrame(cleaned_df)



# Delete the third column if there is one
if cleaned_df.shape[1] == 3:
    cleaned_df = cleaned_df.iloc[:, :-1]
    


# Rename DataFrame columns for further processing
cleaned_df.columns = ['Product', 'Quantity']



# Conditionally clean the new DataFrame
# Remove rows that do not align with individual item criteria
new_df = cleaned_df.dropna(subset=['Quantity'])
new_df = new_df[~new_df['Product'].str.contains('Total|Other', case=False, na=False)]



# Create Code and Description columns
# Separate content inside and outside parentheses and store in 'Code' and 'Description'

# Extract strings outside parentheses from the 'Product' column and save to 'Code'
new_df['Code'] = new_df['Product'].str.extract(r'^(.*?)\s*\(')

# Extract strings inside parentheses from the 'Product' column and save to 'Description'
new_df['Description'] = new_df['Product'].str.extract(r'\((.*?)\)$')



# Description Cleaning

# Clean 'Description' by removing extra parentheses content
new_df['Description'] = new_df['Description'].str.replace(r'\s*\(.*?\)', '', regex=True)



# WI_warehouse_code_merge Function to Merge two items together

def WI_warehouse_code_merge(df, code_pattern, grouped_suffix):
    code = 'Code'
    quantity = 'Quantity'
    
    # Mark the rows that need to be processed
    df['Marked'] = df[code].str.contains(code_pattern)
    
    # Extract the grouped main code (only process rows where Marked=True)
    df['Grouped_Code'] = df.apply(
        lambda row: row[code][:-2] + grouped_suffix if row['Marked'] else row[code],
        axis=1
    )
    
    # Summarize Quantity by Grouped_Code
    grouped = df.groupby('Grouped_Code')[quantity].sum().reset_index()
    
    # Merge the summarized results back into the original DataFrame
    df = df.merge(grouped, on='Grouped_Code', how='left', suffixes=('', '_Summed'))
    
    # Update the original DataFrame with the summarized quantity and updated code
    df[code] = df['Grouped_Code']
    df[quantity] = df[f"{quantity}_Summed"]
    
    # Remove duplicates, keeping only the last record for each group
    df = df.drop_duplicates(subset=[code], keep='last').reset_index(drop=True)
    
    return df


# Apply the WI_warehouse_code_merge function to merge two similar product codes into one item
formatted_df = WI_warehouse_code_merge(new_df, code_pattern='4369|4609', grouped_suffix='96')



# Select only 'Code' and 'Quantity' columns for easy comparison
formatted_df = formatted_df[['Code','Quantity']]



# Output DataFrame to be an Excel File
folder_path = r'C:\Users\18366\Downloads'
file_name = '1.xlsx'

output_file = os.path.join(folder_path, file_name)
formatted_df.to_excel(output_file, index=False)

