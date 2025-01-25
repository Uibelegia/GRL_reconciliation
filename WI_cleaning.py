import numpy as np
import pandas as pd
import re
import os


# Import raw data from the downloaded QuickBooks Excel sheet
df = pd.read_excel(
    r"C:\Users\18366\Downloads\Durasein LDSS Inventory Sheet.xlsx",
    sheet_name="Sheet1",
    header=2
)




# Extract the required columns, create a copy, and store it in a new DataFrame
cleaned_df = df[['Type', 'Code', 'Size', 'Ending Qty']].copy()

# Rename DataFrame columns for further processing
cleaned_df.columns = ['Type', 'Code', 'Description', 'Quantity']



# Remove the last row as it is not needed
new_df = cleaned_df[:-1].copy()



# Remove quotes and leading/trailing spaces
new_df['Description'] = new_df['Description'].str.replace(r"[\"']", '', regex=True).str.strip()

# Remove leading/trailing spaces from 'Type'
new_df['Type'] = new_df['Type'].str.strip()

# Remove leading/trailing spaces from 'Code'
new_df['Code'] = new_df['Code'].str.strip() 




# Fill missing values in the 'Description' column with 'Sinks'
new_df['Description'] = new_df['Description'].fillna('Sinks')



# Assign 'Acrylic Sheet' to the 'Product Family' column for rows where 'Type' contains 'Durasein'
new_df.loc[new_df['Type'].str.contains('Durasein', case=False, na=False), 'Product Family'] = 'Acrylic Sheet'




# Assign 'Corner Caddy' to the 'Product Family'
new_df.loc[
    new_df['Description'].str.contains(
        'Corner Caddy',
        case=False, na=False
    ),
    'Product Family'
] = 'Corner Caddy'


# In[9]:


# Assign 'Soap Dish' to the 'Product Family'
new_df.loc[
    new_df['Description'].str.contains(
        'Corner Soap Dish',
        case=False, na=False
    ),
    'Product Family'
] = 'Soap Dish'



# Assign 'Shower Trim' to the 'Product Family'
new_df.loc[
    new_df['Description'].str.contains(
        'Dado Trim|Corner Cove|L-Trim',
        case=False, na=False
    ),
    'Product Family'
] = 'Shower Trim'



# Assign 'Shower Pan' to the 'Product Family'
new_df.loc[
    new_df['Description'].str.contains(
        '32x 60|36x 36|36x 48',
        case=False, na=False
    ),
    'Product Family'
] = 'Shower Pan'



# Assign 'Recessed Caddies' to the 'Product Family'
new_df.loc[
    new_df['Description'].str.contains(
        '7-1/2x 11|8x 18 w/shelf',
        case=False, na=False
    ),
    'Product Family'
] = 'Recessed Caddies'



# Fill missing values in the 'Product Family' column with 'Acrylic Sinks'
new_df['Product Family'] = new_df['Product Family'].fillna('Acrylic Sinks')



def WI_code_suffix(df):
    # If Product Family is 'Acrylic Sheet' or 'Salvaged Misc Sizes'
    if df['Product Family'] in ['Acrylic Sheet', 'Salvaged Misc Sizes']:
        if pd.notna(df['Description']):  # Ensure 'Description' is not NaN
            # Extract all digits and remove the first digit
            numbers = ''.join([c for c in df['Description'] if c.isdigit()])
            return numbers[1:] if len(numbers) > 1 else None

    # Handle specific 'Description' cases
    elif df['Description'] == 'Corner Caddy':
        return 'CC'
    elif df['Description'] == 'Recessed Caddy':
        return 'RC'
    elif df['Description'] == 'L-Trim':
        return 'LTRIM'
    elif df['Description'] == 'Corner Cove':
        return 'COVE'
    elif df['Description'] == 'Dado Trim':
        return 'DADO'
    elif df['Description'] == '7-1/2x 11':
        return '7.5x11-RC'
    elif df['Description'] == '8x 18 w/shelf':
        return '8x18-RC'
    elif df['Description'] == 'Corner Soap Dish':
        return 'CSD'

    # For other cases, return 'Description' directly if it's not NaN
    return df['Description'] if pd.notna(df['Description']) else None



# Apply the WI_code_suffix function to each row in the DataFrame
new_df['code_suffix'] = new_df.apply(WI_code_suffix, axis=1)



def WI_combine_code(df):
    # Combine 'Type' and 'Code' for 'Acrylic Sinks'
    if df['Product Family'] == 'Acrylic Sinks':
        return df['Type'] + '-' + df['Code']
    
    # Add 'G-' prefix for 'Shower Pan' and combine with 'Type' and 'Code'
    elif df['Product Family'] == 'Shower Pan':
        return 'G-' + df['Type'] + '-' + df['Code']
    
    # For other cases, combine 'Code' with 'code_suffix' if it exists
    else:
        return df['Code'] + '-' + (df['code_suffix'] if pd.notna(df['code_suffix']) else '')



# Apply the WI_combine_code function
new_df['Code'] = new_df.apply(WI_combine_code, axis=1)



# Select only 'Code' and 'Quantity' columns for easy comparison
formatted_df = new_df[['Code', 'Quantity']]



# Output DataFrame to be an Excel File
folder_path = r'C:\Users\18366\Downloads'
file_name = '3.xlsx'

output_file = os.path.join(folder_path, file_name)
formatted_df.to_excel(output_file, index=False)