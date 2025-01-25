import numpy as np
import pandas as pd
import re
import os


# Import raw data from the downloaded QuickBooks Excel sheet
df = pd.read_excel(r"C:\Users\18366\Downloads\Inventory ending December 31, 2025.xlsx",
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



def classify_by_range(df, start_end_rules):
    # Initialize a list to store classification results
    product_family = []  
    current_category = None
    
    for product in df['Product']:
        product = product.strip().lower()  # Remove extra spaces and convert to lowercase
        
        # Check if the current product matches any range's start or end value
        for start, end, category in start_end_rules:
            if product == start.strip().lower():  # If it matches the start of a range
                current_category = category
            elif product == end.strip().lower():  # If it matches the end of a range
                current_category = category
        
        # Record the result based on the current range category
        product_family.append(current_category)
    
    # Add the classification results to a new column
    df['Product Family'] = product_family
    return df



rules = [
    ("Acrylic Sheets", "Total Acrylic Sheets", "Acrylic Sheet"),
    ("Sinks", "Total Sinks", "Acrylic Sinks"),
    ("Salvaged MIsc. Sizes", "Total Salvaged MIsc. Sizes", "Salvaged Misc Sizes"),
    ("Acrylic Stone Samples", "Total Acrylic Stone Samples", "Sample"),
    ("SHOWER ACCESSORIES", "Total SHOWER ACCESSORIES", "SHOWER ACCESSORIES"),
    ("CORNER CADDIES","Total CORNER CADDIES", "CORNER CADDIES"),
    ("RECESSED CADDIES", "Total RECESSED CADDIES", "RECESSED CADDIES"),
    ("SHOWER PAN", "Total SHOWER PAN", "SHOWER PAN"),
    ("SHOWER TRIM","Total SHOWER TRIM", "SHOWER TRIM"),
    ("SOAP DISH","Total SOAP DISH","SOAP DISH")
]


cleaned_df = classify_by_range(cleaned_df, rules)




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

# Remove single and double quotes from 'Description' (standard step)
new_df['Description'] = new_df['Description'].str.replace(r"[\"']", '', regex=True)

# Remove backticks (`) and hyphens (-) from 'Description' (additional step for specific cases like `-)
new_df['Description'] = new_df['Description'].str.replace(r"[\`-]", '', regex=True)

# Remove decimal points and the digit immediately following them from 'Description' (standard step)
new_df['Description'] = new_df['Description'].str.replace(r'\.\d', '', regex=True)



def process_code_suffix(df):
    # Extract Dimension
    dimension = df['Description']
    extracted_dim = None
    if pd.notna(dimension):
        match = pd.Series(dimension).str.extract(r'(\d.*\d)').iloc[0, 0]
        extracted_dim = match if pd.notna(match) else None

    # Logic based on Product Family
    if df['Product Family'] in ['Acrylic Sheet', 'Salvaged Misc Sizes']:
        if extracted_dim:  # Ensure Dimension exists
            # Extract all digits and remove the first digit
            numbers = ''.join([c for c in extracted_dim if c.isdigit()])
            return numbers[1:] if len(numbers) > 1 else None

    elif df['Product Family'] == 'Sample':
        if extracted_dim:
            return f"Sample_{extracted_dim}"

    elif df['Product Family'] == 'Acrylic Sinks':
        # For Acrylic Sinks, return "Acrylic Sinks" as a fixed value
        return "Acrylic Sinks"

    elif df['Product Family'] == 'SHOWER ACCESSORIES':
        # For SHOWER ACCESSORIES, check for specific keywords
        if "corner caddy" in dimension.lower():
            return 'CC'
        elif "recessed caddy" in dimension.lower():
            return 'RC'
        elif "trim" in dimension.lower():
            if "corner" in dimension.lower():
                return 'COVE'
            elif "dado" in dimension.lower():
                return 'DADO'
            else:
                return 'LTRIM'

    # Default to returning the processed Dimension
    return extracted_dim if extracted_dim else None



# Apply function process_code_suffix for creating product code suffix according to its product family
new_df['code_suffix'] = new_df.apply(process_code_suffix, axis=1)



def combine_code(df):
    # Process 'Acrylic Sheet' and 'Salvaged Misc Sizes'
    if df['Product Family'] in ['Acrylic Sheet', 'Salvaged Misc Sizes']:
        if pd.notna(df['Code']):
            if not df['Code'].startswith('SC'):
                # Remove all "-" and content after the first space
                code = df['Code'].split()[0].split('-')[0]
                code = ''.join(c.upper() if c.isalpha() else c for c in code)  # Convert all letters to uppercase
                # Combine Code with code_suffix
                return f"{code}-{df['code_suffix']}" if pd.notna(df['code_suffix']) else code

    # Process 'Sample'
    elif df['Product Family'] == 'Sample':
        if pd.notna(df['Code']):
            # Remove all "-" and content after the first space
            code = df['Code'].split()[0].split('-')[0]
            code = ''.join(c.upper() if c.isalpha() else c for c in code)  # Convert all letters to uppercase
            # Combine Code with Sample_code_suffix
            return f"{code} {df['code_suffix']}" if pd.notna(df['code_suffix']) else code

    # Process 'Acrylic Sinks'
    elif df['code_suffix'] in ['CC', 'RC', 'LTRIM', 'COVE', 'DADO']:
        if pd.notna(df['Code']):
            # Remove all "-" and content after the first space
            code = df['Code'].split()[0].split('-')[0]
            code = ''.join(c.upper() if c.isalpha() else c for c in code)  # Convert all letters to uppercase
            # Combine Code with code_suffix
            return f"{code}-{df['code_suffix']}" if pd.notna(df['code_suffix']) else code

    # Handle other cases
    elif pd.notna(df['Code']):
        # Replace "/" with "-" and remove spaces around "-"
        return df['Code'].replace('/', '-').replace(' - ', '-').replace(' -', '-').replace('- ', '-')

    # Retain original value
    return df['Code']



# Apply function combine_code to combine code and suffix code
new_df['Code'] = new_df.apply(combine_code, axis=1)



# Select only 'Code' and 'Quantity' columns for easy comparison
formatted_df = new_df[['Code','Quantity']]



# Output DataFrame to be an Excel
folder_path = r'C:\Users\18366\Downloads'
file_name = '2.xlsx'

output_file = os.path.join(folder_path, file_name)
formatted_df.to_excel(output_file, index=False)