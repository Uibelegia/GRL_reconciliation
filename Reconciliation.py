import pandas as pd
import os


def reconcile_excel_files():
    try:
        # Step 1: Get file and sheet inputs with validation
        folder_path = r'C:\Users\18366\Downloads'

        file1_name = input("Enter the GRL file name (without .xlsx): ").strip()
        if not file1_name:
            print("GRL file name cannot be empty. Exiting.")
            return
        file1_path = os.path.join(folder_path, f"{file1_name}.xlsx")
        if not os.path.exists(file1_path):
            print(f"File {file1_path} does not exist. Exiting.")
            return

        sheet1_name = input("Enter the GRL sheet name: ").strip()
        if not sheet1_name:
            print("GRL sheet name cannot be empty. Exiting.")
            return
        try:
            df_GRL = pd.read_excel(file1_path, sheet_name=sheet1_name, header=0)
        except Exception as e:
            print(f"Failed to read sheet '{sheet1_name}' from GRL file. Error: {e}. Exiting.")
            return

        GRL_name = input("Enter a name for GRL: ").strip()
        if not GRL_name:
            print("GRL name cannot be empty. Exiting.")
            return
        GRL_Quantity = GRL_name + ' Quantity'

        file2_name = input("Enter the Other file name (without .xlsx): ").strip()
        if not file2_name:
            print("Other file name cannot be empty. Exiting.")
            return
        file2_path = os.path.join(folder_path, f"{file2_name}.xlsx")
        if not os.path.exists(file2_path):
            print(f"File {file2_path} does not exist. Exiting.")
            return

        sheet2_name = input("Enter the Other sheet name: ").strip()
        if not sheet2_name:
            print("Other sheet name cannot be empty. Exiting.")
            return
        try:
            df_Other = pd.read_excel(file2_path, sheet_name=sheet2_name, header=0)
        except Exception as e:
            print(f"Failed to read sheet '{sheet2_name}' from Other file. Error: {e}. Exiting.")
            return

        Other_name = input("Enter a name for Other: ").strip()
        if not Other_name:
            print("Other name cannot be empty. Exiting.")
            return
        Other_Quantity = Other_name + ' Quantity'

        # Step 2: Perform reconciliation
        df_reconciliation = pd.merge(
            df_Other[['Code', 'Quantity']].rename(columns={'Quantity': Other_Quantity}),
            df_GRL[['Code', 'Quantity']].rename(columns={'Quantity': GRL_Quantity}),
            on='Code',
            how='inner'
        )

        df_reconciliation['Diff'] = df_reconciliation[Other_Quantity] - df_reconciliation[GRL_Quantity]
        df_diff = df_reconciliation[df_reconciliation['Diff'] != 0]

        df_GRL_only = df_GRL[~df_GRL['Code'].isin(df_Other['Code'])].copy()
        df_Other_only = df_Other[~df_Other['Code'].isin(df_GRL['Code'])].copy()

        # Step 3: Write results to an Excel file
        output_file = os.path.join(folder_path, f"{GRL_name}_{Other_name}_reconciliation.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_reconciliation.to_excel(writer, sheet_name='Reconciliation', index=False)
            df_diff.to_excel(writer, sheet_name='Difference', index=False)
            df_GRL_only.to_excel(writer, sheet_name=f'{GRL_name} Only', index=False)
            df_Other_only.to_excel(writer, sheet_name=f'{Other_name} Only', index=False)

        print(f"Reconciliation completed. Results saved to {output_file}")

    except Exception as e:
        print(f"An error occurred: {e}")



reconcile_excel_files()





