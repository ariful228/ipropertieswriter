import pandas as pd

# Specify the input CSV and output Excel file paths
input_csv_path = r"D:\PYDATAANALYSIS\iPropertiesProcessdata\export_dump_Monday_18_12_2023_13_05_PM.csv"
output_excel_path = "export_dump_process.xlsx"

# Read the CSV file into a DataFrame using the correct separator (pipe '|')
df = pd.read_csv(input_csv_path, sep='|')

# Write the DataFrame to an Excel file
df.to_excel(output_excel_path, index=False)

print(f"CSV to Excel conversion completed. Output saved to {output_excel_path}")
