import pandas as pd

def process_csv(input_path, output_path):
    try:
        # Read the CSV file into a DataFrame, skipping lines with parsing errors
        df = pd.read_csv(input_path, sep='|')

        # Keep only the specified columns
        columns_to_keep = ["FilePath", "PRODHIERARCHY", "MATGROUP", "MATERIALTYPE", "EXTMATGROUP", "LAB_OFFICE", "HDMATWT"]
        df = df[columns_to_keep]

        # Process PRODHIERARCHY, MATGROUP, MATERIALTYPE, EXTMATGROUP, and LAB_OFFICE columns
        for col in ["PRODHIERARCHY", "MATGROUP", "MATERIALTYPE", "EXTMATGROUP", "LAB_OFFICE"]:
            df[col] = df[col].str.split(" - ").str[0]

        # Process HDMATWT column
        df["HDMATWT"] = df["HDMATWT"].str.replace(",", ".").str.replace("kg", "").str.strip()

        # Convert HDMATWT to float
        df["HDMATWT"] = pd.to_numeric(df["HDMATWT"], errors='coerce')

        # Rename FilePath to Assembly_Path
        df = df.rename(columns={"FilePath": "Assembly_Path"})

        # Replace specific values with empty string
        values_to_replace = ["not_found", "MATGROUP doest not exist", "EXTMATGROUP doest not exist",
                             "LAB_OFFICE doest not exist", "PRODHIERARCHY doest not exist",
                             "MATERIALTYPE doest not exist", "HDMATWT  doest not exist"]
        df.replace(values_to_replace, "", inplace=True)

        # Export the modified DataFrame to a new CSV file
        df.to_csv(output_path, sep="|", index=False)

        print(f"CSV processing completed. Output saved to {output_path}")

    except pd.errors.ParserError as e:
        print(f"Error parsing the CSV file: {e}")
        print("Check the content of the CSV file for inconsistencies.")

# Specify the input and output file paths using raw string literals
input_csv_path = "export_dump_process.csv"
output_csv_path = "input.csv"

# Call the function to process and export the CSV
process_csv(input_csv_path, output_csv_path)
