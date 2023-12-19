import win32com.client
import csv
import pyautogui

def set_iproperties(file_path, properties):
    try:
        # Create an instance of Inventor
        inventor = win32com.client.Dispatch("Inventor.Application")

        # Open the Inventor document
        document = inventor.Documents.Open(file_path)

        # Get the iProperties collection
        iProperties = document.PropertySets.Item("Inventor User Defined Properties")

        # Set each specified property
        for prop_name, prop_value in properties.items():
            try:
                prop = iProperties[prop_name]
            except:
                # If the property doesn't exist, create it
                prop = iProperties.Add(prop_name, prop_value)

            # Set the value of the property
            prop.Value = prop_value
            print(f"Setting {prop_name} to {prop_value} in {file_path}")

        # Set SilentOperation to True before saving to suppress the confirmation dialog
        inventor.SilentOperation = True

        try:
            # Save the document to apply changes
            document.Save()
            print(f"Saved changes in {file_path}")
        except win32com.client.pywintypes.com_error as e:
            # Handle COM errors
            if e.hresult == -2147418113:  # Error code for 'Object is not found'
                print(f"Warning: Object not found in {file_path}")

                # Simulate pressing "Skip" and then "Enter"
                pyautogui.press('Skip All')  # Replace 's' with the key for "Skip" if available
                pyautogui.press('enter')

            else:
                print(f"Error processing {file_path}: {e}")
                
                # Simulate pressing "OK" and then "Enter" for other types of pop-ups
                pyautogui.press('enter')

        # Reset SilentOperation to False after saving
        inventor.SilentOperation = False

        # Close the document
        document.Close()

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

def main():
    # Read the CSV file with file paths and attribute names
    csv_file_path = "input.csv"

    with open(csv_file_path, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for row in csv_reader:
            file_path = row["Assembly_Path"]  # Change this according to your CSV header
            attributes = {key: row[key] for key in row.keys() if key != "Assembly_Path"}

            # Check file extension
            if file_path.lower().endswith((".iam", ".ipt", ".idw", ".dwg", ".ipn")):
                print(f"Processing {file_path}...")
                set_iproperties(file_path, attributes)
                print(f"Finished processing {file_path}")

if __name__ == "__main__":
    main()
