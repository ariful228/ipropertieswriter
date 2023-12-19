import win32com.client
import csv
import pygetwindow as gw
import pyautogui
import time
import logging

# Configure the logging module
logging.basicConfig(filename='logfile.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def set_iproperties(file_path, properties):
    try:
        # Create an instance of Inventor
        inventor = win32com.client.Dispatch("Inventor.Application")

        try:
            # Open the Inventor document
            document = inventor.Documents.Open(file_path)
            popup = gw.getWindowsWithTitle("Resolve Link")[0]
            popup.activate()
            time.sleep(1)  # Add a delay to ensure the window is active
            pyautogui.press('Cancel')  # Press 'Cancel' to skip
        except IndexError:
            logging.error("Error: Unable to find 'Skip All' window.")
            return

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
            logging.info(f"Setting {prop_name} to {prop_value} in {file_path}")
            print(f"Setting {prop_name} to {prop_value} in {file_path}")

        # Set SilentOperation to True before saving to suppress the confirmation dialog
        inventor.SilentOperation = True

        try:
            # Save the iProperties changes
            iProperties.Save()
            logging.info(f"Saved iProperties changes in {file_path}")
        except win32com.client.pywintypes.com_error as e:
            # Handle COM errors
            if e.hresult == -2147418113:  # Error code for 'Object is not found'
                logging.warning(f"Warning: Object not found in {file_path}")
            else:
                logging.error(f"Error processing {file_path}: {e}")
                # Simulate pressing "OK" and then "Enter" for other types of pop-ups
                pyautogui.press('Enter')

        # Reset SilentOperation to False after saving
        inventor.SilentOperation = False

        # Close the document without saving changes
        document.Close(False)
        logging.info(f"Closed {file_path}")

    except Exception as e:
        logging.error(f"Error processing {file_path}: {e}")

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
                logging.info(f"Processing {file_path}...")
                set_iproperties(file_path, attributes)
                logging.info(f"Finished processing {file_path}")
                print(f"Finished processing {file_path}")

    # Close the Inventor application
    # inventor.Quit()

if __name__ == "__main__":
    main()
