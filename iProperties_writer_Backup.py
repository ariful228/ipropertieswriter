import win32com.client
import csv
import logging

# Configure the logging module
logging.basicConfig(filename='processing_log.txt', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

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

        # Save the document to apply changes
        document.Save()

        # Close the document
        document.Close()

        # Log success
        logging.info(f"Success processing {file_path}")

    except win32com.client.pywintypes.com_error as e:
        # Handle COM errors
        if e.hresult == -2147418113:  # Error code for 'Object is not found'
            logging.warning(f"Warning: Object not found in {file_path}")
        else:
            logging.error(f"Error processing {file_path}: {e}")
    except Exception as e:
        # Log general errors
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
            if file_path.lower().endswith((".iam", ".ipt")):
                set_iproperties(file_path, attributes)

if __name__ == "__main__":
    main()
