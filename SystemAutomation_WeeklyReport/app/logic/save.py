import win32com.client as win32  # Import the win32com.client module to interact with Excel
import json
import os

# Class to help with file path management
class ResourceHelper:
    @staticmethod
    def get_path(relative_path: str) -> str:
        """
        Get the absolute path of a file relative to the current script's directory.

        Args:
            relative_path (str): The relative path to the file.

        Returns:
            str: The absolute path to the file.
        """
        # Get the directory of the current script and join it with the relative path
        base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

# Get the path to the JSON configuration file
file_path_json = ResourceHelper.get_path('../config/inputan.json')

# Read the JSON configuration file
with open(file_path_json, 'r') as file:
    json_data = json.load(file)

def main():
    # Extract the path to the final Excel file from the JSON data
    file_path = json_data["final_file"]
    
    # Start Excel application in the background
    excel = win32.DispatchEx("Excel.Application")  # Create a new instance of Excel
    excel.DisplayAlerts = False  # Disable prompts for actions like overwriting files
    excel.Visible = False        # Keep Excel application hidden from the user
    
    # Open the workbook specified in the JSON configuration
    wb = excel.Workbooks.Open(file_path)  # Open the specified Excel workbook
    wb.Save()  # Save the workbook (equivalent to pressing Ctrl+S)
    wb.Close(False)  # Close the workbook without prompting to save changes
    excel.Quit()  # Quit the Excel application

    print("Excel file saved successfully.")  # Confirmation message

# Run the main function if this script is executed directly
if __name__ == "__main__":
    main()