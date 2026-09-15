# main.py
import sys
import io
import json
import os
from pathlib import Path

# Importing various modules for processing different steps
import add_row
import copy_data
import ongoing_month
import month_1
import month_2
import month_3
import month_4
import month_5
import month_6
import save  # Import the save module for saving the final output

def run_step(func, label: str) -> None:
    """
    Run a specified function and print the status.

    Args:
        func: The function to run (should have a main() method).
        label (str): A label for logging purposes to indicate which step is being executed.
    """
    print(f"Running {label}...")  # Print status before starting the function
    func.main()  # Call the main method of the specified function
    print(f"{label} success.", flush=True)  # Indicate that the step was successful

class ResourceHelper:
    @staticmethod
    def get_path(relative_path: str) -> Path:
        """
        Get the absolute path of a file relative to the current script's directory.

        Args:
            relative_path (str): The relative path to the file.

        Returns:
            Path: The absolute path to the file.
        """
        # Use pathlib.Path to get the file path
        base_path = Path(__file__).parent
        return base_path / relative_path
    
if __name__ == "__main__":
    # Ensure stdout is in UTF-8 (in case of non-ASCII characters)
    if sys.stdout.encoding.lower() != "utf-8":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    # Read configuration from the JSON file
    cfg_path = ResourceHelper.get_path('../config/inputan.json')
    json_data = json.loads(cfg_path.read_text(encoding="utf-8"))

    print("Starting execution...\n")  # Indicate the start of the execution process

    # Mandatory steps that must be executed
    run_step(add_row,       "Process add_row")  # Add rows to the Excel file
    run_step(copy_data,     "Process penalty & demurrage")  # Copy penalty and demurrage data
    run_step(ongoing_month, "Process ongoing month")  # Process ongoing month data

    # Conditional steps for processing months 1-4 based on data availability
    months = [
        ("data_count_month1", month_1, "Process month 1"),
        ("data_count_month2", month_2, "Process month 2"),
        ("data_count_month3", month_3, "Process month 3"),
        ("data_count_month4", month_4, "Process month 4"),
        ("data_count_month5", month_5, "Process month 5"),
        ("data_count_month6", month_6, "Process month 6"),
    ]

    # Iterate through each month and run the corresponding module if data is available
    for key, module, label in months:
        if json_data.get(key, 0) > 0:  # Check if there is data to process for the month
            run_step(module, label)  # Run the processing function for the month
        else:
            print(f"{label} skipped, data does not change")  # Indicate that the step was skipped

    # Finally, run the save step to save the changes made to the Excel file
    run_step(save, "Autosave Excel draft")  # This calls the main() function in save.py

    print("\nExecution completed.")  # Indicate that the execution has finished
    print("Automation completed successfully!", flush=True)  # Final success message
    sys.exit(0)  # Exit the program with a success status