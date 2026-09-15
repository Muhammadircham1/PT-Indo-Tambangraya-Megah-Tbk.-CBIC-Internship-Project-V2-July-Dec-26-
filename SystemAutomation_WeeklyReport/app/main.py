import sys  # Import the sys module for system-specific parameters and functions
import os  # Import the os module for interacting with the operating system
from PyQt6.QtWidgets import QApplication  # Import QApplication from PyQt6 for GUI application
from gui.main_gui import MyApp  # Import the main application class from the GUI module

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
        # Check if the application is running in a PyInstaller bundle
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS  # Use the temporary folder created by PyInstaller
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))  # Get the directory of the current script
        return os.path.join(base_path, relative_path)  # Join the base path with the relative path

def main():
    app = QApplication(sys.argv)  # Create a QApplication instance

    # Load the style.qss file using ResourceHelper to get the relative path
    try:
        file_path_qss = ResourceHelper.get_path('style/style.qss')  # Get the path to the stylesheet
        with open(file_path_qss, encoding='utf-8') as f:  # Open the stylesheet file
            app.setStyleSheet(f.read())  # Set the application's stylesheet
    except FileNotFoundError:
        print("Style file not found, continuing without stylesheet.")  # Handle missing stylesheet

    gui = MyApp()  # Create an instance of the main application GUI
    gui.show()  # Show the GUI window
    sys.exit(app.exec())  # Start the application event loop and exit when done

# Run the main function if this script is executed directly
if __name__ == "__main__":
    main()