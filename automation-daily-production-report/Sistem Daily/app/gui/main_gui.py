import sys
import json
import os
import logging
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QPushButton,
    QLineEdit, QMessageBox, QProgressDialog
)
from PyQt6 import QtGui, QtCore

from gui.ui_interface import Ui_MainWindow
from logic.main_logic import run_main_logic  # Make sure this function uses logging


# ---------- Helper for resolving relative paths ----------
class ResourceHelper:
    @staticmethod
    def get_path(relative_path):
        base_path = (
            os.path.dirname(sys.executable)
            if getattr(sys, 'frozen', False)
            else os.path.dirname(os.path.abspath(__file__))
        )
        return os.path.join(base_path, relative_path)


# ---------- Worker Thread to handle background processing ----------
class MainLogicWorker(QtCore.QThread):
    finished = QtCore.pyqtSignal()
    error = QtCore.pyqtSignal(str)

    def __init__(self, input_paths):
        super().__init__()
        self.input_paths = input_paths

    def run(self):
        logging.info("🛠 Starting main logic process...")
        try:
            run_main_logic(self.input_paths)
            logging.info("🟢 Main logic process completed.")
            self.finished.emit()
        except Exception as e:
            logging.exception("🚨 An error occurred during processing:")
            self.error.emit(str(e))



# ---------- Main Application ----------
class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # Setup logging format
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')


        self.progress_dialog = None  # Placeholder for progress popup

        # Button state setup for navigation
        self.pushButton_Home.setCheckable(True)
        self.pushButton_Automation.setCheckable(True)
        self.switch_page(self.page_1, self.pushButton_Home)

        # Isi ComboBox bulan (January - December)
        months = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ]
        self.comboBoxThisMonth.addItems(months)
        self.comboBoxOutlookMonth.addItems(months)

        # Mapping buttons to corresponding line edits
        self.browseButtons = {
            'IMM': (self.findChild(QPushButton, 'btn_IMM'), self.findChild(QLineEdit, 'lineEdit_IMM')),
            'TCM': (self.findChild(QPushButton, 'btn_TCM'), self.findChild(QLineEdit, 'lineEdit_TCM')),
            'BEK': (self.findChild(QPushButton, 'btn_BEK'), self.findChild(QLineEdit, 'lineEdit_BEK')),
            'GPK': (self.findChild(QPushButton, 'btn_GPK'), self.findChild(QLineEdit, 'lineEdit_GPK')),
            'JBG': (self.findChild(QPushButton, 'btn_JBG'), self.findChild(QLineEdit, 'lineEdit_JBG')),
            'TIS': (self.findChild(QPushButton, 'btn_TIS'), self.findChild(QLineEdit, 'lineEdit_TIS')),
            'SUM': (self.findChild(QPushButton, 'btn_SUM'), self.findChild(QLineEdit, 'lineEdit_SUM')),
            'Output': (self.findChild(QPushButton, 'btn_output'), self.findChild(QLineEdit, 'lineEdit_output')),
            'Summary': (self.findChild(QPushButton, 'btn_summary'), self.findChild(QLineEdit, 'lineEdit_summary')),
        }

        # === Placeholder di dalam LineEdit (bukan di tombol) ===
        self.lineEdit_output.setPlaceholderText("*Draft for Power BI Dashboard source")
        self.lineEdit_summary.setPlaceholderText("*File Summary for Pro Outlook")

        # Sembunyikan label keterangan yang redundant
        if hasattr(self, "label_20"):
            self.label_20.hide()
        if hasattr(self, "label_27"):
            self.label_27.hide()

        # Connect browse buttons ke file dialog
        for _, (button, lineedit) in self.browseButtons.items():
            if button:
                button.clicked.connect(lambda _, le=lineedit: self.browse_file(le))

        # Connect main buttons
        self.pushButton_submit.clicked.connect(self.save_to_json_and_goto_page3)
        self.pushButton_Home.clicked.connect(lambda: self.switch_page(self.page_1, self.pushButton_Home))
        self.pushButton_Automation.clicked.connect(lambda: self.switch_page(self.page_2, self.pushButton_Automation))
        self.pushButton_process.clicked.connect(self.run_process)

    def switch_page(self, page, button):
        """Switches to the selected page and updates button states."""
        self.stackedWidget.setCurrentWidget(page)
        self.pushButton_Home.setChecked(button == self.pushButton_Home)
        self.pushButton_Automation.setChecked(button == self.pushButton_Automation)

    def browse_file(self, lineedit):
        """Open a file dialog and set the selected path to the line edit."""
        path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*.*)")
        if path:
            lineedit.setText(path)

    def save_to_json_and_goto_page3(self):
        """Saves all file paths to the JSON config and switches to page 3."""
        try:
            output_data = {"file": {}}
            for key, (_, lineedit) in self.browseButtons.items():
                path = lineedit.text()
                if path:
                    output_data["file"][key] = path

             # === Ambil nilai This Day ===
            selected_day = self.spinBox_thisDay.value()
            output_data["this_Day"] = selected_day

            # Tambahkan nilai dari ComboBox ThisMonth
            selected_month = self.comboBoxThisMonth.currentText()
            output_data["This_Month"] = selected_month

            # Gabungan untuk pro_outlook_sheet_name
            outlook_name = "Outlook"
            outlook_day = f"{self.spinBoxOutlookDate.value()}"
            outlook_month = self.comboBoxOutlookMonth.currentText()
            outlook_year = str(self.spinBoxOutlookYear.value())

            pro_outlook_sheet_name = f"{outlook_name} {outlook_day} {outlook_month} {outlook_year}"
            output_data["Pro_Outlook_Sheet_Name"] = pro_outlook_sheet_name

            json_file_path = ResourceHelper.get_path('../config/inputan.json')

            try:
                with open(json_file_path, "r") as fp:
                    existing_data = json.load(fp)
            except FileNotFoundError:
                existing_data = {}

            existing_data.update(output_data)

            with open(json_file_path, "w") as fp:
                json.dump(existing_data, fp, indent=4)

            logging.info("✅ Input data saved to JSON and switching to page 3.")
            self.stackedWidget.setCurrentWidget(self.page_3)
        except Exception:
            logging.exception("❌ Failed to save input data:")

    def run_process(self):
        """Starts the background processing with a progress dialog."""
        try:
            input_paths = {
                key: le.text()
                for key, (_, le) in self.browseButtons.items()
                if le.text()
            }
            self.pushButton_process.setEnabled(False)

            # Show progress popup
            self.progress_dialog = QProgressDialog("Processing data...", None, 0, 0, self)
            self.progress_dialog.setWindowTitle("Please Wait")
            self.progress_dialog.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
            self.progress_dialog.setCancelButton(None)
            self.progress_dialog.setMinimumDuration(0)
            self.progress_dialog.show()

            self.worker = MainLogicWorker(input_paths)
            self.worker.finished.connect(self.on_process_finished)
            self.worker.error.connect(self.on_process_error)
            self.worker.start()
        except Exception:
            logging.exception("❌ Failed to start processing.")
            self.pushButton_process.setEnabled(True)
            if self.progress_dialog:
                self.progress_dialog.close()

    def on_process_finished(self):
        """Callback when processing completes successfully."""
        if self.progress_dialog:
            self.progress_dialog.close()

        logging.info("✅ Processing complete.")
        self.pushButton_process.setEnabled(True)
        QMessageBox.information(self, "Done ✅", "🎉 Processing completed successfully!\nPlease check the output file.")

    def on_process_error(self, error_message):
        """Callback when an error occurs during processing."""
        if self.progress_dialog:
            self.progress_dialog.close()

        logging.error(f"❌ Error occurred: {error_message}")
        self.pushButton_process.setEnabled(True)
        QMessageBox.critical(self, "Error ❌", f"An error occurred during processing:\n{error_message}")


# ---------- App Entry Point ----------
if __name__ == "__main__":
    app = QApplication(sys.argv)

    try:
        style_file_path = ResourceHelper.get_path('style/style.qss')
        with open(style_file_path, encoding='utf-8') as f:
            app.setStyleSheet(f.read())
    except FileNotFoundError:
        print("Style file not found, continuing without stylesheet.")

    gui = MyApp()
    gui.show()
    sys.exit(app.exec())
