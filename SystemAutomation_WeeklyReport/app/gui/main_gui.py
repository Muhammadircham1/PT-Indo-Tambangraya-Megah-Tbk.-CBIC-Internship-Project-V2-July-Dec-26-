import sys
import json
import os
import pathlib

from PyQt6.QtCore import QThread
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QDialog, QTextEdit, QComboBox, QSpinBox, QCheckBox, QPushButton, QLineEdit, QMessageBox
)
from PyQt6 import QtGui, QtCore

from gui.process import ProcessWorker
from gui.ui_window import Ui_MainWindow
from gui.mini_popup import Ui_dialog
from gui.popup import Ui_Dialog

import importlib.util

# ---------- Helper for relative path ---------------------------------
class ResourceHelper:
    @staticmethod
    def get_path(relative_path):
        base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

# ---------- Dialog "Process Completed" -----------------------------------
class ProcessDoneDialog(QDialog, Ui_dialog):
    def __init__(self, parent=None, status="success"):
        super().__init__(parent)
        self.setupUi(self, status)  # Pass status to setupUi
        self.setModal(True)
        self.setFixedSize(self.size())
        self.pushButton_ok.clicked.connect(self.accept)

# ---------- Confirmation Popup Before Process --------------------------
class ConfirmationPopup(QDialog, Ui_Dialog):
    def __init__(self, data_dict, confirm_callback):
        super().__init__()
        self.setupUi(self)
        self.confirm_callback = confirm_callback

        te: QTextEdit = self.findChild(QTextEdit, "textEdit_popup")
        if te:
            te.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
            te.setFont(QtGui.QFont("Consolas", 10))
            te.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
            te.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)

            max_key_len = max(len(k) for k in data_dict)
            lines = []
            for k, v in data_dict.items():
                v_disp = os.path.basename(v) if 'file' in k.lower() else str(v)
                key = k.ljust(max_key_len)
                lines.append(f"{key} : {v_disp}")

            te.setPlainText("\n".join(lines))

        self.btn_send.clicked.connect(self.send_data)
        self.btn_cancel.clicked.connect(self.reject)

    def send_data(self):
        self.confirm_callback()
        self.accept()

# ---------- Main Application ---------------------------------------------
class MyApp(QMainWindow, Ui_MainWindow):

    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(__file__)

    SCRIPT_MAIN = (pathlib.Path(base_path) / '../logic/main_logic.py').resolve().as_posix()

    def __init__(self): 
        super().__init__()
        self.setupUi(self)

        # Set buttons as checkable
        self.pushButton_Home.setCheckable(True)
        self.pushButton_Performance.setCheckable(True)
        self.pushButton_3rdParty.setCheckable(True)

        # Initialize the UI
        self.stackedWidget.setCurrentWidget(self.page_1)
        self.comboBox_week.addItems([f"W{i}" for i in range(6)])
        self.textEdit_log.setReadOnly(True)

        # Connect buttons to their respective methods
        self.btn_summary.clicked.connect(self.browse_summary_file)
        self.btn_final.clicked.connect(self.browse_final_file)
        self.pushButton_submit.clicked.connect(self.collect_and_confirm)
        self.pushButton_Home.clicked.connect(lambda: self.switch_page(self.page_1, self.pushButton_Home))
        self.pushButton_Performance.clicked.connect(lambda: self.switch_page(self.page_2, self.pushButton_Performance))
        self.pushButton_3rdParty.clicked.connect(lambda: self.switch_page(self.page_4, self.pushButton_3rdParty))
        self.pushButton_process.clicked.connect(self.run_main_program)
        self.pushButton_end.clicked.connect(self.end_process)  # Connect "End" button

        # 3rd Party Buttons
        self.pushButton_Raw3rdParty.clicked.connect(self.browse_raw_file)
        self.pushButton_Draft3rdParty.clicked.connect(self.browse_draft_file)
        self.pushButton_Start3rdParty.clicked.connect(self.program_3rdParty)
        self.lineEdit_Raw3rdParty.textChanged.connect(self.update_start_button_state)
        self.lineEdit_Draft3rdParty.textChanged.connect(self.update_start_button_state)

        self.summary_file = ""
        self.final_file = ""
        self.output_data = {}

        # Threading for process
        self.thread: QThread | None = None
        self.worker: ProcessWorker | None = None

        # Connect checkbox to SpinBox for month 4
        self.checkBox_enableMonth4.toggled.connect(self.toggle_month4_spinboxes)
        self.toggle_month4_spinboxes()

        # Connect checkbox to SpinBox for month 4
        self.checkBox_enableMonth5.toggled.connect(self.toggle_month5_spinboxes)
        self.toggle_month5_spinboxes()

        # Connect checkbox to SpinBox for month 4
        self.checkBox_enableMonth6.toggled.connect(self.toggle_month6_spinboxes)
        self.toggle_month6_spinboxes()

        # Ensure the Home button is checked and the corresponding page is active
        self.switch_page(self.page_1, self.pushButton_Home)

    def switch_page(self, page, button):
        # Set the current widget to the specified page
        self.stackedWidget.setCurrentWidget(page)

        # Set only the clicked button checked, disable others
        self.pushButton_Home.setChecked(button == self.pushButton_Home)
        self.pushButton_Performance.setChecked(button == self.pushButton_Performance)
        self.pushButton_3rdParty.setChecked(button == self.pushButton_3rdParty)

        # Extra safety: disable other buttons that aren't clicked
        if button != self.pushButton_Home:
            self.pushButton_Home.setChecked(False)
        if button != self.pushButton_Performance:
            self.pushButton_Performance.setChecked(False)
        if button != self.pushButton_3rdParty:
            self.pushButton_3rdParty.setChecked(False)

    # Toggle enabling/disabling spinboxes for month 4
    def toggle_month4_spinboxes(self):
        is_checked = self.checkBox_enableMonth4.isChecked()
        self.spinBox_headerMonth4.setEnabled(is_checked)
        self.spinBox_dataCountMonth4.setEnabled(is_checked)

    # Toggle enabling/disabling spinboxes for month 5
    def toggle_month5_spinboxes(self):
        is_checked = self.checkBox_enableMonth5.isChecked()
        self.spinBox_headerMonth5.setEnabled(is_checked)
        self.spinBox_dataCountMonth5.setEnabled(is_checked)

    # Toggle enabling/disabling spinboxes for month 6
    def toggle_month6_spinboxes(self):
        is_checked = self.checkBox_enableMonth6.isChecked()
        self.spinBox_headerMonth6.setEnabled(is_checked)
        self.spinBox_dataCountMonth6.setEnabled(is_checked)

    # Browse summary input file
    def browse_summary_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Source File", "", "Excel (*.xlsx)")
        if path:
            self.summary_file = path
            self.lineEdit_summary.setText(path)

    # Browse final output file
    def browse_final_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Output File Location", "", "Excel (*.xlsx)")
        if path:
            self.final_file = path
            self.lineEdit_final.setText(path)

    # Collect user inputs and show confirmation
    def collect_and_confirm(self):
        try:
            self.new_output_data = {
                "summary_file": self.summary_file,
                "final_file": self.final_file,
                "header_month1": int(self.spinBox_headerMonth1.value()),
                "header_month2": int(self.spinBox_headerMonth2.value()),
                "header_month3": int(self.spinBox_headerMonth3.value()),
                "header_month4": int(self.spinBox_headerMonth4.value()) if self.checkBox_enableMonth4.isChecked() else 0,
                "header_month5": int(self.spinBox_headerMonth5.value()) if self.checkBox_enableMonth5.isChecked() else 0,
                "header_month6": int(self.spinBox_headerMonth6.value()) if self.checkBox_enableMonth6.isChecked() else 0,
                "data_count_month1": int(self.spinBox_dataCountMonth1.value()),
                "data_count_month2": int(self.spinBox_dataCountMonth2.value()),
                "data_count_month3": int(self.spinBox_dataCountMonth3.value()),
                "data_count_month4": int(self.spinBox_dataCountMonth4.value()) if self.checkBox_enableMonth4.isChecked() else 0,
                "data_count_month5": int(self.spinBox_dataCountMonth5.value()) if self.checkBox_enableMonth5.isChecked() else 0,
                "data_count_month6": int(self.spinBox_dataCountMonth6.value()) if self.checkBox_enableMonth6.isChecked() else 0,
                "selected_week": self.comboBox_week.currentText()
            }
        except ValueError:
            self.textEdit_log.append("Numeric input is incomplete.")
            return

        ConfirmationPopup(self.new_output_data, self.save_to_json_and_goto_page3).exec()

    # Save inputs to JSON file and change to page 3
    def save_to_json_and_goto_page3(self):
        json_file_path = ResourceHelper.get_path('../config/inputan.json')
        try:
            with open(json_file_path, "r") as fp:
                existing_data = json.load(fp)
        except FileNotFoundError:
            existing_data = {}

        existing_data.update(self.new_output_data)

        with open(json_file_path, "w") as fp:
            json.dump(existing_data, fp, indent=4)

        self.stackedWidget.setCurrentWidget(self.page_3)

    # Run main logic script in a separate thread
    def run_main_program(self):
        if self.thread and self.thread.isRunning():
            self.textEdit_log.append("Process is still running.")
            return

        self.textEdit_log.clear()
        self.pushButton_process.setEnabled(False)
        self.pushButton_end.setEnabled(False)

        self.thread = QThread(self)
        self.worker = ProcessWorker(self.SCRIPT_MAIN)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log.connect(self.append_log)
        self.worker.error.connect(self.append_error)
        self.worker.finished.connect(self.on_finished)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    # Append normal log message
    def append_log(self, txt: str):
        self.textEdit_log.append(txt)

    # Append error log message
    def append_error(self, txt: str):
        self.textEdit_log.append(f"<span style='color:red'>{txt}</span>")

    # Called when background process finished
    def on_finished(self, exit_code: int):
        status_text = "Normal" if exit_code == 0 else "Error"
        self.textEdit_log.append(f"\nFinished (exit code {exit_code}, status {status_text}).")
        
        # Show the popup with appropriate status
        status = "success" if exit_code == 0 else "error"
        ProcessDoneDialog(self, status).exec()
        
        self.pushButton_process.setEnabled(True)
        self.pushButton_end.setEnabled(True)
        self.thread.quit()

    # End button handler: clear SSO form and go home
    def end_process(self):
        self.clear_sso_form()
        self.switch_page(self.page_1, self.pushButton_Home)

    def clear_sso_form(self):
        self.lineEdit_summary.clear()
        self.lineEdit_final.clear()
        self.spinBox_headerMonth1.setValue(0)
        self.spinBox_headerMonth2.setValue(0)
        self.spinBox_headerMonth3.setValue(0)
        self.spinBox_headerMonth4.setValue(0)
        self.spinBox_headerMonth5.setValue(0)
        self.spinBox_headerMonth6.setValue(0)
        self.spinBox_dataCountMonth1.setValue(0)
        self.spinBox_dataCountMonth2.setValue(0)
        self.spinBox_dataCountMonth3.setValue(0)
        self.spinBox_dataCountMonth4.setValue(0)
        self.spinBox_dataCountMonth5.setValue(0)
        self.spinBox_dataCountMonth6.setValue(0)
        self.checkBox_enableMonth4.setChecked(False)
        self.checkBox_enableMonth5.setChecked(False)
        self.checkBox_enableMonth6.setChecked(False)
        self.comboBox_week.setCurrentIndex(0)
        self.textEdit_log.clear()

    # ================ Third Party Integration ============
    def program_3rdParty(self):
        try:
            # Determine base path depending on frozen state
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            # Get file paths from user input
            raw_file_path = self.lineEdit_Raw3rdParty.text()
            draft_file_path = self.lineEdit_Draft3rdParty.text()
            third_party_logic_path = os.path.join(base_path, '../logic/3rd_party.py')

            # Validate both files are selected
            if not raw_file_path or not draft_file_path:
                print("Please select both files first.")
                return

            # Dynamically load third_party module
            spec = importlib.util.spec_from_file_location("third_party_logic", third_party_logic_path)
            third_party_logic = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(third_party_logic)

            # Call move_data function from third_party module
            third_party_logic.move_data(raw_file_path, draft_file_path, sheet_a='YTD', sheet_b='3rd Party')

            # Show success message
            ProcessDoneDialog(self).exec()

            # Clear inputs after process completion
            self.lineEdit_Raw3rdParty.clear()
            self.lineEdit_Draft3rdParty.clear()

        except Exception as e:
            print(f"Error occurred: {str(e)}")
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    # Browse 3rd party raw file
    def browse_raw_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Source File", "", "Excel (*.xlsx)")
        if path:
            self.raw_file = path
            self.lineEdit_Raw3rdParty.setText(path)

    # Browse 3rd party draft output file
    def browse_draft_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Output File Location", "", "Excel (*.xlsx)")
        if path:
            self.draft_file = path
            self.lineEdit_Draft3rdParty.setText(path)

    # Enable start button only if both lineEdits have text
    def update_start_button_state(self):
        enabled = bool(self.lineEdit_Raw3rdParty.text()) and bool(self.lineEdit_Draft3rdParty.text())
        self.pushButton_Start3rdParty.setEnabled(enabled)

# ---------- Main Entry Point ---------------------------------------------
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