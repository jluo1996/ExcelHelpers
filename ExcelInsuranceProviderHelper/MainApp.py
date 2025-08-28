import os
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QLineEdit, QLabel, QComboBox, QDesktopWidget, QTextEdit
import tkinter as tk
from tkinter import filedialog
from InsuranceStatusHelper import InsuranceStatusHelper
from InsuranceStatusHelperEnum import INSURANCE_FORMAT_ENUM, PLAN_TYPE_ENUM
from logger import Logger

WINDOW_WIDTH = 600
WINDOW_HEIGHT = 200

class MyWindow(QWidget):
    def __init__(self, logger : Logger = None):
        super().__init__()

        self.logger = logger
        self._init_UI()

        self.logger.text_edit = self.log_textedit


    def _init_UI(self):
        self.setWindowTitle("Insurance Provider Excel Helper")
        self.resize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.center()

        # ADP file
        self.adp_file_path_textedit = QLineEdit()
        self.adp_file_path_textedit.setReadOnly(True)
        self.adp_file_path_textedit.setPlaceholderText("Select ADP file")
        adp_file_browse_button = QPushButton("Browse")
        adp_file_browse_button.clicked.connect(self.adp_file_browse_button_clicked)

        input_file_h_layout = QHBoxLayout()
        input_file_h_layout.addWidget(self.adp_file_path_textedit)
        input_file_h_layout.addWidget(adp_file_browse_button)
        # --------- end of Input file


        # Insurance file 
        self.insurance_file_path_textedit = QLineEdit()
        self.insurance_file_path_textedit.setReadOnly(True)
        self.insurance_file_path_textedit.setPlaceholderText("Select insurance file")
        insurance_file_browse_button = QPushButton("Browse")
        insurance_file_browse_button.clicked.connect(self.insurance_file_browse_button_clicked)

        insurance_file_h_layout = QHBoxLayout()
        insurance_file_h_layout.addWidget(self.insurance_file_path_textedit)
        insurance_file_h_layout.addWidget(insurance_file_browse_button)
        # ---------- end of insurance file


        # Output folder path
        self.output_folder_path_textedit = QLineEdit()
        self.output_folder_path_textedit.setReadOnly(True)
        self.output_folder_path_textedit.setPlaceholderText("Select output folder")
        output_folder_path_button = QPushButton("Browse")
        output_folder_path_button.clicked.connect(self.output_folder_path_button_clicked)

        output_folder_path_h_layout = QHBoxLayout()
        output_folder_path_h_layout.addWidget(self.output_folder_path_textedit)
        output_folder_path_h_layout.addWidget(output_folder_path_button)
        # ---------- end of Output folder path


        # Insurance Provider
        insurance_provider_label = QLabel()
        insurance_provider_label.setText("Select Insurance Provider:")
        self.insurance_provider_combobox = QComboBox()
        for provider in INSURANCE_FORMAT_ENUM:
            self.insurance_provider_combobox.addItem(provider.name, provider.value)

        insurance_provider_h_layout = QHBoxLayout()
        insurance_provider_h_layout.addWidget(insurance_provider_label)
        insurance_provider_h_layout.addWidget(self.insurance_provider_combobox)
        # ---------- end of insurance provider


        # Insurance Plan Type
        insurance_plan_type_label = QLabel()
        insurance_plan_type_label.setText("Select Insurance Type:")
        self.insurance_plan_type_combobox = QComboBox()
        for plan_type in PLAN_TYPE_ENUM:
            self.insurance_plan_type_combobox.addItem(plan_type.name, plan_type.value)

        insurance_plan_type_h_layout = QHBoxLayout()
        insurance_plan_type_h_layout.addWidget(insurance_plan_type_label)
        insurance_plan_type_h_layout.addWidget(self.insurance_plan_type_combobox)
        # ---------- end of Insurance Plan Type


        # Log area
        self.log_textedit = QTextEdit()
        self.log_textedit.setReadOnly(True)




        # Function buttons
        generate_status_report_button = QPushButton("Generate Status Report")
        generate_status_report_button.clicked.connect(self.generate_status_report_button_clicked)

        function_button_h_box = QHBoxLayout() # Keep this layout for more functions in the future
        function_button_h_box.addWidget(generate_status_report_button)
        # ---------- end of Function buttons

        main_v_layout = QVBoxLayout()
        main_v_layout.addLayout(input_file_h_layout)
        main_v_layout.addLayout(insurance_file_h_layout)
        main_v_layout.addLayout(output_folder_path_h_layout)
        main_v_layout.addLayout(insurance_provider_h_layout)
        main_v_layout.addLayout(insurance_plan_type_h_layout)
        main_v_layout.addWidget(self.log_textedit)
        main_v_layout.addLayout(function_button_h_box)
        self.setLayout(main_v_layout)

    def center(self):
        """
        Used to center the window
        """
        # Get window rectangle
        qr = self.frameGeometry()

        # Get screen center point
        cp = QDesktopWidget().availableGeometry().center()

        # Move rectangle center to screen center
        qr.moveCenter(cp)

        # Move top-left of window to rectangle's top-left
        self.move(qr.topLeft())

    def output_folder_path_button_clicked(self):
        output_folder_path = self.get_folder_from_user()
        self.output_folder_path_textedit.setText(output_folder_path)

    def adp_file_browse_button_clicked(self):
        adp_file_full_path = self.get_excel_file_from_user()
        self.adp_file_path_textedit.setText(adp_file_full_path)

    def generate_status_report_button_clicked(self):
        proceed, error_msgs = self.get_is_ready_to_generate_status_report()
        
        if not proceed:
            self.logger.log_error("Cannot proceed. See reasons below.")
            for msg in error_msgs:
                self.logger.log_error(f"- {msg}")
            return
        
        adp_file_path = self.get_adp_file_full_path()
        insurance_file_path = self.get_insurance_file_path()
        insurance_provider_type = INSURANCE_FORMAT_ENUM(self.get_selected_insurance_provider_index())
        plan_type = PLAN_TYPE_ENUM(self.get_selected_insurance_plan_type_index())
        output_folder = self.get_output_folder_path()
        helper = InsuranceStatusHelper(adp_file_path, insurance_file_path, insurance_provider_type, plan_type, output_folder, self.logger)
        helper.generate_status_report(False)

    def insurance_file_browse_button_clicked(self):
        insurance_file_full_path = self.get_excel_file_from_user()
        self.insurance_file_path_textedit.setText(insurance_file_full_path)

    def get_is_ready_to_generate_status_report(self) -> tuple[bool, list[str]]:
        error_msg = []

        # Validate ADP file
        adp_file_full_path = self.get_adp_file_full_path()
        if not adp_file_full_path:
            error_msg.append("No ADP file selected.")
        elif not os.path.isabs(adp_file_full_path):
            error_msg.append(f"{adp_file_full_path} is not an absolute path.")
        elif not adp_file_full_path.lower().endswith(".xlsx"):
            error_msg.append(f"{adp_file_full_path} is not a xlsx file.")
        # ------- end of Validate ADP file

        # Validate Insurance file
        insurance_file_full_path = self.get_insurance_file_path()
        if not insurance_file_full_path:
            error_msg.append("No insurance file selected.")
        elif not os.path.isabs(insurance_file_full_path):
            error_msg.append(f"{insurance_file_full_path} is not an absolute path.")
        elif not insurance_file_full_path.lower().endswith(".xlsx"):
            error_msg.append(f"{insurance_file_full_path} is not a xlsx file.")
        # ------- end of Validate Insurance file

        # Validate Output folder
        output_folder_path = self.get_output_folder_path()
        if not output_folder_path:
            error_msg.append("No output folder selected.")
        elif not os.path.isabs(output_folder_path):
            error_msg.append(f"{output_folder_path} is not an absolute path.")
        elif not os.path.isdir(output_folder_path):
            error_msg.append(f"{output_folder_path} is not a valid directory.")
        # ------- end of Validate Output folder

        is_ready = len(error_msg) == 0

        return is_ready, error_msg

    def is_valid_folder_path(self, path : str) -> bool:
        # Empty string is not a valid folder path
        if not path:
            return False

        # Check if it's an absolute path and points to a directory
        return os.path.isabs(path) and os.path.isdir(path)

    def is_valid_xlsx_file_full_path(self, path : str) -> bool:
        # Empty string is not valid
        if not path:
            return False
        
        # Must be absolute path and end with .xlsx
        return os.path.isabs(path) and path.lower().endswith(".xlsx")
    
    def get_adp_file_full_path(self):
        return self.adp_file_path_textedit.text()
    
    def get_insurance_file_path(self):
        return self.insurance_file_path_textedit.text()
    
    def get_output_folder_path(self):
        return self.output_folder_path_textedit.text()

    def get_selected_insurance_provider_index(self):
        return self.insurance_provider_combobox.currentIndex()

    def get_selected_insurance_plan_type_index(self):
        return self.insurance_plan_type_combobox.currentIndex()

    def get_excel_file_from_user(self):
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        return file_path
    
    def get_folder_from_user(self):
        root = tk.Tk()
        root.withdraw()
        folder_path = filedialog.askdirectory(title="Select a folder")
        return folder_path


if __name__ == "__main__":
    app = QApplication(sys.argv)
    logger = Logger()
    window = MyWindow(logger)
    window.show()
    sys.exit(app.exec_())
