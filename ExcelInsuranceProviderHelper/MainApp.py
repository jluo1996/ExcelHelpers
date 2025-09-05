import os
import sys
from PyQt5.QtCore import QDate, QSharedMemory, QSystemSemaphore, QUrl
from PyQt5.QtWidgets import QApplication, QDateEdit, QTextBrowser, QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QLineEdit, QLabel, QComboBox, QDesktopWidget
import tkinter as tk
from tkinter import filedialog
from InsuranceStatusHelper import InsuranceStatusHelper
from InsuranceStatusHelperEnum import INSURANCE_FORMAT_ENUM
from logger import Logger

WINDOW_WIDTH = 600
WINDOW_HEIGHT = 200

class MyWindow(QWidget):
    def __init__(self, logger : Logger = None):
        super().__init__()

        self.logger = logger
        self._init_UI()

        self.logger.text_browser = self.log_textbrowser


    def _init_UI(self):
        self.setWindowTitle("Insurance Provider Excel Helper")
        self.resize(WINDOW_WIDTH, WINDOW_HEIGHT)
        self.center()

        # ADP file
        self.adp_file_path_textedit = QLineEdit()
        self.adp_file_path_textedit.setReadOnly(True)
        self.adp_file_path_textedit.setPlaceholderText("Select ADP file")
        adp_file_browse_button = QPushButton("Select ADP file")

        input_file_h_layout = QHBoxLayout()
        input_file_h_layout.addWidget(self.adp_file_path_textedit)
        input_file_h_layout.addWidget(adp_file_browse_button)
        # --------- end of Input file


        # Insurance file 
        self.insurance_file_path_textedit = QLineEdit()
        self.insurance_file_path_textedit.setReadOnly(True)
        self.insurance_file_path_textedit.setPlaceholderText("Select Insurance File")
        insurance_file_browse_button = QPushButton("Select Insurance File")

        insurance_file_h_layout = QHBoxLayout()
        insurance_file_h_layout.addWidget(self.insurance_file_path_textedit)
        insurance_file_h_layout.addWidget(insurance_file_browse_button)
        # ---------- end of insurance file


        # Output folder path
        self.output_folder_path_textedit = QLineEdit()
        self.output_folder_path_textedit.setReadOnly(True)
        self.output_folder_path_textedit.setPlaceholderText("Select Output Folder")
        output_folder_path_button = QPushButton("Select Output Folder")

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


        # Insurance ID file 
        self.insurance_id_ile_path_textedit = QLineEdit()
        self.insurance_id_ile_path_textedit.setReadOnly(True)
        self.insurance_id_ile_path_textedit.setPlaceholderText("Select Insurance ID file")
        insurance_id_file_browse_button = QPushButton("Select ID File")

        self.insurance_id_file_container = QWidget()
        insurance_id_file_h_layout = QHBoxLayout(self.insurance_id_file_container)
        insurance_id_file_h_layout.addWidget(self.insurance_id_ile_path_textedit)
        insurance_id_file_h_layout.addWidget(insurance_id_file_browse_button)
        insurance_id_file_h_layout.setContentsMargins(0, 0, 0, 0)  # left, top, right, bottom
        insurance_id_file_h_layout.setSpacing(0)  # optional, removes spacing between child widgets
        # ---------- end of insurance file  


        # Date selection
        self.start_date_label = QLabel()
        self.start_date_label.setText("Start Date")
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        
        self.end_date_label = QLabel()
        self.end_date_label.setText("End Date")
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        
        self.data_selection_container = QWidget()
        date_selection_h_layout = QHBoxLayout(self.data_selection_container)
        date_selection_h_layout.addWidget(self.start_date_label)
        date_selection_h_layout.addWidget(self.start_date_edit)
        date_selection_h_layout.addWidget(self.end_date_label)
        date_selection_h_layout.addWidget(self.end_date_edit)


        # Log area
        self.log_textbrowser = QTextBrowser()
        self.log_textbrowser.setReadOnly(True)
        self.log_textbrowser.setOpenLinks(False)
        self.log_textbrowser.setOpenExternalLinks(False)
        self.log_textbrowser.anchorClicked.connect(self.output_file_path_text_clicked)
        self.log_textbrowser.setHtml("")


        # Function buttons
        generate_status_report_button = QPushButton("Generate Status Report")

        function_button_h_box = QHBoxLayout() # Keep this layout for more functions in the future
        function_button_h_box.addWidget(generate_status_report_button)
        # ---------- end of Function buttons


        # Command connect
        adp_file_browse_button.clicked.connect(self.adp_file_browse_button_clicked)
        insurance_file_browse_button.clicked.connect(self.insurance_file_browse_button_clicked)
        output_folder_path_button.clicked.connect(self.output_folder_path_button_clicked)
        self.insurance_provider_combobox.currentIndexChanged.connect(self.insurance_provider_selection_changed)
        insurance_id_file_browse_button.clicked.connect(self.insurance_id_file_browse_button_clicked)
        generate_status_report_button.clicked.connect(self.generate_status_report_button_clicked)
        # --------- end of Command connect

        main_v_layout = QVBoxLayout()
        main_v_layout.addLayout(input_file_h_layout)
        main_v_layout.addLayout(insurance_file_h_layout)
        main_v_layout.addLayout(output_folder_path_h_layout)
        main_v_layout.addLayout(insurance_provider_h_layout)
        main_v_layout.addWidget(self.insurance_id_file_container)
        main_v_layout.addWidget(self.data_selection_container)
        main_v_layout.addWidget(self.log_textbrowser)
        main_v_layout.addLayout(function_button_h_box)
        self.setLayout(main_v_layout)


        self.insurance_provider_selection_changed(self.insurance_provider_combobox.currentIndex()) # manually trigger index change

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

    def output_file_path_text_clicked(self, url: QUrl):
        output_path = url.toLocalFile()
        output_path = os.path.dirname(output_path)
        if os.path.exists(output_path):
            os.startfile(output_path)
        else:
            self.logger.log_error(f"Failed to open {output_path}")

    def output_folder_path_button_clicked(self):
        output_folder_path = self.get_folder_from_user()
        self.output_folder_path_textedit.setText(output_folder_path)
    
    def insurance_provider_selection_changed(self, index):
        self.insurance_id_file_container.setVisible(index == INSURANCE_FORMAT_ENUM.CIGNA.value)
        self.data_selection_container.setVisible(index == INSURANCE_FORMAT_ENUM.CIGNA.value)

    def insurance_id_file_browse_button_clicked(self):
        insurance_id_file_path = self.get_excel_file_from_user()
        if insurance_id_file_path:
            self.insurance_id_ile_path_textedit.setText(insurance_id_file_path)

    def adp_file_browse_button_clicked(self):
        adp_file_full_path = self.get_excel_file_from_user()
        if adp_file_full_path:
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
        id_file_path = self.get_insurance_id_file_path()
        start_date = self.get_selected_start_date()
        end_date = self.get_selected_end_date()
        insurance_provider_type = INSURANCE_FORMAT_ENUM(self.get_selected_insurance_provider_index())
        output_folder = self.get_output_folder_path()
        self.helper = InsuranceStatusHelper(adp_file_full_path=adp_file_path,
                                            insurance_file_full_path=insurance_file_path, 
                                            id_file_full_path=id_file_path, 
                                            start_date=start_date, 
                                            end_date=end_date,
                                            insurance_provider_type=insurance_provider_type, 
                                            output_folder=output_folder, 
                                            logger=self.logger)
        
        run_as_thread = not __debug__
        if run_as_thread:
            self.helper.set_finish_method(self.job_completed)
            self.enable_all_interactive_UI(False)
        output_file_path = self.helper.generate_status_report(run_as_thread)
        if not run_as_thread and output_file_path:
            self.job_completed(output_file_path)

    def enable_all_interactive_UI(self, enable : bool = True):
        self.enable_all_buttons(enable)
        self.enable_all_comboboxes(enable)

    def enable_all_buttons(self, enable : bool = True):
        for btn in self.findChildren(QPushButton):
            btn.setEnabled(enable)

    def enable_all_comboboxes(self, enable: bool = True):
        for combobox in self.findChildren(QComboBox):
            combobox.setEnabled(enable)

    def job_completed(self, path :str ):
        self.enable_all_interactive_UI()
        if path:
            self.output_file_path_text_clicked(QUrl.fromLocalFile(path))

    def insurance_file_browse_button_clicked(self):
        insurance_file_full_path = self.get_excel_file_from_user()
        if insurance_file_full_path:
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

        # Validate ID file 
        if self.get_selected_insurance_provider_index() == INSURANCE_FORMAT_ENUM.CIGNA.value:
            id_file_path = self.get_insurance_id_file_path()
            if not id_file_path:
                error_msg.append(f"ID file is reqiured for {((INSURANCE_FORMAT_ENUM)(self.get_selected_insurance_provider_index())).get_string()}.")
            elif not os.path.isabs(id_file_path):
                error_msg.append(f"{id_file_path} is not an absolute path.")
            elif not id_file_path.lower().endswith(".xlsx"):
                error_msg.append(f"{id_file_path} is not a xlsx file.")
        # ------- end of Validate ID file 

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
    
    def get_selected_start_date(self) -> str:
        return self.start_date_edit.date().toString("MM/dd/yyyy")
    
    def get_selected_end_date(self) -> str:
        return self.end_date_edit.date().toString("MM/dd/yyyy")
    
    def get_adp_file_full_path(self):
        return self.adp_file_path_textedit.text()
    
    def get_insurance_id_file_path(self):
        return self.insurance_id_ile_path_textedit.text()
    
    def get_insurance_file_path(self):
        return self.insurance_file_path_textedit.text()
    
    def get_output_folder_path(self):
        return self.output_folder_path_textedit.text()

    def get_selected_insurance_provider_index(self):
        return self.insurance_provider_combobox.currentIndex()

    def get_excel_file_from_user(self):
        root = tk.Tk()
        root.geometry("+{}+{}".format(int(root.winfo_screenwidth() / 2 - root.winfo_reqwidth() / 2) - 400, int(root.winfo_screenheight() / 2 - root.winfo_reqheight() / 2) - 300))
        root.withdraw()
        # TODO: to use width/height of the pyqt window

        self.setEnabled(False) # disable main window
        file_path = filedialog.askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")], parent=root)
        self.setEnabled(True)
        return file_path
    
    def get_folder_from_user(self):
        root = tk.Tk()
        root.withdraw()
        folder_path = filedialog.askdirectory(title="Select a folder")
        return folder_path

APP_ID = "ExcelInsuranceProviderHelper"
class SingleInstance:
    def __init__(self, key):
        self.key = key
        self.semaphore = QSystemSemaphore(self.key + "_sem", 1)
        self.semaphore.acquire()

        self.shared_memory = QSharedMemory(self.key)
        if self.shared_memory.attach():
            # Another instance already running
            self.is_running = True
        else:
            self.shared_memory.create(1)  # create memory block
            self.is_running = False

        self.semaphore.release()

if __name__ == "__main__":
    instance = SingleInstance(APP_ID)
    if instance.is_running:
        import ctypes  # An included library with Python install.   
        ctypes.windll.user32.MessageBoxW(0, "Another instance is already running.", "Information", 0)
        sys.exit(0)

    app = QApplication(sys.argv)
    logger = Logger()
    window = MyWindow(logger)
    window.show()
    sys.exit(app.exec_())
