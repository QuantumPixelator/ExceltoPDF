import sys
import os
import threading
import logging
import pythoncom
import win32com.client as win32
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, 
    QFileDialog, QTextEdit, QLabel, QProgressBar
)
from PySide6.QtCore import Qt, Signal, QObject, QMutex
import time

class ConversionWorker(QObject):
    progress_update = Signal(str)
    conversion_complete = Signal(str)
    progress = Signal(int)

    def __init__(self, files, output_dir, retries=3):
        super().__init__()
        self.files = files
        self.output_dir = output_dir
        self.retries = retries
        self.cancelled = False
        self.mutex = QMutex()
        self.saved_pdf_paths = []  # Track saved PDF paths to use in finalize_filenames

    def reset_permissions(self, file):
        try:
            os.chmod(file, 0o600)  # Reset to read-write permissions for the owner
        except Exception as e:
            logging.error(f"Failed to reset permissions for {file} - {e}")

    def convert_files(self):
        pythoncom.CoInitialize()
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = False
        total_files = len(self.files)
        processed_files = 0

        for file in self.files:
            if self.is_cancelled():
                self.progress_update.emit("Conversion cancelled by user.")
                break
            
            file_name = os.path.basename(file)
            # Prepare the PDF path with the proper name (stripping .xls/.xlsx if present)
            if file_name.endswith((".xls", ".xlsx")):
                pdf_name = file_name.rsplit('.', 1)[0] + ".pdf"
            else:
                pdf_name = file_name + ".pdf"

            pdf_path = os.path.join(self.output_dir, pdf_name)
            pdf_path = os.path.normpath(pdf_path)  # Normalize path for consistent handling
            success = False

            self.reset_permissions(file)
            for attempt in range(1, self.retries + 1):
                if self.is_cancelled():
                    break

                try:
                    self.progress_update.emit(f"Attempting conversion ({attempt}/{self.retries}) for: {file_name}")
                    
                    workbook = excel_app.Workbooks.Open(file)
                    workbook.ExportAsFixedFormat(0, pdf_path)
                    workbook.Close(False)
                    
                    self.progress_update.emit(f"Successfully converted: {file_name} to PDF at {pdf_path}")
                    self.saved_pdf_paths.append(pdf_path)  # Track the exact saved PDF path
                    success = True
                    break
                except Exception as e:
                    logging.error(f"Attempt {attempt}: Failed to convert {file_name} - {e}")
                    self.progress_update.emit(f"Attempt {attempt}: Error converting {file_name} - {e}")
                    time.sleep(1)
                finally:
                    workbook = None

            if not success and not self.is_cancelled():
                self.progress_update.emit(f"Failed to convert {file_name} after {self.retries} attempts.")
            
            if not self.is_cancelled():
                processed_files += 1
                progress_percent = int((processed_files / total_files) * 100)
                self.progress.emit(progress_percent)

        excel_app.Quit()
        pythoncom.CoUninitialize()
        self.finalize_filenames()
        self.conversion_complete.emit("Conversion cancelled." if self.is_cancelled() else "All files processed.")

    def finalize_filenames(self):
        """
        After all files have been converted, rename each file by replacing '%20' with spaces
        and removing any '.xls' or '.xlsx' extension from the PDF file name.
        """
        for pdf_path in self.saved_pdf_paths:
            original_pdf_name = os.path.basename(pdf_path)
            
            # Replace '%20' with spaces
            new_pdf_name = original_pdf_name.replace('%20', ' ')

            new_pdf_path = os.path.join(self.output_dir, new_pdf_name)

            if pdf_path != new_pdf_path:
                try:
                    if os.path.exists(pdf_path):
                        os.rename(pdf_path, new_pdf_path)
                        self.progress_update.emit(f"Renamed {original_pdf_name} to {new_pdf_name}")
                    else:
                        self.progress_update.emit(f"File not found for renaming: {original_pdf_name}")
                except Exception as rename_error:
                    self.progress_update.emit(f"Failed to rename {original_pdf_name}: {rename_error}")

    def cancel_conversion(self):
        self.mutex.lock()
        self.cancelled = True
        self.mutex.unlock()

    def is_cancelled(self):
        self.mutex.lock()
        cancelled = self.cancelled
        self.mutex.unlock()
        return cancelled

class ExcelToPDFConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.selected_files = []
        self.output_dir = ""
        self.conversion_thread = None

    def init_ui(self):
        self.setWindowTitle("Excel to PDF Converter")
        self.setGeometry(300, 300, 600, 400)
        
        self.layout = QVBoxLayout()
        
        self.info_label = QLabel("Select Excel files to convert to PDF", self)
        self.info_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.info_label)
        
        self.file_selector_btn = QPushButton("Browse Files", self)
        self.file_selector_btn.setStyleSheet("padding: 10px; font-size: 14px;")
        self.file_selector_btn.clicked.connect(self.select_files)
        self.layout.addWidget(self.file_selector_btn)
        
        self.output_selector_btn = QPushButton("Select Output Folder", self)
        self.output_selector_btn.setStyleSheet("padding: 10px; font-size: 14px;")
        self.output_selector_btn.clicked.connect(self.select_output_folder)
        self.layout.addWidget(self.output_selector_btn)
        
        self.convert_btn = QPushButton("Convert to PDF", self)
        self.convert_btn.setStyleSheet("padding: 10px; font-size: 16px; font-weight: bold;")
        self.convert_btn.clicked.connect(self.toggle_conversion)
        self.layout.addWidget(self.convert_btn)
        
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("QProgressBar { height: 20px; }")
        self.layout.addWidget(self.progress_bar)
        
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("background-color: #f9f9f9; color: black; font-size: 12px; padding: 5px;")
        self.layout.addWidget(self.log_text)
        
        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

        # Set up logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

        # Apply a modern Fusion style
        QApplication.setStyle("Fusion")
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QLabel {
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45A049;
            }
            QProgressBar {
                text-align: center;
                color: white;
                background-color: #D3D3D3;
                border: 1px solid #4CAF50;
                border-radius: 5px;
            }
            QTextEdit {
                border: 1px solid #D3D3D3;
                border-radius: 5px;
            }
        """)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Files", "", "Excel Files (*.xls *.xlsx)"
        )
        if files:
            self.selected_files = files
            self.log(f"Selected files: {len(files)}")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_dir = folder
            self.log(f"Selected output folder: {folder}")

    def log(self, message):
        logging.info(message)
        self.log_text.append(message)

    def toggle_conversion(self):
        if self.convert_btn.text() == "Convert to PDF":
            # Start conversion process
            if not self.selected_files:
                self.log("No files selected. Please select Excel files first.")
                return
            if not self.output_dir:
                self.log("No output folder selected. Please select an output folder.")
                return

            self.convert_btn.setText("Cancel")
            self.progress_bar.setValue(0)

            self.worker = ConversionWorker(self.selected_files, self.output_dir)
            self.worker.progress_update.connect(self.log)
            self.worker.conversion_complete.connect(self.conversion_complete)
            self.worker.progress.connect(self.update_progress)

            self.conversion_thread = threading.Thread(target=self.worker.convert_files)
            self.conversion_thread.start()
        else:
            # Cancel conversion process
            if self.worker:
                self.worker.cancel_conversion()
            self.convert_btn.setEnabled(False)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def conversion_complete(self, message):
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("Convert to PDF")
        self.progress_bar.setValue(100)
        self.log(message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_win = ExcelToPDFConverterApp()
    main_win.show()
    sys.exit(app.exec())
