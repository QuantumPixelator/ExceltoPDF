
# Excel to PDF Converter

This project provides a Windows-based application that converts Excel files (.xls and .xlsx) to PDF format. The application uses the `win32com` library to interface with Microsoft Excel, allowing for high-quality PDF exports directly from Excel files. Built with PySide6, it provides a simple, user-friendly GUI to select files, choose an output folder, and monitor conversion progress.

## Features

- **Batch Conversion**: Convert multiple Excel files at once.
- **Output Folder Selection**: Specify the destination folder for PDFs.
- **Progress Tracking**: Visual progress bar and log messages for each file.
- **Renaming Options**: Automatically removes `%20` from filenames and unwanted Excel extensions (`.xls` or `.xlsx`).

## Requirements

To run this application, install the dependencies listed in `requirements.txt`:

```text
pypiwin32==223
PySide6==6.8.0.2
PySide6_Addons==6.8.0.2
PySide6_Essentials==6.8.0.2
pywin32==308
shiboken6==6.8.0.2
```

Install the requirements using pip:

```bash
pip install -r requirements.txt
```

## Getting Started

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/excel-to-pdf-converter.git
   cd excel-to-pdf-converter
   ```

2. **Install dependencies**:
   Run the following command to install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Application**:
   Launch the application by running the `main.pyw` file:
   ```bash
   python main.pyw
   ```

4. **Using the Application**:
   - Select Excel files for conversion by clicking "Browse Files."
   - Choose an output directory with "Select Output Folder."
   - Click "Convert to PDF" to start the batch conversion.

## File Renaming and Path Management

- During conversion, the application ensures file paths are Windows-compatible.
- Files are saved with `.pdf` as the sole extension, removing any preceding `.xls` or `.xlsx` extensions.
- Spaces replace URL-encoded `%20` in file names.

## Known Issues

- This application requires Microsoft Excel to be installed on the machine, as it uses `win32com` to interact with Excel.
- Currently designed only for Windows due to reliance on Windows-specific libraries (`win32com`).

## License

MIT License. See `LICENSE` file for more information.
