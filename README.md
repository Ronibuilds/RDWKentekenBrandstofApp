# RDW Kenteken Brandstof Checker

A user-friendly desktop application to extract Dutch license plates (kentekens) from PDF/Word documents, look up their fuel type using the official RDW Open Data API, and export the results in a stylish Excel report. The tool features a modern dark mode interface.

---

## Features

- **Extract License Plates:** Automatically finds Dutch license plates (format: `Kenteken: XX999X`) in PDF or Word (.docx) documents.
- **RDW API Integration:** Retrieves fuel type (brandstoftype) for each plate from the real RDW API.
- **Excel Reporting:** Generates beautifully formatted and readable Excel files, with colors per fuel type.
- **Configurable Output:** Lets you choose the output folder for saving reports.
- **Dark Mode GUI:** Beautiful modern interface with dark color scheme and clickable author attribution.
- **Logging:** Creates a log file for troubleshooting.
- **Threaded Operation:** Remains responsive during long tasks.
- **Help/Support:** Easy to use, with built-in help dialog and error messages.

---

## Installation

1. **Clone this repository:**

   ```bash
   git clone https://github.com/Ronibuilds/RDWKentekenBrandstofApp.git
   cd RDWKentekenBrandstofApp
   ```

2. **Install required dependencies:**

   ```bash
   Dependencies:

   ```
   requests
   pandas
   openpyxl
   docx2txt
   pymupdf
   pillow
   ```

   Install with:
   ```bash
   pip install requests pandas openpyxl docx2txt pymupdf pillow
   ```

3. **Run the app:**

   ```bash
   python script.py
   ```

---

## Usage

1. **Choose Output Folder:**  
   Click "Opslag Map Kiezen" to set where results are saved.
2. **Choose File:**  
   Click "Bestand Selecteren" and pick a PDF or DOCX document to analyze.
3. **Processing:**  
   The app automatically scans the document for license plates, queries the RDW API, and saves the results.
4. **Result:**  
   An Excel file is created in the chosen output folder, with found license plates and their corresponding fuel types.

For more info, click "Help" in the app.

---

## File Formats

- **Supported Input:** `.pdf`, `.docx`
- **Output:** `Excel (.xlsx)` and a plain text dump of extracted document content.

---

## Logging & Troubleshooting

- Log files are stored in the `logs` directory.
- If something goes wrong, you’ll see an error message and details are available in the logs.


---

## License

This project is licensed under the MIT License.

---

## Disclaimer

This project is not affiliated with RDW. Be mindful of the RDW open data terms of use.
