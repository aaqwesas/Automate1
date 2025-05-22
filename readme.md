# PDF Certificate & Receipt Automation

This project automates the process of merging and splitting PDF files (certificates and receipts) based on a provided Excel file containing participant information. It is useful for bulk operations such as matching certificates/receipts to participant names after a training or event.

---

## **Table of Contents**

- [Features](#features)
- [Directory Structure](#directory-structure)
- [Setup Instructions](#setup-instructions)
- [How to Run](#how-to-run)
- [Input File Descriptions](#input-file-descriptions)
- [Output File Descriptions](#output-file-descriptions)
- [Troubleshooting](#troubleshooting)
- [Customization](#customization)

---

## **Features**

- Merges all certificate PDFs in `data/Certificate` into a single PDF.
- Merges all receipt PDFs in `data/Receipt` into a single PDF.
- Splits the merged PDFs into individual files, named according to the participant names in the Excel file.
- Output files are saved in `Certificate_Result` and `Receipt_Result` folders.

---

## **Directory Structure**

```
project_root/
├── data/
│   ├── Certificate/
│   │   ├── cert1.pdf
│   │   ├── cert2.pdf
│   │   └── ...
│   ├── Receipt/
│   │   ├── receipt1.pdf
│   │   ├── receipt2.pdf
│   │   └── ...
│   └── whatevername.csv
├── main.py
└── README.md
```

---

## **Setup Instructions**

### 1. **Clone or Download This Repository**

```bash
git clone https://github.com/aaqwesas/Automate1.git
cd Automate1/
```

### 2. **Create a Python Virtual Environment**

```bash
python -m venv venv
```

### 3. **Activate the Virtual Environment**

- **Windows:**
  ```bash
  venv\Scripts\activate
  ```
- **macOS/Linux:**
  ```bash
  source venv/bin/activate
  ```

### 4. **Install Required Packages**

```bash
pip freeze > requirements.txt
pip install -r requirements.txt
```

---

## **How to Run**

1. **Prepare your input files:**

   - Place all certificate PDFs in `data/Certificate/`
   - Place all receipt PDFs in `data/Receipt/`
   - Place your Excel file (e.g., `whatevername.csv`) in `data/`
     - The Excel file should have two columns: one for labels (like "Course Code") and one for values (participant name).

2. **Make sure your directory structure matches [Directory Structure](#directory-structure).**

3. **Run the script:**

```bash
python main.py
```

- The script will create `combined/combined_certificates.pdf` and `combined/combined_receipts.pdf`.
- It will then split those into individual PDFs in `Certificate_Result/` and `Receipt_Result/`.

---

## **Input File Descriptions**

- **Excel File (`data/whatevername.csv`):**

  - Column 1: Values (e.g., course code, participant names)
  - Column 2: Field names (e.g., "Course Code: ", "Name", etc.)

- **Certificate PDFs:**  
  Place each participant’s certificate PDF in `data/Certificate/`.

- **Receipt PDFs:**  
  Place each participant’s receipt PDF in `data/Receipt/`.

---

## **Output File Descriptions**

- **Combined PDFs** (in `combined/`):

  - `combined_certificates.pdf`: All certificate PDFs merged together.
  - `combined_receipts.pdf`: All receipt PDFs merged together.

- **Individual Result PDFs** (in `Certificate_Result/` and `Receipt_Result/`):
  - Each PDF is named in the format:  
    `<ParticipantName>_<CourseCode>_Certificate.pdf`  
    `<ParticipantName>_<CourseCode>_Receipt.pdf`

---

## **Troubleshooting**

- **No Excel files found:**  
  Make sure your `.csv` file is in the `data/` folder.

- **Mismatch between number of names and PDF pages:**  
  Ensure the number of participant names in the Excel file matches the number of pages in the merged PDF.

- **Permission errors:**  
  Make sure you have write permissions for the output folders.

- **Other issues:**
  Make sure there is only one Excel file in the `data/` folder.

---

## **Customization**

- **Change output folders:**  
  Modify the `output_folder` arguments in the script as desired.

- **Change file naming scheme:**  
  Edit the filename format in the `split_pdf_to_pages` function in `main.py`.
