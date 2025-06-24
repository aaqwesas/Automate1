# PDF Certificate & Receipt Automation
Personal use only, it is specific for automating some trivial jobs I need to perform in day-to-day. No sure if you really can apply it to anything else.

---

## **Table of Contents**

- [Features](#features)
- [Directory Structure](#directory-structure)
- [Setup Instructions](#setup-instructions)

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
pip install -r requirements.txt
```

---


