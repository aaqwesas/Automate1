import pandas as pd
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import glob
import shutil
import os
import re
import logging
from functools import wraps

logging.basicConfig(
    level=logging.ERROR,
    filename="app_errors.log",
    filemode="w",
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


def logging_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            logger.info(f"Running {func.__name__} with args: {args}, kwargs: {kwargs}")
            result = func(*args, **kwargs)
            logger.info(f"Result from {func.__name__}: {result}")
            return result
        except Exception as e:
            logger.error(f"Error in {func.__name__}: {e}")
            raise

    return wrapper


# -----------------------
# Pure Helper Functions
# -----------------------


def clean_filename(name: str) -> str:
    return re.sub(r"[^\w\-_\. ]", "", name).strip()


def validate_length(names: list[str], pages: list) -> None:
    if len(names) != len(pages):
        raise ValueError(
            f"Length of name_list ({len(names)}) does not match number of PDF pages ({len(pages)})"
        )


# -----------------------
# IO / Side-Effect Functions
# -----------------------


def ensure_dir(path: str) -> None:
    """Ensure that a directory exists."""
    os.makedirs(path, exist_ok=True)


def remove_folder(folder_path: str) -> None:
    """Remove a folder if it exists."""
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)


def list_files(pattern: str) -> list[str]:
    """Return a list of file paths matching the given glob pattern."""
    return glob.glob(pattern)


def check_pdf_path(path: str) -> bool:
    """Check if the PDF file exists."""
    if not os.path.exists(path):
        logger.info(f"No PDF file found in {path}.")
        return False
    return True


# -----------------------
# PDF-Processing Functions
# -----------------------


@logging_decorator
def merge_pdfs(pdf_files: list[str], output_file: str) -> str:
    """Merge a list of PDF files into one PDF file."""
    ensure_dir(os.path.dirname(output_file))
    if not pdf_files:
        print("No PDF files found to merge.")
        return ""
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    merger.write(output_file)
    merger.close()
    return output_file


@logging_decorator
def split_pdf_to_pages(
    combined_pdf_path: str,
    output_folder: str,
    name_list: list[str],
    course_code: str,
    prefix: str,
) -> list[str]:
    """
    Split a combined PDF into individual files by page.
    Each file is named using a value from name_list along with course_code and prefix.
    """
    if not check_pdf_path(combined_pdf_path):
        return []

    ensure_dir(output_folder)
    reader = PdfReader(combined_pdf_path)
    validate_length(name_list, reader.pages)

    def write_page(page_info):
        i, name = page_info
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        safe_name = clean_filename(str(name))
        output_filename = os.path.join(
            output_folder, f"{safe_name}_{course_code}_{prefix}.pdf"
        )
        with open(output_filename, "wb") as f:
            writer.write(f)
        return output_filename

    return list(map(write_page, enumerate(name_list)))


# -----------------------
# CSV and File List Processing
# -----------------------


def extract_names_from_csv(csv_file: str, start_row: int = 0) -> list[str]:
    """Extract the 'User Name' field from CSV rows starting at start_row."""
    df = pd.read_csv(csv_file)
    return [str(row["User Name"]).strip() for _, row in df.iloc[start_row:].iterrows()]


# -----------------------
# ZIP Archive Function
# -----------------------


@logging_decorator
def create_zip_from_folder(folder: str) -> str:
    """Create a ZIP archive of the specified folder if it contains files."""
    if os.path.exists(folder) and os.listdir(folder):
        return shutil.make_archive(folder, "zip", folder)
    return ""


# -----------------------
# Main Pipeline Function
# -----------------------


def main():
    # 1. Get the CSV file and extract names.
    csv_files = list_files(os.path.join("data", "*.csv"))
    if not csv_files:
        print("No CSV files found in the 'data' folder.")
        return
    csv_file = csv_files[0]
    name_list = extract_names_from_csv(csv_file)

    # 2. Get the course code from user input.
    course_code = input("Enter the course code: ").strip()

    # 3. Create lists of certificate and receipt PDF files.
    cert_list = list_files(os.path.join("data", "Certificate", "*.pdf"))
    receipt_list = list_files(os.path.join("data", "Receipt", "*.pdf"))

    # 4. Merge PDFs for certificates and receipts.
    combined_cert = merge_pdfs(
        cert_list, os.path.join("combined", "combined_certificates.pdf")
    )
    combined_receipt = merge_pdfs(
        receipt_list, os.path.join("combined", "combined_receipts.pdf")
    )

    # 5. Split the combined PDFs into individual pages using the name list.
    split_pdf_to_pages(
        combined_pdf_path=combined_cert,
        output_folder="Certificate_Result",
        name_list=name_list,
        course_code=course_code,
        prefix="Certificate",
    )
    split_pdf_to_pages(
        combined_pdf_path=combined_receipt,
        output_folder="Receipt_Result",
        name_list=name_list,
        course_code=course_code,
        prefix="Receipt",
    )

    # 6. Archive the result folders as ZIP files.
    for folder in ("Certificate_Result", "Receipt_Result"):
        zip_archive = create_zip_from_folder(folder)
        if zip_archive:
            print(f"Created zip archive: {zip_archive}")
        else:
            print(f"No files found in {folder}. Skipping zip creation.")

    # 7. Clean up temporary folders.
    for folder in ("combined", "Certificate_Result", "Receipt_Result"):
        remove_folder(folder)


if __name__ == "__main__":
    main()
