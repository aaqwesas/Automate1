import pandas as pd
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import glob
import shutil
import os
import re
import logging
from functools import wraps

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


# functional error handling
class Result:
    def __init__(self, is_ok, value):
        self.is_ok = is_ok
        self.value = value

    def __repr__(self):
        return f"Ok({self.value})" if self.is_ok else f"Error({self.value})"

    @staticmethod
    def ok(value):
        return Result(True, value)

    @staticmethod
    def error(error_value):
        return Result(False, error_value)

    def bind(self, func):
        if not self.is_ok:
            return self
        try:
            return func(self.value)
        except Exception as e:
            return Result.error(str(e))


def logging_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            logger.info(f"Result from {func.__name__}: {result}")
            return result
        except Exception as e:
            return Result.error(f"Error in {func.__name__}: {e}")

    return wrapper


# -----------------------
# Pure Helper Functions
# -----------------------


def clean_filename(name: str) -> str:
    return re.sub(r"[^\w\-_\. ]", "", name).strip()


def validate_length(names: list[str], pages: list) -> Result:
    if len(names) != len(pages):
        return Result.error(
            "The number of participant names does not match the number of pages"
        )
    return Result.ok(None)


# -----------------------
# IO / Side-Effect Functions
# -----------------------


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def remove_folder(folder_path: str) -> None:
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)


def list_files(pattern: str) -> list[str]:
    return glob.glob(pattern)


def check_pdf_path(path: str) -> Result:
    """Check if the PDF file exists."""
    if not os.path.exists(path):
        return Result.error(f"PDF file not found: {path}")
    return Result.ok(path)


# -----------------------
# PDF-Processing Functions
# -----------------------


@logging_decorator
def merge_pdfs(pdf_files: list[str], output_file: str) -> Result:
    try:
        merger = PdfMerger()
        for pdf_file in pdf_files:
            merger.append(pdf_file)
        merger.write(output_file)
        merger.close()
        return Result.ok(output_file)
    except Exception as e:
        return Result.error(f"Error merging PDFs: {e}")


def safe_split_pdf(
    combined_pdf_path: str,
    output_folder: str,
    name_list: list[str],
    course_code: str,
    prefix: str,
) -> Result:
    try:
        reader = PdfReader(combined_pdf_path)
    except Exception as e:
        return Result.error(f"Error reading PDF: {e}")

    # Validate that the number of names matches the number of PDF pages.
    res = validate_length(name_list, reader.pages)
    if not res.is_ok:
        return res

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

    try:
        # Lazy evaluation
        list(map(write_page, enumerate(name_list)))
        return Result.ok(None)
    except Exception as e:
        return Result.error(f"Error splitting PDF pages: {e}")


@logging_decorator
def split_pdf_to_pages(
    combined_pdf_path: str,
    output_folder: str,
    name_list: list[str],
    course_code: str,
    prefix: str,
) -> Result:
    return check_pdf_path(combined_pdf_path).bind(
        lambda _: safe_split_pdf(
            combined_pdf_path, output_folder, name_list, course_code, prefix
        )
    )


# -----------------------
# CSV and File List Processing
# -----------------------


def extract_names_from_csv(csv_file: str, start_row: int = 0) -> list[str]:
    df = pd.read_csv(csv_file)
    return [str(row["User Name"]).strip() for _, row in df.iloc[start_row:].iterrows()]


# -----------------------
# ZIP Archive Function
# -----------------------


def create_zip_from_folder(folder: str) -> Result:
    """Create a ZIP archive of the specified folder if it contains files."""
    if os.path.exists(folder) and os.listdir(folder):
        try:
            zip_path = shutil.make_archive(folder, "zip", folder)
            return Result.ok(zip_path)
        except Exception as e:
            return Result.error(f"Error creating zip archive for {folder}: {e}")
    return Result.error(f"No files found in {folder}")


# -----------------------
# Main Pipeline Function
# -----------------------


@logging_decorator
def main() -> Result | None:
    # 1. Get the CSV file and extract names.
    csv_files = list_files(os.path.join("data", "*.csv"))
    if not csv_files:
        return Result.error("No CSV file found")
    csv_file = csv_files[0]
    name_list = extract_names_from_csv(csv_file)

    # 2. Get the course code from user input and define output folder names.
    course_code = input("Enter the course code: ").strip()
    output_folders = {
        "cert": f"Certificate_{course_code}",
        "receipt": f"Receipt_{course_code}",
        "combined": f"combined_{course_code}",
    }

    # Create all necessary directories.
    for folder in output_folders.values():
        ensure_dir(folder)

    # 3. Get lists of certificate and receipt PDF files.
    cert_list = list_files(os.path.join("data", "Certificate", "*.pdf"))
    receipt_list = list_files(os.path.join("data", "Receipt", "*.pdf"))

    # 4. Merge PDFs for certificates and receipts.
    combined_cert_file = os.path.join(
        output_folders["combined"], "combined_certificates.pdf"
    )
    combined_receipt_file = os.path.join(
        output_folders["combined"], "combined_receipts.pdf"
    )

    res_cert_merge = merge_pdfs(cert_list, combined_cert_file)
    if not res_cert_merge.is_ok:
        return Result.error(f"Error merging certificates: {res_cert_merge.value}")

    res_receipt_merge = merge_pdfs(receipt_list, combined_receipt_file)
    if not res_receipt_merge.is_ok:
        return Result.error(f"Error merging receipts: {res_receipt_merge.value}")

    # 5. Split the combined PDFs into individual pages.
    res_cert_split = split_pdf_to_pages(
        combined_cert_file,
        output_folders["cert"],
        name_list,
        course_code,
        "Certificate",
    )
    if not res_cert_split.is_ok:
        return Result.error(f"Error splitting certificate PDF: {res_cert_split.value}")

    res_receipt_split = split_pdf_to_pages(
        combined_receipt_file,
        output_folders["receipt"],
        name_list,
        course_code,
        "Receipt",
    )
    if not res_receipt_split.is_ok:
        return Result.error(f"Error splitting receipt PDF: {res_receipt_split.value}")

    # 6. Archive each of the output folders.
    for folder in output_folders.values():
        if folder == output_folders["combined"]:
            continue
        res_zip = create_zip_from_folder(folder)
        if res_zip.is_ok:
            logger.info(f"Created zip archive: {res_zip.value}")
        else:
            logger.warning(
                f"ZIP archive creation skipped for {folder}: {res_zip.value}"
            )

    # 7. Clean up the temporary folders
    for folder in output_folders.values():
        remove_folder(folder)


if __name__ == "__main__":
    main()
