import pandas as pd
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import glob
import shutil
import os
import re
import logging


logging.basicConfig(
    level=logging.ERROR,
    filename="app_errors.log",
    filemode="w",
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def error_handler(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            # Unified error handling
            logging.error(
                f"An error occurred in '{func.__name__}': {str(e)}", exc_info=True
            )
            raise e

    return wrapper


@error_handler
def split_pdf_to_pages(
    combined_pdf_path, output_folder, name_list, course_code, prefix
):
    if not os.path.exists(combined_pdf_path):
        print(f"No combined PDF file found in {combined_pdf_path}.")
        return

    os.makedirs(output_folder, exist_ok=True)
    reader = PdfReader(combined_pdf_path)

    if len(name_list) != len(reader.pages):
        raise ValueError(
            f"Length of name_list ({len(name_list)}) does not match number of PDF pages ({len(reader.pages)})"
        )

    output_files = []
    for page_num, name in enumerate(name_list):
        writer = PdfWriter()
        writer.add_page(reader.pages[page_num])
        safe_name = re.sub(r"[^\w\-_\. ]", "", str(name)).strip()
        output_filename = os.path.join(
            output_folder, f"{safe_name}_{course_code}_{prefix}.pdf"
        )
        with open(output_filename, "wb") as out_file:
            writer.write(out_file)
        output_files.append(output_filename)
        print(f"Saved: {output_filename}")
    return output_files


@error_handler
def merge_pdfs(pdf_files, output_file):
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    if not pdf_files:
        print("No PDF files found in the given path")
        return
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    merger.write(output_file)
    merger.close()


@error_handler
def main():
    csv_files = glob.glob("data/*.csv")  # this will automatically match your csv file
    if not csv_files:
        print("No CSV files found in the 'data' folder.")
        return
    file_path = csv_files[0]
    df = pd.read_csv(file_path)
    course_code = df["Course Code"].iloc[0]
    start_row = 0  # this is the row where the participant names start, modify as needed
    name_list = []
    name_list = [row["User Name"].strip() for _, row in df.iloc[start_row:].iterrows()]

    cert_list = sorted(glob.glob("data/Certificate/*.pdf"))
    receipt_list = sorted(glob.glob("data/Receipt/*.pdf"))

    merge_pdfs(pdf_files=cert_list, output_file="combined/combined_certificates.pdf")
    merge_pdfs(pdf_files=receipt_list, output_file="combined/combined_receipts.pdf")

    split_pdf_to_pages(
        combined_pdf_path="combined/combined_certificates.pdf",
        output_folder="Certificate_Result",
        name_list=name_list,
        course_code=course_code,
        prefix="Certificate",
    )

    split_pdf_to_pages(
        combined_pdf_path="combined/combined_receipts.pdf",
        output_folder="Receipt_Result",
        name_list=name_list,
        course_code=course_code,
        prefix="Receipt",
    )

    # Convert Certificate_Result to a zip file
    for folder in ("Certificate_Result", "Receipt_Result"):
        if os.path.exists(folder) and os.listdir(folder):
            shutil.make_archive(folder, "zip", folder)
            print(f"Created {folder}.zip")
        else:
            print(f"No files found in {folder} folder. Skipping zip creation.")

    # if os.path.exists("combined"):
    #     shutil.rmtree("combined")
    #     print("Deleted combined folder")

    print("Done!")


if __name__ == "__main__":
    main()
