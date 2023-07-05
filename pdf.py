import os
import PyPDF2


PDF_DIR_NAME = "pdf_to_txt"

def pdf_to_txt(dir_path):
    """
    This function converts all .pdf files in the specified directory into .txt files
    and store them into a newly created directory called pdf_to_txt.
    """
    new_dir_path = os.path.join(dir_path, PDF_DIR_NAME)
    os.makedirs(new_dir_path, exist_ok=True)
    print(f"[INFO] directory {new_dir_path} created")

    found_pdf_file = False

    for filename in os.listdir(dir_path):
        # Checks if file name ends with .PDF
        if filename.endswith(".pdf"):

            found_pdf_file = True

            pdf_file_path = os.path.join(dir_path, filename)

            # Specifying the output path for the TXT file
            txt_file_path = os.path.join(new_dir_path, filename.replace(".pdf", ".txt"))

            # Creating a PDF file object
            pdfFileObj = open(pdf_file_path, "rb")

            # Creating a PDF reader object
            pdf_reader = PyPDF2.PdfReader(pdfFileObj)

            # Getting the number of pages in the PDF file
            pdf_file_pages = len(pdf_reader.pages)

            text_content = []

            # Looping through each of the pages and extracting the text
            for page_num in range(0, pdf_file_pages):

                pdf_file_page= pdf_reader._get_page(page_num)
                text = pdf_file_page.extract_text()
                text_content.append(text)

            # Joining the text content
            text_content = "\n".join(text_content)

            # Writing the text content to a TXT file
            with open(txt_file_path, "w", encoding="utf-8-sig") as file:
                file.write(text_content)
            # Reporting the created file
            print(f"[INFO] file {txt_file_path} created")
    
    if not found_pdf_file:
        print(f"[WARN] there are no .pdf files in {dir_path}")