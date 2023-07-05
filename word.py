import os

from docx import Document

WORD_DIR_NAME = "word_to_txt"

def docx_to_txt(dir_path):
    """
    This function converts all .docx files in the specified directory into .txt files
    and store them into a newly created directory called word_to_txt.
    """
    new_dir_path = os.path.join(dir_path, WORD_DIR_NAME)
    os.makedirs(new_dir_path, exist_ok=True)
    print(f"[INFO] directory {new_dir_path} created")

    found_docx_file = False

    for filename in os.listdir(dir_path):
        # Checks if file name ends with .DOCX
        if filename.endswith(".docx"):

            found_docx_file = True
    
            file_path = os.path.join(dir_path, filename)

            # Specifying the output path for the TXT file
            txt_file_path = os.path.join(new_dir_path, filename.replace(".docx", ".txt"))

            # Opening the DOCX file
            doc = Document(file_path)

            # Getting the text from each paragraph
            text_content = []
            for paragraph in doc.paragraphs:
                text_content.append(paragraph.text)

            # Joining the text content 
            text_content = "\n".join(text_content)
            # Writing the text content to a TXT file
            with open(txt_file_path, "w", encoding="utf-8-sig") as file:
                file.write(text_content)
            # Reporting the created file
            print(f"[INFO] file {txt_file_path} created")
        
    if not found_docx_file:
        print(f"[WARN] there are no .docx files in {dir_path}")