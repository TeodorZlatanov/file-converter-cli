import argparse
from pathlib import Path

from excel import excel_to_csv
from word import docx_to_txt
from pdf import pdf_to_txt


def main():

    parser = argparse.ArgumentParser(
        usage="docx to txt / xlsx to csv / pdf to txt -  converter"
    )

    parser.add_argument(
        "--type",
        type=str,
        required=True,
        help="type of file to be converted: 'docx' / 'xlsx' / 'pdf'"
    )

    parser.add_argument(
        "--src-path",
        type=str,
        required=True,
        help="directory path to from wh ere to convert the specified file type"
    )

    args = parser.parse_args()
    dir_path = Path(args.src_path)
    prog_type = args.type

    print(f"[INFO] function type: converting from {prog_type}")
    print(f"[INFO] directory passed to the program: {dir_path}")

    #Check if the directory exists; If Not -> throw exception
    if not dir_path.exists():
        print(f"[ERROR] The target {dir_path} directory doesn't exist!")
        raise SystemExit(1)

    if prog_type == "docx":
        docx_to_txt(dir_path=dir_path)

    elif prog_type == "xlsx":
        excel_to_csv(dir_path)

    elif prog_type == "pdf":
        pdf_to_txt(dir_path=dir_path)

    else:
        print(f"[ERROR] There is no '{prog_type}' type!")
        

if __name__=='__main__':
    main()