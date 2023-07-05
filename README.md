## Converter
- Convertes every xls/xlsx to csv, docx to txt, pdf to txt in the specified directory 
### Dependencies

```ps
pip install -r requirements.txt
```

### Usage

```ps
python.exe main.py -h
usage: docx to txt / xlsx to csv / pdf to txt -  converter                                          

options:
  -h, --help           show this help message and exit
  --type TYPE          type of file to be converted: 'docx' / 'xlsx' / 'pdf'
  --src-path SRC_PATH  directory path to from were to convert the specified file type
```

- Convert from docx to txt

```ps
python.exe .\main.py --type="docx" --src-path="C:/"
```

- Convert from xlsx to csv

```ps
python.exe .\main.py --type="xlsx" --src-path="C:/"
```

- Convert from pdf to txt

```ps
python.exe .\main.py --type="pdf" --src-path="C:/"
```