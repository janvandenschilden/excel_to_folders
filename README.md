# excel_to_folders
# libraries
```bash
pip install pandas openpyxl
```

## To get help
```bash
python excel_to_folders.py -h
``` 

## Full command
```bash
python excel_to_folders.py \
    [--excel_file EXCEL_FILE] \
    [--sheet_name SHEET_NAME] \
    [--output_folder OUTPUT_FOLDER] \
    [--document_folder DOCUMENT_FOLDER]
```

## Defaults
```bash
python excel_to_folders.py
```

will do the same as

```bash
python excel_to_folders.py --excel_file example.xlsx --sheet_name Sheet1 --output_folder output --document_folder documents
```
