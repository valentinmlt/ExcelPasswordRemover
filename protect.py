import json
import os.path

from excel import WorkSheet, ExcelDocument
import sys

if __name__ == "__main__":

    if not len(sys.argv) == 3:
        print('USAGE : python script.py <unprotected_file.xlsx> <security_file.json>')
        exit(-1)

    UNPROTECTED_FILE_PATH = sys.argv[1]
    SAVING_FILE_PATH = sys.argv[2]

    if not os.path.exists(UNPROTECTED_FILE_PATH) or not os.path.isfile(UNPROTECTED_FILE_PATH):
        print(f"File don't exist ({UNPROTECTED_FILE_PATH})")
        exit(-1)

    if not os.path.exists(SAVING_FILE_PATH) or not os.path.isfile(SAVING_FILE_PATH):
        print(f"File don't exist ({SAVING_FILE_PATH})")
        exit(-1)

    if not os.path.splitext(UNPROTECTED_FILE_PATH)[1] == '.xlsx':
        print('File should be an Excel document (xlsx)')
        exit(-1)

    if not os.path.splitext(SAVING_FILE_PATH)[1] == '.json':
        print('File should be an json document (json)')
        exit(-1)

    excel_document = ExcelDocument(UNPROTECTED_FILE_PATH)
    with open(SAVING_FILE_PATH, 'r') as saving_file:
        SECURITY_SAVING = json.load(saving_file)

    DOCUMENT_FILE_NAME = excel_document.file_name
    DOCUMENT_NAME = excel_document.name
    PROTECTED_DOCUMENT_NAME = 'protected_' + DOCUMENT_FILE_NAME

    if os.path.exists(PROTECTED_DOCUMENT_NAME):
        print(f'There is already a unprotected document, it will be overwritten ({PROTECTED_DOCUMENT_NAME})')
        exit(-1)

    for worksheet in excel_document.get_worksheets():
        print(worksheet.get_security_str())
        security_associated = SECURITY_SAVING.get(worksheet.name, None)
        worksheet.reload_security(security_associated)
        print(worksheet.get_security_str())

    excel_document.save_file(PROTECTED_DOCUMENT_NAME)
    print(f'New protected document was created : {PROTECTED_DOCUMENT_NAME}')
