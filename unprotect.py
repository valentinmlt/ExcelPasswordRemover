import json
import os.path

from excel import WorkSheet, ExcelDocument
import sys

if __name__ == "__main__":

    if not len(sys.argv) == 2:
        print('USAGE : python script.py <file.xlsx>')
        exit(-1)

    FILE_PATH = sys.argv[1]

    if not os.path.exists(FILE_PATH) or not os.path.isfile(FILE_PATH):
        print(f"File don't exist ({FILE_PATH})")
        exit(-1)

    if not os.path.splitext(FILE_PATH)[1] == '.xlsx':
        print('File should be an Excel document (xlsx)')
        exit(-1)

    excel_document = ExcelDocument(FILE_PATH)

    DOCUMENT_FILE_NAME = excel_document.file_name
    DOCUMENT_NAME = excel_document.name
    UNPROTECTED_DOCUMENT_NAME = 'unprotected_' + DOCUMENT_FILE_NAME
    SAVING_FILE_NAME = f'security_{DOCUMENT_NAME}.json'

    if os.path.exists(UNPROTECTED_DOCUMENT_NAME):
        print(f'There is already a unprotected document, it will be overwritten ({UNPROTECTED_DOCUMENT_NAME})')
        exit(-1)

    if os.path.exists(SAVING_FILE_NAME):
        print(f'There is already a saving document, it will be overwritten ({SAVING_FILE_NAME})')
        exit(-1)

    securitys_saver = {}

    for worksheet in excel_document.get_worksheets():
        saved_security = worksheet.remove_security()
        worksheet_name = worksheet.name

        securitys_saver[worksheet_name] = saved_security

        if saved_security:
            print(f'Security removed from {worksheet_name}')
        else:
            print(f'No security needed to be removed from {worksheet_name}')

    with open(SAVING_FILE_NAME, 'w') as saving_file:
        json.dump(securitys_saver, saving_file)
        print(f'security file backup created : {SAVING_FILE_NAME}')

    excel_document.save_file(UNPROTECTED_DOCUMENT_NAME)
    print(f'New unprotected document was created : {UNPROTECTED_DOCUMENT_NAME}')
