import os
import zipfile
import xml

from xml.dom import minidom

import tempfile


class ExcelDocument:
    def __init__(self, path):

        if not os.path.exists(path):
            raise FileNotFoundError

        if not os.path.splitext(path)[1] == '.xlsx':
            raise NotAnXlsxFileError

        self.path = path

        self.file_name = os.path.basename(path)              # /home/dir/finance.xlsx => finance.xlsx
        self.name = os.path.splitext(self.file_name)[0]      # /home/dir/finance.xlsx => finance

        self.extract_folder = os.path.join(tempfile.gettempdir(), self.name )

        with zipfile.ZipFile(self.path, 'r') as zip_file:
            zip_file.extractall(self.extract_folder)

    def get_worksheets(self):
        return_list = []
        worksheet_directory = os.path.join(self.extract_folder, 'xl/worksheets/')

        if not os.path.exists(worksheet_directory):
            print(worksheet_directory)
            raise WorkSheetDirectoryError

        if not os.path.isdir(worksheet_directory):
            raise WorkSheetDirectoryError

        for file_name in os.listdir(worksheet_directory):

            absolut_path = os.path.join(worksheet_directory, file_name)

            if os.path.isfile(absolut_path):
                return_list.append(WorkSheet(absolut_path))

        return return_list

    def save_file(self, file_name):
        with zipfile.ZipFile(file_name, 'w') as updated_zip_file:
            for root, _, files in os.walk(self.extract_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, self.extract_folder)
                    updated_zip_file.write(file_path, arcname)


class WorkSheet:
    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError

        if not os.path.isfile(file_path):
            raise FileExistsError

        self.path = file_path

        file_name = os.path.basename(file_path)              # /home/dir/finance.xlsx => finance.xlsx
        self.name = os.path.splitext(file_name)[0]          # /home/dir/finance.xlsx => finance

        self.dom = minidom.parse(self.path).documentElement

    def remove_security(self):

        return_protection = self.get_security_str()

        dom = self.dom
        protections = dom.getElementsByTagName('sheetProtection')

        if not protections:
            return

        for protection in protections:
            protection.parentNode.removeChild(protection)

        self.rewrite_file()

        return return_protection

    def reload_security(self, security_node_str):

        dom = self.dom
        protections = dom.getElementsByTagName('sheetProtection')

        if protections:
            print('THIS SHOULD NOT HAPPEN')
            self.remove_security(self.path)

        new_security = xml.dom.minidom.parseString(security_node_str).documentElement

        dom.insertBefore(new_security, dom.childNodes[-1])      # Put the security right in -2 position

        self.rewrite_file()

    def get_security_str(self):
        dom = self.dom
        protections = dom.getElementsByTagName('sheetProtection')

        if protections:
            return protections[0].toxml()
        else:
            return None

    def rewrite_file(self):
        xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

        with open(self.path, "w") as file_obj:
            file_obj.write(xml_declaration)
            self.dom.writexml(file_obj)





# ERROR

class NotAnXlsxFileError(Exception):
    pass


class WorkSheetDirectoryError(Exception):
    pass