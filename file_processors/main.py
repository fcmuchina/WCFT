from openpyxl import load_workbook, workbook


class Dependencies:
    def __init__(self, orgin_data_file):
        self.orgin_data_file = orgin_data_file
    load_workbook_func = load_workbook
    create_workbook_func = workbook



class ExcelFileOps:
    def __init__(self, dependencies):
        self.load_workbook_func = dependencies.load_workbook_func
        self.orgin_data_file = dependencies.orgin_data_file
        self.create_workbook_func = dependencies.create_workbook_func


    def load_origin_data_file(self):
        workbook = self.load_workbook_func(self.orgin_data_file)
        return workbook

    def get_sheet_names(self, workbook):
        return workbook.sheetnames

    def create_new_workbook(self):
        return self.create_workbook_func()

    def create_worksheet(self, workbook, sheet_name):
        return workbook.create_sheet(sheet_name)

    def save_workbook(self, workbook, file_name):
        workbook.save(file_name)


    

    

    