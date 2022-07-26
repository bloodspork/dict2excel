# author: bloodspork
# date  : 20220726

from openpyxl import Workbook

def dict2excel(dct: dict, filepath: str):
    wb = Workbook()
    for sheetname in dct:
        sheet = wb.create_sheet(sheetname)
        row_list = dct[sheetname]

        fieldname_pos_dict = {}
        colnum = 1
        rownum = 1
        for row in row_list:
            if rownum == 1:
                for fieldname in row:
                    sheet.cell(rownum, colnum, fieldname)
                    fieldname_pos_dict['fieldname'] = colnum
                    colnum += 1
                rownum += 1
            
            for fieldname in row:
                sheet.cell(rownum, colnum, row[fieldname])
            rownum += 1

    wb.save(filepath)

