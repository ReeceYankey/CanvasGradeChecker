from GradeSheets import *
import pandas as pd
import openpyxl as xl


def is_weighted_sheet(worksheet):
    return '(WP)' in worksheet['A2'].value


def update_from_csv(class_names):
    wb = xl.load_workbook(filename='Galipatia Database Template.xlsx')
    try:
        for name in class_names:
            table = pd.read_csv('ClassData\\{}.csv'.format(name))
            print(type(table['type']))
            # create new sheet if it doesn't exist
            if name not in wb:
                inp = input("does " + name + " use weighted grades (w) or a point system (p)?")
                if inp == 'p':
                    ws = wb.copy_worksheet("Point Template")
                    ws.title = name
                else:
                    ws = wb.copy_worksheet("Weighted Template")
                    ws.title = name
            else:
                ws = wb[name]
            # check type of point system
            if is_weighted_sheet(ws):
                sheet = WeightedSheetHandler(ws)
            else:
                sheet = PointSheetHandler(ws)
            sheet.update(table)

    finally:
        wb.save('updated.xlsx')


if __name__ == '__main__':
    class_names = ['ENGE 1215', 'ENGR 1054', 'CHEM 1035', 'MATH 2204', 'CHEM 1045', 'GEOG 1014']
    update_from_csv(class_names)
