import os
import openpyxl

for excel in os.listdir():
    if excel.endswith('.xlsx'):
        videoFolder = f'{excel[:-10]}_videos'
        wbObj = openpyxl.load_workbook(excel)
        sheet = wbObj.active
        print(f'LOOKING INTO {excel}, FILES TO BE REMOVED FROM {videoFolder}/:')
        for i in range(2, sheet.max_row + 1):
            if sheet[f"F{i}"].value == '+':
                fileToRemove = sheet[f'B{i}'].value
                print(fileToRemove)
                try:
                    os.remove(fileToRemove)
                    sheet[f"F{i}"].value = 'DELETED'
                    sheet.row_dimensions[i].height = 1
                    print(f'Successfully deleted')
                except Exception as e:
                    print(e)
                    sheet[f"F{i}"].value = str(e)
        wbObj.save(excel)

