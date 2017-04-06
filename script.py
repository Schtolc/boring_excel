from openpyxl import Workbook, load_workbook

wb = load_workbook('input.xlsx')
input = wb.active

wb2 = load_workbook('output.xlsx')
output = wb2.active

for article in output['E5':'E3646']:
    art = str(article[0].value).strip()
    for article2 in input['B4':'B4798']:
        art2 = str(article2[0].value).strip()
        if art == art2:
            output['G' + str(article[0].row)].value = str(input['B' + str(article2[0].row)].value).strip()
            output['H' + str(article[0].row)].value = str(input['C' + str(article2[0].row)].value).strip()
            output['I' + str(article[0].row)].value = str(input['D' + str(article2[0].row)].value).strip()
            tmp = str(input['E' + str(article2[0].row)].value).strip()
            output['J' + str(article[0].row)].value = '' if tmp == 'None' else tmp
            break

wb2.save('result.xlsx')

