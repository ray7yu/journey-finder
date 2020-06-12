import openpyxl, os, csv
os.chdir('./files/')
workbook = input('Filename that you want to convert: ')
wb = openpyxl.load_workbook(workbook)
sh = wb.active
workbook = workbook[:-5] + '.csv'
with open(workbook, 'w', newline="") as f:
    c = csv.writer(f)
    for r in sh.rows:
        c.writerow([cell.value for cell in r])