import openpyxl, os, csv
os.chdir('./files/')
wb = openpyxl.load_workbook('tallahassee_tourism.xlsx')
sh = wb.active
with open('tallahassee_tourism.csv', 'w', newline="") as f:
    c = csv.writer(f)
    for r in sh.rows:
        c.writerow([cell.value for cell in r])