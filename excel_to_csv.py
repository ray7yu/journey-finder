import openpyxl, os, csv
os.chdir('./files/')
workbook = input('Filename that you want to convert: ')
wb = openpyxl.load_workbook(workbook)
'''First Excel Sheet'''
sh = wb.active

'''New names for files'''
attractions = workbook[:-5] + ' Points of Interest.csv'
restaurants = workbook[:-5] + ' Restaurants.csv'
with open(attractions, 'w', newline="") as f:
    c = csv.writer(f)
    for r in sh.rows:
        c.writerow([cell.value for cell in r])
'''Second Excel Sheet'''
sh = wb['Restaurants']
with open(restaurants, 'w', newline="") as f:
    c = csv.writer(f)
    for r in sh.rows:
        c.writerow([cell.value for cell in r])