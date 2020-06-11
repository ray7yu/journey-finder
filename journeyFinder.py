#! /usr/bin/env python3
import time, os
from openpyxl import Workbook
from selenium import webdriver

driver = webdriver.Chrome('/usr/local/bin/chromedriver') 
driver.set_window_size(1200, 700)
driver.implicitly_wait(10)
driver.get('http://www.tripadvisor.com/') 

#Major issue was that first search bar was hidden and not interactable
#Search bar
elements = driver.find_elements_by_name('q')
# for i in elements:
#     print(i.get_attribute('aria-label'))
search = elements[1]
search.click()
search.send_keys('Tallahassee')
search.submit()

'''Overview of City'''
results = driver.find_elements_by_class_name('result-title')
city = results[0]
city.click()

'''Go to Things To Do'''
allWindows = driver.window_handles
parentWindow = allWindows[0]
currWindow = allWindows[1]
driver.switch_to.window(currWindow)

'''Go to attractions'''
attractions = driver.find_elements_by_class_name('_1yB-kafB')
if not attractions:
    print('empty')
attractions[1].click()

'''Scrolls and makes more locations available'''
print(driver.window_handles)
time.sleep(3)
see_more = driver.find_element_by_class_name('attractions-attraction-overview-main-TopPOIs__see_more--2Vsb-')
#scrolls to element
driver.execute_script("arguments[0].scrollIntoView();", see_more)
#scrolls a little up
driver.execute_script("window.scrollBy(0, -80);")
time.sleep(3)
see_more.click()

'''Collect attractions'''
names = []
reviews = []
urls = []
anchors = driver.find_elements_by_class_name('attractions-attraction-overview-pois-PoiInfo__name--SJ0a4')
for a in anchors:
    urls.append(a.get_attribute('href'))
    place = a.find_element_by_tag_name('h3')
    names.append(place.text)
numbers = driver.find_elements_by_class_name('reviewCount')
for n in numbers:
    reviews.append(n.text)
print(urls)
print(names)
print(reviews)

'''End Scraper'''
time.sleep(3) # Let the user actually see something!
# driver.quit()

'''Write to Excel file'''
if names:
    os.chdir('./files/')
    wb = Workbook()
    ws = wb.active
    ws.title = "Tallahassee Points of Interest"
    ws['A1'] = "Location Names"
    ws['B1'] = "Review Count"
    ws['C1'] = "Links"
    for x in range(len(names)):
        ws.cell(row=x+2,column=1,value=names[x])
        ws.cell(row=x+2,column=2,value=reviews[x])
        ws.cell(row=x+2,column=3,value=urls[x])
    wb.save('tallahassee_tourism.xlsx')