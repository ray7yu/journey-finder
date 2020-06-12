#! /usr/bin/env python3
'''
Raymond Yu
06/2020
Script uses the openpyxl and selenium libaries of python
in order to scrape data about points of interest and restaurants
from TripAdvisor to put into an excel worksheet
'''
import time, os
from openpyxl import Workbook
from selenium import webdriver

'''Prompt Input'''
yes = False
acceptable = ['y', 'Y', 'Yes', 'yes']
print('Outputted Excel files will be in ./files')
while not yes:
    print('What location would you like to visit?')
    print('Format: City State/Country') 
    print('Example: Dallas Texas, London United Kingdom')
    location = input('Enter: ')
    confirm = input('Are you sure? (y/n) ')
    if confirm in acceptable:
        yes = True

'''Setup driver for Chrome'''
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
search.send_keys(location)
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
if attractions[0].get_attribute('title') == 'Travel Advice':
    index = 2
else:
    index = 1
attractions[index].click()

'''Scrolls and makes more locations available'''
#Fixes major bug by checking alternate class name
see_more = driver.find_elements_by_class_name('attractions-attraction-overview-main-TopPOIs__see_more--2Vsb-')
if not see_more:
    see_more = driver.find_elements_by_class_name('_3EzqC0V4')
#scrolls to element
driver.execute_script('arguments[0].scrollIntoView();', see_more[0])
#scrolls a little up
driver.execute_script('window.scrollBy(0, -80);')
time.sleep(1)
see_more[0].click()

'''Collect attractions'''
names = []
reviews = []
urls = []
#alternate class name
anchors = driver.find_elements_by_class_name('attractions-attraction-overview-pois-PoiInfo__name--SJ0a4')
if anchors:
    for a in anchors:
        urls.append(a.get_attribute('href'))
        place = a.find_element_by_tag_name('h3')
        names.append(place.text)
else:
    anchors = driver.find_elements_by_class_name('_255i5rcQ')
    for a in anchors:
        urls.append(a.get_attribute('href'))
        names.append(a.text)
numbers = driver.find_elements_by_class_name('reviewCount')
if not numbers:
    numbers = driver.find_elements_by_class_name('_1DasOrRF')
for n in numbers:
    reviews.append(n.text)
# print(urls)
print(names)
print(reviews)

'''Collect Restaurants'''
driver.execute_script('window.scrollTo(0, 0)')
restaurant = driver.find_elements_by_id('global-nav-restaurants')
if not restaurant:
    restaurant = driver.find_elements_by_class_name('_1T4t-FiN')
    restaurant = restaurant[3]
else:
    restaurant = restaurant[0]
restaurant.click()
dining_names = []
dining_categories = []
dining_prices = []
dining_reviews = []
dining_urls = []

time.sleep(1.5)
for d in driver.find_elements_by_class_name('_15_ydu6b'):
    dining_names.append(d.text)
    dining_urls.append(d.get_attribute('href'))
dining_reviews = [r.text for r in driver.find_elements_by_class_name('w726Ki5B')]
for style in driver.find_elements_by_class_name('_3d9EnJpt'):
    stats = style.find_elements_by_class_name('EHA742uW')
    if len(stats) < 2:
        dining_categories.append('')
        dining_prices.append(stats[0].text)
    else:
        dining_categories.append(stats[0].text)
        dining_prices.append(stats[1].text)

print(dining_names)
'''Quit driver'''
driver.quit()

'''Write to Excel file'''
if names:
    os.chdir('./files/')
    wb = Workbook()
    ws = wb.active
    ws.title = 'Points of Interest'
    ws['A1'] = 'Location Name'
    ws['B1'] = 'Review Count'
    ws['C1'] = 'Links'
    for x in range(len(names)):
        ws.cell(row=x+2,column=1,value=names[x])
        ws.cell(row=x+2,column=2,value=reviews[x])
        ws.cell(row=x+2,column=3,value=urls[x])
    if dining_names:
        ws1 = wb.create_sheet('Restaurants')
        ws1['A1'] = 'Restaurant Name'
        ws1['B1'] = 'Cuisine Categories'
        ws1['C1'] = 'Price Level'
        ws1['D1'] = 'Review Count'
        ws1['E1'] = 'Links'
        for y in range(len(dining_names)):
            ws1.cell(row=y+2,column=1,value=dining_names[y])
            ws1.cell(row=y+2,column=2,value=dining_categories[y])
            ws1.cell(row=y+2,column=3,value=dining_prices[y])
            ws1.cell(row=y+2,column=4,value=dining_reviews[y])
            ws1.cell(row=y+2,column=5,value=dining_urls[y])
    excel_name = location + ' Tourism.xlsx'
    wb.save(excel_name)