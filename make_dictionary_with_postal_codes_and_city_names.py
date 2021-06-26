# This program only needs to run once to make a dictionary with all the postal codes and their city names
# If there are more cities with the same postal code, only the last one will be added (because adding an element with
#   existing key will overwrite the value.

from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys

dictionary_postal_codes_and_city_names = dict()

# Add the cities of Brussels by hand...
dictionary_postal_codes_and_city_names[1000] = 'Brussel'
dictionary_postal_codes_and_city_names[1070] = 'Anderlecht'
dictionary_postal_codes_and_city_names[1050] = 'Elsene'
dictionary_postal_codes_and_city_names[1040] = 'Etterbeek'
dictionary_postal_codes_and_city_names[1140] = 'Evere'
dictionary_postal_codes_and_city_names[1083] = 'Ganshoren'
dictionary_postal_codes_and_city_names[1130] = 'Haren'
dictionary_postal_codes_and_city_names[1090] = 'Jette'
dictionary_postal_codes_and_city_names[1081] = 'Koekelberg'
dictionary_postal_codes_and_city_names[1020] = 'Laken'
dictionary_postal_codes_and_city_names[1120] = 'Neder-Over-Heembeek'
dictionary_postal_codes_and_city_names[1160] = 'Oudergem'
dictionary_postal_codes_and_city_names[1030] = 'Schaarbeek'
dictionary_postal_codes_and_city_names[1082] = 'Sint-Agatha-Berchem'
dictionary_postal_codes_and_city_names[1060] = 'Sint-Gillis'
dictionary_postal_codes_and_city_names[1080] = 'Sint-Jans-Molenbeek'
dictionary_postal_codes_and_city_names[1210] = 'Sint-Joost-Ten-Node'
dictionary_postal_codes_and_city_names[1200] = 'Sint-Lambrechts-Woluwe'
dictionary_postal_codes_and_city_names[1150] = 'Sint-Pieters-Woluwe'
dictionary_postal_codes_and_city_names[1180] = 'Ukkel'
dictionary_postal_codes_and_city_names[1190] = 'Vorst'
dictionary_postal_codes_and_city_names[1170] = 'Watermaal-Bosvoorde'


link1 = 'https://belgie-postcodes.be/provincies/antwerpen'
link2 = 'https://belgie-postcodes.be/provincies/henegouwen'
link3 = 'https://belgie-postcodes.be/provincies/limburg'
link4 = 'https://belgie-postcodes.be/provincies/luik'
link5 = 'https://belgie-postcodes.be/provincies/luxemburg'
link6 = 'https://belgie-postcodes.be/provincies/namen'
link7 = 'https://belgie-postcodes.be/provincies/oost-vlaanderen'
link8 = 'https://belgie-postcodes.be/provincies/vlaams-brabant'
link9 = 'https://belgie-postcodes.be/provincies/waals-brabant'
link10 = 'https://belgie-postcodes.be/provincies/west-vlaanderen'

dict1 = dict()

link = link10

list_postal_codes = []
list_city_names = []

driver = webdriver.Chrome(
    executable_path=r'C:\Users\jensb\OneDrive\Documenten\Programming\Python\renewi_distance_calculator_excel\templates\chromedriver.exe')
driver.get(link)

# Get page html
page_soup = BeautifulSoup(driver.page_source, 'html.parser')

containers = page_soup.findAll('td', {'class': 'views-field views-field-field-postcode active'})

for index in range(0, len(containers)):
    list_postal_codes.append(int(containers[index].text[2:]))

containers = page_soup.findAll('td', {'class': 'views-field views-field-field-gemeente'})

for index in range(0, len(containers)):
    list_city_names.append(containers[index].find('a').text)

driver.close()

for index in range(len(list_postal_codes)):
    dict1[list_postal_codes[index]] = list_city_names[index]


print(dict1)

