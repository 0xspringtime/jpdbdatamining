import requests
from bs4 import BeautifulSoup
import openpyxl

# Fetch the webpage
base_url = 'https://jpdb.io/visual-novel-difficulty-list?offset='
pages = ['0#a', '50#a', '100#a', '150#a', '200#a','250#a', '300#a',
         '350#a', '400#a','450#a', '500#a','550#a', '600#a',
         '650#a', '700#a','750#a', '800#a', '850#a', '900#a', '950#a', '1000#a',
         '1050#a', '1100#a','1150#a', '1200#a','1250#a', '1300#a',
         '1350#a']


# Extract the title and length values for each entry
titles = []
td_values = []
for page in pages:
    url = base_url + page
    response = requests.get(url)
    # Parse the HTML content
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find all the entries on the webpage
    entries = soup.find_all('div', style='display: flex; flex-wrap: wrap;')

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet.append(['Title', 'Length (in words)', 'Unique words', 'Unique words (used once)', 'Unique words (used once %)', 'Unique kanji', 'Unique kanji (used once)', 'Unique kanji readings', 'Difficulty', 'MAL avg. rating', 'MAL rating count'])





    for entry in entries:
      title = entry.find('h5').text
      tds = entry.find_all('td')
      td_values_for_entry = [td.text for td in tds]
      td_values_for_entry.insert(0, title)
      titles.append(td_values_for_entry)

    print(f'Titles: {titles}')

    for t in titles:
      sheet.append(t)

    workbook.save('vns.xlsx')