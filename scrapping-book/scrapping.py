from selenium import webdriver
from bs4 import BeautifulSoup
import time
import os

# Path to the Edge browser driver
edge_driver_path = r"msedgedriver.exe"

# Initialize the Edge driver
driver = webdriver.Edge()

# URL of the first page
first_page = 'https://www.kostenlosonlinelesen.net/kostenlose-dune-01-der-wuestenplanet'
driver.get(first_page)

# Wait for the page to fully load if necessary
time.sleep(5)

# Get the HTML of the page
html = driver.page_source
# Create a BeautifulSoup object with the obtained HTML
soup = BeautifulSoup(html, 'html.parser')

# Use BeautifulSoup to search and extract data from the HTML
divs = soup.find_all('div', class_='txt-cnt')

# Create a text string to store the result
texto_completo = ""

# Iterate over the found <div> elements and extract their text
for div in divs:
    # Get the text from the <div> element with line breaks
    texto = str(div).replace("<br/>", "\n")
    # Remove the opening and closing <div> tags
    texto = texto[21:-6].strip() + " "
    # Add the text to the complete result
    texto_completo += texto

# Process the rest of the pages
for i in range(2, 329):
    page = f"https://www.kostenlosonlinelesen.net/kostenlose-dune-01-der-wuestenplanet/lesen/{i}"
    driver.get(page)
    time.sleep(2)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    divs = soup.find_all('div', class_='txt-cnt')
    for div in divs:
        texto = str(div).replace("<br/>", "\n")
        texto = texto[21:-6].strip() + " "
        texto_completo += texto

# Save the result to a text file
with open("resultado.txt", "w", encoding="utf-8") as file:
    file.write(texto_completo)

driver.quit()
