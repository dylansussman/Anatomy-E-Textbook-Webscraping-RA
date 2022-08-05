import csv
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from itertools import zip_longest

DRIVER_PATH = '/usr/local/bin/chromedriver'
driver = webdriver.Chrome(executable_path=DRIVER_PATH)
# Link to go directly to e-texbook page
driver.get('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/content.aspx?sectionid=168703105&bookid=2212#168703125')

# Get thru login page
username: WebElement = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
password: WebElement = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'password')))
username.send_keys('sussman.47')
password.send_keys('Banbury3117&')
driver.find_element(By.ID, 'submit').click()

# Get down the page (scroll past chapter outline) to the beginning of text
intro_link: WebElement = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Introduction")))

# Get headers for each chapter section
headers: list[WebElement] = driver.find_elements(By.CLASS_NAME, 'scrollTo')
section_ids: list[str] = []
for element in headers:
    if element.accessible_name != 'Chapter Outline':
        href: str = element.get_attribute('href')
        pound_index: int = href.find('#')
        id: str = f'section_{href[pound_index + 1:]}'
        section_ids.append(id)

# Creart 2D list with all bolded terms
keywords: list[list[str]] = []
for id in section_ids:
    path: str = f"//div[@id='{id}']/descendant::div[not(@class='caption-legend')]/div[@class='para']/strong | //div[@id='{id}']/descendant::div[not(@class='caption-legend')]/ul[@class='bullet']/li/div[@class='para']/strong"
    bold_web_elements: list[WebElement] = driver.find_elements(By.XPATH, path)
    bold_words: list[str] = []
    for element in bold_web_elements:
        bold_words.append(element.text)
    keywords.append(bold_words)

# Create list of headers (just text)
headers_text: list[str] = []
for header in headers:
    if header.accessible_name != 'Chapter Outline':
        headers_text.append(header.accessible_name)

# Transpose keywords in order to write it to .csv file
transposed_tuples_list: list[list[str]] = list(zip_longest(*keywords, fillvalue=''))
keywords_transposed = [list(sublist) for sublist in transposed_tuples_list]

# Write out terms to .csv file
with open('Chapter_5_Abdomen_Terms.csv', 'w') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(headers_text)
    writer.writerows(keywords_transposed)
    





