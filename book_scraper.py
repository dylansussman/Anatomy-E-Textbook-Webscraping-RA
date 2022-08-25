from lib2to3.pgen2 import driver
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from striprtf.striprtf import rtf_to_text
from chapter_scraper import chapterScraper

class bookScraper:
    DRIVER_PATH: str = '/usr/local/bin/chromedriver'

    def __init__(self, book_url: str) -> None:
        self.driver = webdriver.Chrome(executable_path=self.DRIVER_PATH)
        self.driver.get(book_url)

    # Get thru the login page to get to the specified textbook
    # First the credentials are read in. Then the username and password are
    # entered allowing access to the page
    def login(self) -> None:
        username, password = '', ''
        with open('../Documents/Private/login_credentials.rtf', 'r') as credentials:
            while username == '':
                username = rtf_to_text(credentials.readline())
            username = username[:username.find('/')]
            password = rtf_to_text(credentials.readline())
        
        username_element: WebElement = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
        password_element: WebElement = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'password')))
        username_element.send_keys(username)
        password_element.send_keys(password)
        self.driver.find_element(By.ID, 'submit').click()

    # Could return dictionary with key as chapter title and values as another dictionary of
    # Each header as a key with the bold list from that subsection as the value 
    def get_book_data(self) -> dict[str, dict[str, list[str]]]:
        chapter_1: WebElement = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.LINK_TEXT, 'Chapter 1: Histologic Methods')))
        # chapter_list: list[WebElement] = self.driver.find_elements(By.CLASS_NAME, 'tocLink_wrap')
        chapter_list: list[str] = []
        for element in self.driver.find_elements(By.CLASS_NAME, 'tocLink_wrap'):
            chapter_list.append(element.text)
        chapter_1.click()
        chapter_dict: dict[str, dict[str, list[str]]] = {}
        for chapter_title in chapter_list:
            # create a chapter_scraper object and call the necessary methods to 
            # get the needed info from each chapter and write out to Excel file
            # as I go (i.e., in this loop call some method to make Excel file one chapter at a time)
            # Try find a way to make excel methods associated with both scraper objects (i.e., create/edit
            # workbook in book_scraper and create/edit sheets in chapter_scraper)

            chapter: chapterScraper = chapterScraper(chapter_title, self.driver)
            headers: dict[str, WebElement] = chapter.get_headers()
            header_term_dict: dict[str, list[str]] = {}
            path: str = ''
            for section_id in headers.keys():
                path = f"//div[@id='{section_id}']/descendant::div[not(@class='boxed-content')]/div[@class='para']/strong"
                bold_terms: list[str] = chapter.get_section_bold_terms(path)
                header_term_dict.update({headers.get(section_id).text:bold_terms})
            chapter_dict.update({chapter_title:header_term_dict})
            next_chapter: WebElement = self.driver.find_element(By.XPATH, '//span[text()="Next Chapter"]')
            next_chapter.click()
