from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from striprtf.striprtf import rtf_to_text


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

    def get_book_data(self) -> None:
        chapter_1: WebElement = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, 'Chapter 1: Histologic Methods')))
        chapter_list: list[WebElement] = self.driver.find_elements(By.CLASS_NAME, 'tocLink_wrap')
        chapter_1.click()
        for chapter in chapter_list:
            # create a chapter_scraper object and call the necessary methods to 
            # get the needed info from each chapter and write out to Excel file
            # as I go (i.e., in this loop call some method to make Excel file one chapter at a time)
            # Try find a way to make excel methods associated with both scraper objects (i.e., create/edit
            # workbook in book_scraper and create/edit sheets in chapter_scraper)
            pass
