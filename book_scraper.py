from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from striprtf.striprtf import rtf_to_text
from chapter_scraper import chapterScraper
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

class bookScraper:
    # DRIVER_PATH: str = '/usr/local/bin/chromedriver'

    def __init__(self, book_url: str) -> None:
        # self.driver = webdriver.Chrome(executable_path=self.DRIVER_PATH)
        self.driver = webdriver.Chrome(ChromeDriverManager().install())
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
    def get_book_data(self, first_chapter: str, ) -> dict[str, dict[str, list[str]]]:
        chapter_1: WebElement = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.LINK_TEXT, first_chapter)))
        chapter_list: list[str] = []
        for element in self.driver.find_elements(By.CLASS_NAME, 'tocLink_wrap'):
            if not 'Appendix' in element.text:
                chapter_list.append(element.text)
        chapter_1.click()
        chapter_dict: dict[str, dict[str, list[str]]] = {}
        for chapter_title in chapter_list:
            chapter: chapterScraper = chapterScraper(chapter_title, self.driver)
            headers: dict[str, WebElement] = chapter.get_headers()
            header_term_dict: dict[str, list[str]] = {}
            path: str = ''
            for section_id in headers.keys():
                path1 = f"//div[@id='{section_id}']/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/div[@class='para']/strong"
                path2 = f"//div[@id='{section_id}']/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/ul[@class='bullet']/li/div[@class='para']/strong"
                # path3 = f"//div[@id='{section_id}']/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ul[@class='bullet']/li/div[@class='para']/strong"
                # path4 = f"//div[@id='{section_id}']/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ul[@class='bullet']/li/strong"
                # path5 = f"//div[@id='{section_id}']/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ol[@class='roman-upper' or @class='alpha-upper' or @class='number']/li/div[@class='para']/strong"
                # path6 = f"//div[@id='{section_id}']/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/ul[@class='bullet']/li/strong"
                # path7 = f"//div[@id='{section_id}']/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/ul[@class='bullet']/li/ul[@class='bullet']/li/div[@class='para']/strong"
                path = f"{path1} | {path2}"
                bold_terms: list[str] = chapter.get_section_bold_terms(path)
                if header_term_dict.get(headers.get(section_id).text) == None or len(bold_terms) > 0:
                    header_term_dict.update({headers.get(section_id).text:bold_terms})
            chapter_dict.update({chapter_title:header_term_dict})
            # Last chapter doesn't have a next chapter button
            try:
                next_chapter: WebElement = self.driver.find_element(By.XPATH, '//span[text()="Next Chapter"]')
            except NoSuchElementException:
                pass
            else:
                next_chapter.click()
        return chapter_dict

    def create_workbook(self, data: dict[str, dict[str, list[str]]], workbook_name: str) -> None:
        wb: Workbook = Workbook()
        wb.remove(wb.active)
        for chapter_title, chapter_data in data.items():
            title: str = chapter_title[:chapter_title.find(':')]
            ws: Worksheet = wb.create_sheet(title)
            chapter: chapterScraper = chapterScraper(chapter_title, self.driver)
            chapter.create_worksheet(ws, chapter_data)
        wb.save(workbook_name)
        wb.close()


