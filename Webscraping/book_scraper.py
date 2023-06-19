from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from striprtf.striprtf import rtf_to_text
from Webscraping.chapter_scraper import chapterScraper
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

    # NOTE Only needed for Elsevier textbooks
    # Gets from main webpage to page for specified book
    def get_to_elsevier_book(self, book_title: str):
        search_bar: WebElement = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'search-bar')))
        search_bar.send_keys(book_title)
        search_bar.click()
        clickable_book: WebElement = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'span.suggestion__search-term')))
        while clickable_book.text != book_title:
            clickable_book = self.driver.find_element(By.CSS_SELECTOR, 'span.suggestion__search-term')
        clickable_book.click()
        # self.driver.find_element(By.CSS_SELECTOR, 'span.suggestion__search-term').click()

    # Could return dictionary with key as chapter title and values as another dictionary of
    # Each header as a key with the bold list from that subsection as the value 
    def get_book_data(self, first_chapter: str, path_list: list[str]) -> dict[str, dict[str, list[str]]]:
        # NOTE For scraping LWW Textbooks
        # chapter_1: WebElement = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.LINK_TEXT, first_chapter)))

        # NOTE For scraping Elsevier Textbooks
        elements: list[WebElement] = WebDriverWait(self.driver, 30).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, 'a.has-chapter')))

        chapter_list: list[str] = []
        # NOTE For scraping LWW Textbooks
        # for element in self.driver.find_elements(By.CLASS_NAME, 'tocLink_wrap'):
        #     if not 'Appendix' in element.text:
        #         chapter_list.append(element.text)
        
        # NOTE For scraping Elsevier Textbooks
        # elements = self.driver.find_elements(By.CSS_SELECTOR, 'a.has-chapter')
        chapter_1: WebElement
        for element in elements:
            if not "Case" in element.accessible_name and not "Review" in element.accessible_name:
                chapter_list.append(element.accessible_name)
                # NOTE For scraping Elsevier Textbooks
                if first_chapter in element.accessible_name:
                    chapter_1 = element

        chapter_1.click()
        chapter_dict: dict[str, dict[str, list[str]]] = {}
        for chapter_title in chapter_list:
            chapter: chapterScraper = chapterScraper(chapter_title, self.driver)            
            headers: dict[str, WebElement] = chapter.get_headers()
            header_term_dict: dict[str, list[str]] = {}
            bold_terms: list[str] = []
            # NOTE Get bold terms at the beginning of chapter that aren't in a section; occur before any of the chapter's sections
            # NOTE Only needed for Textbook of Histology
            bold_terms = chapter.get_section_bold_terms(f"//div[@class='s-content ng-scope early-item']/div/p/b")

            if len(bold_terms) > 0: header_term_dict.update({"Chapter Introduction":bold_terms})
            path: str = ''
            for section_id in headers.keys():
                path = ""
                for i, p in enumerate(path_list):
                    if i > 0:
                        path += " | "
                    if "/a[section]" in p:
                        path += f"//a[@id='{section_id}']{p[p.find(']')+1:]}"
                    else:
                        path += f"//section[@id='{section_id}']{p}"
                # NOTE Next three lines only matter for Textbook of Histology (first condition in if statement checks what textbook)
                # Because of HTML structure, only outermost sections will be checked for bold terms, so the outputted sheets will have
                # less columns
                parent_tag: str = self.driver.find_element(By.ID, section_id).find_element(By.XPATH, "./..").tag_name
                if first_chapter == "Introduction to Histology and Basic Histological Techniques" and parent_tag == "section":
                    continue

                bold_terms = chapter.get_section_bold_terms(path)
                section_title: str = headers.get(section_id).accessible_name
                # NOTE Commented out for Stevens & Lowe's only, uncomment for next textbook
                # if header_term_dict.get(section_title) == None and len(bold_terms) > 0:
                if len(bold_terms) > 0:
                    if header_term_dict.get(section_title) != None:
                        header_term_dict.update({section_title:header_term_dict.get(section_title) + bold_terms})
                    else:
                        header_term_dict.update({section_title:bold_terms})
                # NOTE For scraping Elsevier Textbooks
                # chapter.get_figure_terms(section_id, header_term_dict)

            chapter_dict.update({chapter_title:header_term_dict})
            
            # NOTE For scraping Elsevier Textbooks
            # Checking for pop-up and exiting out of it if it's present
            try:
                pop_up: WebElement = self.driver.find_element(By.ID, 'acsFocusFirst')
            except NoSuchElementException:
                pass
            else:
                pop_up.click()

            # Last chapter doesn't have a next chapter button
            try:
                # NOTE For scraping Elsevier Textbooks
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
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
            # NOTE For scraping LWW Textbooks
            # title: str = chapter_title[:chapter_title.find(':')]

            # NOTE for scraping Elsevier Textbooks
            title: str = chapter_title[:chapter_title.find('.')]
            
            ws: Worksheet = wb.create_sheet(title)
            chapter: chapterScraper = chapterScraper(chapter_title, self.driver)
            chapter.create_worksheet(ws, chapter_data)
        wb.save(workbook_name)
        wb.close()


