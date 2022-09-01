from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font

# Object to represent a chapter of a textbook
# Big picture: chapter_scraper is within book_scraper

class chapterScraper:
    BAD_HEADERS: list[str] = ['Summary', 'Review Questions', 'Additional Histologic Images']
    ROMAN_NUMERALS: list[str] = ['I', 'V', 'X', 'L', 'C', 'D']
    
    def __init__(self, title: str, web_driver: webdriver.Chrome ) -> None:
        self.chapter_title = title
        self.driver = web_driver

    def get_headers(self) -> dict[str, WebElement]:
        headers: dict[str, WebElement] = {}
        sections: list[WebElement] = self.driver.find_elements(By.CLASS_NAME, 'scrollTo')
        for section in sections:
            if not any(map(lambda header: header in section.text, self.BAD_HEADERS)):
                href: str = section.get_attribute('href')
                pound_index: int = href.find('#')
                id: str = f'section_{href[pound_index + 1:]}'
                headers.update({id:section})
        return headers

    def get_section_bold_terms(self, path: str) -> list[str]:
        bold_web_elements: list[WebElement] = self.driver.find_elements(By.XPATH, path)
        bold_words: list[str] = []
        for element in bold_web_elements:
            if len(element.text) > 1 and not any(char.isnumeric() for char in element.text) and not '.' in element.text and not self.is_roman_numeral(element.text):
                    bold_words.append(element.text)
        return bold_words


    # TODO Formatting sheet: autowidth columns
    # TODO Add chapter name to each sheet
    def create_worksheet(self, ws: Worksheet, data: dict[str, list[str]]) -> None:
        row: int = 1
        col: int = 1
        for header, terms in data.items():
            ws.cell(row=row, column=col, value=header)
            for term in terms:
                row += 1
                ws.cell(row=row, column=col, value=term)
            row = 1
            col += 1
        for cell in ws[1]:
            cell.font = Font(bold=True)

    def is_roman_numeral(self, word: str) -> bool:
        word_chars_bool: list[bool] = []
        for char in word:
            word_chars_bool.append(char in self.ROMAN_NUMERALS)
        return all(word_chars_bool)
