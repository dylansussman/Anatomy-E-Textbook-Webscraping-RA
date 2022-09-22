from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Object to represent a chapter of a textbook
# Big picture: chapter_scraper is within book_scraper
class chapterScraper:
    BAD_HEADERS: list[str] = ['Summary', 'Review Questions', 'Additional Histologic Images', 'Outline']
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
            word = element.text
            if ':' in word:
                word = word[:word.find(':')]
            if len(word) > 1 and self.add_with_number(word) and not '.' in word and not self.is_roman_numeral(word) and word != "Image:":
                if not word.lower() in bold_words:
                    bold_words.append(word.lower())
        return bold_words


    def create_worksheet(self, ws: Worksheet, data: dict[str, list[str]]) -> None:
        # Add chapter name to first row of sheet and merge cells
        ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=len(data))
        ws.cell(row=1, column=1, value=self.chapter_title).font = Font(bold=True, size=14)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal='center')
        ws.cell(row=1, column=1).fill = PatternFill('solid', fgColor='BFBFBF')
        row: int = 2
        col: int = 1
        max_width: int = 0
        for header, terms in data.items():
            ws.cell(row=row, column=col, value=header).font = Font(bold=True)
            ws.cell(row=row, column=col).alignment = Alignment(horizontal='center')
            max_width = len(header)
            for term in terms:
                row += 1
                ws.cell(row=row, column=col, value=term)
                if len(term) > max_width:
                    max_width = len(term)
            ws.column_dimensions[get_column_letter(col)].width = max_width
            row = 2
            col += 1

    def is_roman_numeral(self, word: str) -> bool:
        word_chars_bool: list[bool] = []
        for char in word:
            word_chars_bool.append(char in self.ROMAN_NUMERALS)
        return all(word_chars_bool)

    def add_with_number(self, word: str) -> bool:
        words_chars_bool: list[bool] = []
        for char in word:
            words_chars_bool.append(char.isnumeric())
        if any(words_chars_bool):
            if len(words_chars_bool) < 3:
                return False
            else:
                return True
        return True

    # NOTE Not using this function, but could need later
    # def add_with_parentheses(self, word: str) -> bool:
    #     if word.startswith('(') and word.endswith(')'):
    #         word_wo_paren: str = word[1:len(word) - 1]
    #         return self.add_with_number(word_wo_paren)
    #     return True
            

