from selenium.webdriver.remote.webelement import WebElement
from selenium import webdriver
from selenium.webdriver.common.by import By

# Object to represent a chapter of a textbook
# Big picture: chapter_scraper is within book_scraper

class chapterScraper:
    def __init__(self, title: str, web_driver: webdriver.Chrome ) -> None:
        self.chapter_title = title
        self.driver = web_driver

    def get_headers(self) -> dict[str, WebElement]:
        headers: dict[str, WebElement] = {}
        sections: list[WebElement] = self.driver.find_elements(By.CLASS_NAME, 'scrollTo')
        for section in sections:
            href: str = section.get_attribute('href')
            pound_index: int = href.find('#')
            id: str = f'section_{href[pound_index + 1:]}'
            headers.update({id, section})
        return headers

    def get_bold_terms(self, headers: dict[str, WebElement], path: str) -> list[list[str]]:
        keywords: list[list[str]] = []
        for id in headers.keys():
            bold_web_elements: list[WebElement] = self.driver.find_elements(By.XPATH, path)
            bold_words: list[str] = []
            for element in bold_web_elements:
                if len(element.text) > 1:
                    bold_words.append(element.text)
            keywords.append(bold_words)
        return keywords
