from book_scraper import bookScraper
from chapter_scraper import chapterScraper

scraper: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=2992')
scraper.login()
data = scraper.get_book_data()
scraper.create_workbook(data)