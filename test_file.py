from book_scraper import bookScraper
from chapter_scraper import chapterScraper
import json

scraper: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=2992')
scraper.login()
data = scraper.get_book_data()
with open('atlas_of_histology_with_function_correlations_13e_data.txt', 'w') as file:
    file.write(json.dumps(data))
# with open('atlas_of_histology_with_function_correlations_13e_data.txt', 'r') as file:
#     data = file.read()
# data = json.loads(data)
scraper.create_workbook(data)