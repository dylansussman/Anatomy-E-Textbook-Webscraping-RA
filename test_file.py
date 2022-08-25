from book_scraper import bookScraper

scraper: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=2992')
scraper.login()
scraper.get_book_data()