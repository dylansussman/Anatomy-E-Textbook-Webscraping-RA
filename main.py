from book_scraper import bookScraper
import json

# TODO Pass paths into get_book_data and bring all pathes out of book_scraper and into the main file
    # In order to do this will have to pass in array of path of part after section_id; will have to add
    # that in inside the book scraper
# TODO Then, continue on to next series of textbooks

'''NOTE LWW Textbooks'''

'''Scraping & Writing Data for Atlas of Histology with Functional Correlations, 13e'''
'''NOTE Pathes used: path1-4'''
# scraper1: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=2992')
# scraper1.login()
# data1 = scraper1.get_book_data('Chapter 1: Histologic Methods')
# with open('Textbook_Data/atlas_of_histology_with_function_correlations_13e_data.txt', 'w') as file1:
#     file1.write(json.dumps(data1))
# # with open('Textbook_Data/atlas_of_histology_with_function_correlations_13e_data.txt', 'r') as file1:
# #     data1 = file1.read()
# # data1 = json.loads(data1)
# scraper1.create_workbook(data1, '../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Atlas of Histology with Functional Correlations, 13e.xlsx')

'''Scraping & Writing Data for Gartner and Hiatt's Atlas and Text of Histology, 8e'''
'''NOTE Pathes used: path1-4'''
# scraper2: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=3216')
# scraper2.login()
# data2 = scraper2.get_book_data('CHAPTER 1: INTRODUCTION TO HISTOLOGIC TECHNIQUES')
# with open("Textbook_Data/gartner_and_hiatt's_atlas_and_text_of_histology_8e_data.txt", 'w') as file2:
#     file2.write(json.dumps(data2))
# # with open("Textbook_Data/gartner_and_hiatt's_atlas_and_text_of_histology_8e_data.txt", 'r') as file2:
# #     data2 = file2.read()
# # data2 = json.loads(data2)
# scraper2.create_workbook(data2, "../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Gartner and Hiatt's Atlas and Text of Histology, 8e.xlsx")

'''Scraping & Writing Data for Histology: A Text and Atlas: With Correlated Cell and Molecular Biology, 8e'''
'''NOTE Pathes used: path1-2'''
# scraper3: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=2583')
# scraper3.login()
# data3 = scraper3.get_book_data('1: Methods')
# with open('Textbook_Data/histology_a_text_and_atlas_with_correlated_cell_and_molecular_biology_8e_data.txt', 'w') as file3:
#     file3.write(json.dumps(data3))
# # with open('Textbook_Data/histology_a_text_and_atlas_with_correlated_cell_and_molecular_biology_8e_data.txt', 'r') as file3:
# #     data3 = file3.read()
# # data3 = json.loads(data3)
# scraper3.create_workbook(data3, '../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Histology A Text and Atlas With Correlated Cell and Molecular Biology, 8e.xlsx')

'''Scraping & Writing Data for Histology From a Clinical Perspective, 2e'''
'''NOTE Pathes used: path1-2, path5'''
# scraper4: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=3201')
# scraper4.login()
# data4 = scraper4.get_book_data('1: Illustrated Glossary of Histologic and Pathologic Terms')
# with open('Textbook_Data/histology_from_a_clinical_perspective_2e.txt', 'w') as file4:
#     file4.write(json.dumps(data4))
# # with open('Textbook_Data/histology_from_a_clinical_perspective_2e.txt', 'r') as file4:
# #     data4 = file4.read()
# # data4 = json.loads(data4)
# scraper4.create_workbook(data4, '../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Histology from a Clinical Perspective, 2e.xlsx')

'''NOTE Elsevier Textbooks'''

'''Scraping & Writing Data for Wheater's Functional Histology, Sixth Edition'''
scraper5: bookScraper = bookScraper("https://www-clinicalkey-com.proxy.lib.ohio-state.edu/#!/browse/book/3-s2.0-C20090600258")
scraper5.login()
scraper5.get_to_elsevier_book("Wheater's Functional Histology")
data5 = scraper5.get_book_data("Cell structure and function")
with open("Textbook_Data/wheater's_functional_histology_sixth_edition.txt", 'w') as file5:
    file5.write(json.dumps(data5))
# with open("Textbook_Data/wheater's_functional_histology_sixth_edition.txt", 'r') as file5:
#     data5 = file5.read()
# data5 = json.loads(data5)
scraper5.create_workbook(data5, "../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Wheater's Functional Histology, Sixth Edition.xlsx")