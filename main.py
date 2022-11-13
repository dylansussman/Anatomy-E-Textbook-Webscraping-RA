from book_scraper import bookScraper
import json

'''NOTE LWW Textbooks'''

'''Scraping & Writing Data for Atlas of Histology with Functional Correlations, 13e'''
# scraper1: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=2992')
# scraper1.login()
# path1 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/div[@class='para']/strong"
# path2 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/ul[@class='bullet']/li/div[@class='para']/strong"
# path3 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ul[@class='bullet']/li/div[@class='para']/strong"
# path4 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ul[@class='bullet']/li/strong"
# data1 = scraper1.get_book_data('Chapter 1: Histologic Methods', [path1, path2, path3, path4])
# with open('Textbook_Data/atlas_of_histology_with_function_correlations_13e_data.txt', 'w') as file1:
#     file1.write(json.dumps(data1))
# # with open('Textbook_Data/atlas_of_histology_with_function_correlations_13e_data.txt', 'r') as file1:
# #     data1 = file1.read()
# # data1 = json.loads(data1)
# scraper1.create_workbook(data1, '../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Atlas of Histology with Functional Correlations, 13e.xlsx')

'''Scraping & Writing Data for Gartner and Hiatt's Atlas and Text of Histology, 8e'''
# scraper2: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=3216')
# scraper2.login()
# path1 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/div[@class='para']/strong"
# path2 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/ul[@class='bullet']/li/div[@class='para']/strong"
# path3 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ul[@class='bullet']/li/div[@class='para']/strong"
# path4 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ul[@class='bullet']/li/strong"
# data2 = scraper2.get_book_data('CHAPTER 1: INTRODUCTION TO HISTOLOGIC TECHNIQUES', [path1, path2, path3, path4])
# with open("Textbook_Data/gartner_and_hiatt's_atlas_and_text_of_histology_8e_data.txt", 'w') as file2:
#     file2.write(json.dumps(data2))
# # with open("Textbook_Data/gartner_and_hiatt's_atlas_and_text_of_histology_8e_data.txt", 'r') as file2:
# #     data2 = file2.read()
# # data2 = json.loads(data2)
# scraper2.create_workbook(data2, "../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Gartner and Hiatt's Atlas and Text of Histology, 8e.xlsx")

'''Scraping & Writing Data for Histology: A Text and Atlas: With Correlated Cell and Molecular Biology, 8e'''
# scraper3: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=2583')
# scraper3.login()
# path1 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/div[@class='para']/strong"
# path2 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/ul[@class='bullet']/li/div[@class='para']/strong"
# data3 = scraper3.get_book_data('1: Methods', [path1, path2])
# with open('Textbook_Data/histology_a_text_and_atlas_with_correlated_cell_and_molecular_biology_8e_data.txt', 'w') as file3:
#     file3.write(json.dumps(data3))
# # with open('Textbook_Data/histology_a_text_and_atlas_with_correlated_cell_and_molecular_biology_8e_data.txt', 'r') as file3:
# #     data3 = file3.read()
# # data3 = json.loads(data3)
# scraper3.create_workbook(data3, '../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Histology A Text and Atlas With Correlated Cell and Molecular Biology, 8e.xlsx')

'''Scraping & Writing Data for Histology From a Clinical Perspective, 2e'''
# scraper4: bookScraper = bookScraper('https://meded-lwwhealthlibrary-com.proxy.lib.ohio-state.edu/book.aspx?bookid=3201')
# scraper4.login()
# path1 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/div[@class='para']/strong"
# path2 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/ul[@class='bullet']/li/div[@class='para']/strong"
# path3 = f"/descendant::div[not(@class='caption-legend' or @class='boxed-content')]/descendant::ol[@class='roman-upper' or @class='alpha-upper' or @class='number']/li/div[@class='para']/strong"
# data4 = scraper4.get_book_data('1: Illustrated Glossary of Histologic and Pathologic Terms', [path1, path2, path3])
# with open('Textbook_Data/histology_from_a_clinical_perspective_2e.txt', 'w') as file4:
#     file4.write(json.dumps(data4))
# # with open('Textbook_Data/histology_from_a_clinical_perspective_2e.txt', 'r') as file4:
# #     data4 = file4.read()
# # data4 = json.loads(data4)
# scraper4.create_workbook(data4, '../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Histology from a Clinical Perspective, 2e.xlsx')

'''NOTE Elsevier Textbooks'''

'''Scraping & Writing Data for Wheater's Functional Histology, Sixth Edition'''
# TODO Grab words from beginning of chapter 20 that aren't in a section; put them in own column w/o header in spreadsheet
# TODO Order words on the sheet in descending order of term length
    # Check if there are any other chapters like this
scraper5: bookScraper = bookScraper("https://www-clinicalkey-com.proxy.lib.ohio-state.edu/#!/browse/book/3-s2.0-C20090600258")
scraper5.login()
scraper5.get_to_elsevier_book("Wheater's Functional Histology")
path1 = f"/p/b/i"
path2 = f"/p/i/b"
path3 = f"/ul/li/p/b"
path4 = f"/ul/li/p/i"
path5 = f"//a[section]/following-sibling::div[1]/div/p/b"
path6 = f"//a[section]/following-sibling::div[1]/div/p/i"
path7 = f"//a[section]/parent::p/following-sibling::div[1]/div/p/b"
path8 = f"//a[section]/parent::p/following-sibling::div[1]/div/p/i"
data5 = scraper5.get_book_data("Cell structure and function", [path1, path2, path3, path4, path5, path6, path7, path8])
with open("Textbook_Data/wheater's_functional_histology_sixth_edition.txt", 'w') as file5:
    file5.write(json.dumps(data5))
# with open("Textbook_Data/wheater's_functional_histology_sixth_edition.txt", 'r') as file5:
#     data5 = file5.read()
# data5 = json.loads(data5)
scraper5.create_workbook(data5, "../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Wheater's Functional Histology, Sixth Edition.xlsx")

'''Scraping & Writing Data Stevens & Lowe's Human Histology, Fifth Edition'''
# scraper6: bookScraper = bookScraper("https://www-clinicalkey-com.proxy.lib.ohio-state.edu/#!/browse/book/3-s2.0-C20170016105")
# scraper6.login()
# scraper6.get_to_elsevier_book("Stevens & Lowe's Human Histology")
# data6 = scraper6.get_book_data("Histology")
# with open("Textbook_Data/stevens_&_lowe's_human_histology_fifth_edition.txt", "w") as file6:
#     file6.write(json.dumps(data6))
# # with open("Textbook_Data/stevens_&_lowe's_human_histology_fifth_edition.txt", "r") as file6:
# #     data6 = file6.read()
# # data6 = json.loads(data6)
# scraper6.create_workbook(data6, "../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Stevens & Lowe's Human Histology, Fifth Edition.xlsx")

