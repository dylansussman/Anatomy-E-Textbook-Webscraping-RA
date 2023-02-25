"""
NOTE
Initial idea:
  Basically read through the the textbook chapter comparisons sheet and go
  book to book (and chapter by chapter within each book) checking against the key.
  First thing to do would probably be read in first row (returns tuple of values) since those
  are all the book titles and go from there thru the chapters book by book.
"""

"""
TODO
Create two mappings from Textbook Chapter Comparisons Excel File:
  - General "chapter" names -> corresponding chapters from key
  - Key chapter names -> corresponding chapters from all other textbooks
"""

# TODO Read in Histology A Text and Atlas With Correlated Cell and Molecular Biology, 8e as word bank

"""
Order of main method:
  1. Create mappings
  2. Instantiate comparisonKey to create key
    - Determine textbook for key by first row of B column (A column has general chapter names for output sheet)
  3. Read in first row from Textbook Chapter Comparisons.xlsx
  4. Go book by book, chapter by chapter to find common terms
"""