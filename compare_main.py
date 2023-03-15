"""
NOTE
Trying to create reusable code where only thing needed as input (besides sheets with terms) is 
a file formatted the same as Textbook Chapter Comparisons.xlsx - going to base everything off this file
Initial idea:
  Basically read through the the textbook chapter comparisons sheet and go
  book to book (and chapter by chapter within each book) checking against the key.
  First thing to do would probably be read in first row using iter_rows (returns tuple of values) since those
  are all the book titles and go from there thru the chapters book by book.
  Read each row in that book column, check against mapping to see which chapter from the key it matches with and then check against
  that chapter key
  Output can be represented as a mapping where the key is a term and it's values are the textbooks it appeared in.
  Could probably create enum to represent textbook names and have mapping from ennum to its equivalent textbook name (as a string)
  Probably want output file to be consistent (manually create file load file to edit)
"""

"""
Order of main method:
  1. Create mappings
  2. Instantiate comparisonKey to create key
    - Determine textbook for key by first row of B column (A column has general chapter names for output sheet)
  3. Read in first row from Textbook Chapter Comparisons.xlsx
  4. Go book by book, chapter by chapter to find common terms
"""

from comparison_mappings import comparisonMappings
from comparison_key import comparisonKey

mappings: comparisonMappings = comparisonMappings("../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Textbook Chapter Comparisons.xlsx")
key: comparisonKey = comparisonKey("../OneDrive - The Ohio State University/Survey Development - Dylan/Textbooks Data/Histology A Text and Atlas With Correlated Cell and Molecular Biology, 8e.xlsx")
"""
TODO
Using mappings.chapters_key, go key by key comparing all values (and their terms) to the correspondinf values in
key.key and keep track of which terms appear in which textbooks using: dict[str, list[str]] where the key is the term
and the value is a list of textbooks it occurs in
Once the data structure is created just output it to Comparison Output.xlsx 
"""