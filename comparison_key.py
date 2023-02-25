from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

"""
This class represents the bank of words that will be tested against.
An instantiation of this class will contain a dictionary with every term from the textbook designated
as the key to compare terms from other textbooks against. These terms will be orgnanized by chapter
and this class contains methods to compare terms from other textbooks against the key
"""
class comparisonKey:
  """
  file_name is the name of the sheet designated as the key
  """
  def __init__(self, file_name: str) -> None:
    # Class variable to hold all terms from sheet designated as the comparison key
    self.key = self.__intialize_key(file_name)
    pass

  def __intialize_key(self, file_name: str) -> dict[str, list[str]]:
    key: dict[str, list[str]] = {}
    wb: Workbook = load_workbook(file_name)
    for sheet_name in wb.sheetnames:
      sheet: Worksheet = wb[sheet_name]
      # TODO Finish loading entire sheet into key

    