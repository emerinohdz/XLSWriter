# coding: utf-8

from xlwt import Workbook
from xlrd import open_workbook
from xlutils.copy import copy

class XLSWriter:

    DEFAULT_SHEET_NAME = "Hoja 1"

    def __init__(self, input_filename=None):
        self.__input_filename = input_filename
        self.__rd_book = None

        if not input_filename:
            self.__wt_book = Workbook()
            self.__wt_book.add_sheet(XLSWriter.DEFAULT_SHEET_NAME)
        else:
            self.__wt_book = copy(open_workbook(input_filename, formatting_info=True))
        
        self.active_sheet = self.__wt_book.get_active_sheet()

    def __get_row_index(self, col_index, col_value):
        """
            Get the row index where col_value is found at the given column index
        """
        if not self.__input_filename:
            raise Exception("No filename given, i.e. no data to search for")

        if not self.__rd_book:
            self.__rd_book = open_workbook(self.__input_filename, on_demand=True)

        row_index = None
        sheet = self.__rd_book.get_sheet(self.active_sheet)

        for i in range(sheet.nrows):
            if sheet.cell_value(i, col_index) == col_value:
                row = i
                break

        if not row_index:
            raise Exception("Row to replace not found!")

        return row_index

    def append(self, cols):
        """
            Append a new row containing cols data at the end of the document.
        """
        row_count = len(self.__wt_book.get_sheet(self.active_sheet).get_rows())
        self.write(row_count, cols)

    def replace(self, col_index, col_value, cols):
        """
            Replace the contents of the row where col_value is found at the given
            column index.
        """
        self.write(self.__get_row_index(col_index, col_value), cols)
        
    def write(self, row_index, cols):
        """
            Write the data contained in cols in the row with index row_index
            and replace the data found there.
        """
        sheet = self.__wt_book.get_sheet(self.active_sheet)

        for i in range(len(cols)):
            sheet.write(row_index, i, cols[i])

    def save(self, output_filename=None):
        """
            Serialize the workbook to the given output_filename or
            to input_filename if not given
        """
        if not output_filename:
            self.__wt_book.save(self.__input_filename)
        else:
            self.__wt_book.save(output_filename)

