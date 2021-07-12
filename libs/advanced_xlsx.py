from bs4 import BeautifulSoup, element
from openpyxl import Workbook, worksheet
import xlrd
from openpyxl.utils.cell import column_index_from_string

class AdvancedXlsx:

    def __init__(self, wb=Workbook(), sheet=None):
        self.wb = wb
        if sheet is None:
           sheet = self.wb.active

        self.sheet = sheet
        

    def open_xls(self, path: str)->Workbook:
        with open(path, "r", encoding="latin-1") as f:
            soup = BeautifulSoup(f.read(), 'html') 
        
        if soup.find_all("table"):
            self.get_from_html(soup)
            return self.wb

        self.convert_xls(path)
        return self.wb
        
    def convert_xls(self, path:str)->None:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)
        for i in range(sheet.nrows):
            row = [sheet.cell_value(rowx=i, colx=j) for j in range(sheet.ncols)]
            self.sheet.append(row)

    def get_from_html(self, soup: BeautifulSoup)->None:
        tables = tables = soup.findChildren("table", recursive=False)
        for table in tables:
            self.get_table(table)
                     
    def get_table(self, table: element.Tag)->None:
        count = 1
        for row in table.findChildren("tr", recursive=False):
            for j, col in enumerate(row.findChildren("td", recursive=False), start=1):
                if col.table:
                    i = count
                    for ii, sub_row in enumerate(col.table.findChildren("tr", recursive=False)):
                        for jj, sub_col in enumerate(sub_row.findChildren("td", recursive=False), start=1):
                            self.sheet.cell(row=i+ii, column=j).value = sub_col.text.strip()
                else:
                    self.sheet.cell(row=count, column=j).value = col.text.strip()
            count = self.sheet.max_row + 1

    def change_sheet(self, sheetname: str)->worksheet:
        self.sheet = self.wb.get_sheet_by_name(sheetname)
        return self.sheet

    def delete_columns(self, col_range: str)->None:
        col_range = col_range.split(":")
        start = column_index_from_string(col_range[0])
        try:
            end = column_index_from_string(col_range[1])
        except IndexError:
            end = start + 1
        self.sheet.delete_cols(start, end - start)

    def delete_rows(self, row_range: str)->None:
        for row in row_range.split(":"):
            self.sheet.delete_rows(int(row))

    def count_by_range(self, range_:str)->tuple:
        column = self.sheet[range_].column
        row = self.sheet[range_].row
        col_length = len([column for column in self.sheet.columns][column - 1])
        row_length = len([row for row in self.sheet.rows][row - 1])

        return (col_length, row_length)
    
    def get_cells_by_filter(self, filter_:list, only_data=True)->list:
        raise NotImplementedError

if __name__ == '__main__':
    path = "REMADV202106110102_1047427.xls"
    advanced_xlsx = AdvancedXlsx()
    wb = advanced_xlsx.open_xls(path)
    wb.save("test.xlsx")