import itertools
import openpyxl
from bs4 import BeautifulSoup

class DataInconsistentError(Exception):
    pass

class Extractor:
    def __init__(self, data):
        self.raw_data = data

        #This list content multiple sublists correspond to the html tables which are inputed
        self.processed_data = []

        #This list content data that will be export
        self.contents = []

        self.is_merged = False
        self.proceed()
        self.table_type = self.get_type()

    def proceed(self):
        if isinstance(self.raw_data, list):
            html_list = self.raw_data
        else:
            html_list = [self.raw_data]

        for html in html_list:
            grid = []
            html = str(html)
            html = ''.join(line.strip() for line in html.split('\n'))

            soup = BeautifulSoup(html, 'html.parser')
            table_tag = soup.table

            #Adding rows
            for row_tag in table_tag.contents:
                row = []
                for data_tag in row_tag.children:
                    col = [data_tag.text.strip(),
                           (int(data_tag.attrs['colspan']) if 'colspan' in data_tag.attrs else 1),
                           (int(data_tag.attrs['rowspan']) if 'rowspan' in data_tag.attrs else 1)]
                    row.append(col)
                grid.append(row)

            #Make symmetrical grid
            for row in grid:
                for col in row:
                    if col[1] > 1: #colspan > 1
                        for _ in range(1, col[1]):
                            row.insert(row.index(col)+1, [None, 1, col[2]])
                    if col[2] > 1: #rowspan > 1
                        curt_col = row.index(col)
                        curt_row = grid.index(row)
                        for i in range(1, col[2]):
                            grid[curt_row+i].insert(curt_col, [None, 1, 1])

            #Remove colspan and rowspan values
            for row in grid:
                for i in range(0, len(row)):
                    row[i] = row[i][0]

            self.processed_data.append(grid)

    @staticmethod
    def invert(lst):
        return [list(x) for x in zip(*lst)]

    def get_type(self):
        header = self.processed_data[0][0]
        if all(table[0] == header for table in self.processed_data):
            return 'horizontal_header'
        else:
            header = [self.processed_data[0][i][0] for i in range(len(self.processed_data[0]))]
            if all(self.invert(table)[0] == header for table in self.processed_data):
                return 'vertical_header'
            else:
                return 'type_inconsistent'

    def merge(self):
        if self.table_type == 'horizontal_header':
            #Remove heads of all tables after the first table
            for i in range(1, len(self.processed_data)):
                del self.processed_data[i][0]
            #Combine all tables into one
            self.contents = list(itertools.chain.from_iterable(self.processed_data))


        elif self.table_type == 'vertical_header':
            for i in range(len(self.processed_data)):
                self.processed_data[i] = self.invert(self.processed_data[i])
            for i in range(1, len(self.processed_data)):
                del self.processed_data[i][0]
            self.contents = self.invert(list(itertools.chain.from_iterable(self.processed_data)))

        elif self.table_type == 'type_inconsistent':
            raise DataInconsistentError('Unable to merge tables')

        self.is_merged = True

    def dump(self, spreadsheet_name):
        #Combine all tables in case self.merge() has't been called
        if not self.is_merged:
            self.contents = list(itertools.chain.from_iterable(self.processed_data))

        wb = openpyxl.Workbook()
        sheet = wb.active

        for r in range(0, len(self.contents)):
            for c in range(0, len(self.contents[r])):
                sheet.cell(row=r+1, column=c+1).value = self.contents[r][c]

        wb.save(spreadsheet_name + '.xlsx')
