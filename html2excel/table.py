from bs4 import BeautifulSoup
import openpyxl
import itertools

class DataInconsistentError(Exception):
    pass

class Table:
    def __init__(self):
        self.contents = [] #list of tables
        self.html_list = []
        # self.contents = []

    # def merge(self, table_list):
    #     for table in table_list:
    #         self.contents += self.parse(table)
    #     # self.contents = sum(self.parse(table) for table in table_list)

    # def load(self, table_doc):
    #     self.contents = self.parse(table_doc)

    # @staticmethod
    # def parse(table_doc):
    #     grid = []
    #     table_doc = str(table_doc)
    #     table_doc = ''.join(line.strip() for line in table_doc.split('\n'))

    #     soup = BeautifulSoup(table_doc, 'html.parser')
    #     table_tag = soup.table

    #     #Adding values
    #     for row_tag in table_tag.contents:
    #         row = []
    #         for data_tag in row_tag.children:
    #             col = [data_tag.text,
    #                    (int(data_tag.attrs['colspan']) if 'colspan' in data_tag.attrs else 1),
    #                    (int(data_tag.attrs['rowspan']) if 'rowspan' in data_tag.attrs else 1)]
    #             row.append(col)
    #         grid.append(row)

    #     #Make symmetrical table
    #     for row in grid:
    #         for col in row:
    #             if col[1] > 1: #colspan > 1
    #                 for _ in range(1, col[1]):
    #                     row.insert(row.index(col)+1, [None, 1, col[2]])
    #             if col[2] > 1: #rowspan > 1
    #                 curt_col = row.index(col)
    #                 curt_row = grid.index(row)
    #                 for i in range(1, col[2]):
    #                     grid[curt_row+i].insert(curt_col, [None, 1, 1])

    #     #Remove colspan and rowspan values
    #     for row in grid:
    #         for i in range(0, len(row)):
    #             row[i] = row[i][0]

    #     return grid


    def loads(self, table_doc):
        if isinstance(table_doc, list):
            self.html_list = table_doc
        else:
            self.html_list.append(table_doc)

        for html in self.html_list:
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

            #Make symmetrical table
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

            self.contents.append(grid)

    def merge(self):
        head = self.contents[0][0]
        # print(head)
        if all(table[0] == head for table in self.contents):
            pass
        for i in range(len(self.contents)):
            if self.contents[i][0] == head:
                del self.contents[i][0]
            else:
                # print('Unable to merge tables')
                # return
                raise DataInconsistentError('Unable to merge tables')

        self.contents.insert(0, [head])


    def dump(self, spreadsheet_name):
        contents = list(itertools.chain.from_iterable(self.contents))
        wb = openpyxl.Workbook()
        sheet = wb.active

        for r in range(0, len(contents)):
            for c in range(0, len(contents[r])):
                sheet.cell(row=r+1, column=c+1).value = contents[r][c]

        wb.save(spreadsheet_name + '.xlsx')

    def del_row(self, row_num):
        pass


class SingleTable(Table):
    def __init__(self, table_doc):
        self.contents = []
        self.html_list = []
        self.loads(table_doc)


def load(something):
    return SingleTable(something)