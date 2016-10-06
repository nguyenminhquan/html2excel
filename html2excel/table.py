from bs4 import BeautifulSoup
import openpyxl

class Table:
    def __init__(self):
        self.content = []

    def load(self, table_doc):
        table_doc = str(table_doc)
        table_doc = ''.join(line.strip() for line in table_doc.split('\n'))

        soup = BeautifulSoup(table_doc, 'html.parser')
        table_tag = soup.table

        #Adding values
        for row_tag in table_tag.contents:
            row = []
            for data_tag in row_tag.children:
                col = [data_tag.text,
                       (int(data_tag.attrs['colspan']) if 'colspan' in data_tag.attrs else 1),
                       (int(data_tag.attrs['rowspan']) if 'rowspan' in data_tag.attrs else 1)]
                row.append(col)
            self.content.append(row)

        #Make symmetrical table
        for row in self.content:
            for col in row:
                if col[1] > 1: #colspan > 1
                    for _ in range(1, col[1]):
                        row.insert(row.index(col)+1, [None, 1, col[2]])
                if col[2] > 1: #rowspan > 1
                    curt_col = row.index(col)
                    curt_row = self.content.index(row)
                    for i in range(1, col[2]):
                        self.content[curt_row+i].insert(curt_col, [None, 1, 1])

        #Remove colspan and rowspan values
        for row in self.content:
            for i in range(0, len(row)):
                row[i] = row[i][0]

    def dump(self, spreadsheet_name):

        wb = openpyxl.Workbook()
        sheet = wb.active

        for r in range(0, len(self.content)):
            for c in range(0, len(self.content[r])):
                sheet.cell(row=r+1, column=c+1).value = self.content[r][c]

        wb.save(spreadsheet_name + '.xlsx')