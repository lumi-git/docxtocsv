
#"C:\Users\ronan\AppData\Local\pandoc\pandoc.exe "

import csv
import subprocess
from docx import Document

def table_to_csv(table, csv_filename):
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        for row in table.rows:
            writer.writerow([cell.text for cell in row.cells])


doc = Document('testDoc.docx')


for i, table in enumerate(doc.tables):
    table_to_csv(table, f'table_{i}.csv')

