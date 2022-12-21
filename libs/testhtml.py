from tablepyxl import tablepyxl

doc = open(r"C:/Users/jmsir/Downloads/html_table.xls", "r")
table = doc.read()

wb = tablepyxl.document_to_workbook(table)
wb.save(r"C:/Users/jmsir/Downloads/html_table.xlsx")