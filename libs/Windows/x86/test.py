import pandas as pd
from lxml import html
from lxml.html import parse
from io import StringIO
from pandas.io.parsers import TextParser
from bs4 import BeautifulSoup

def _unpack(row, kind='td'):
   elts = row.findall('.//%s' % kind)
   return [val.text_content() for val in elts]

def parse_options_data(table):
  rows = table.findall('.//tr')
  header = _unpack(rows[0], kind='th')
  data = [_unpack(r) for r in rows[1:]]
  return TextParser(data, names=header).get_chunk()

parsed = html.parse(open(r"C:\Users\jmsir\Downloads\ArchivoExcel.xls", 'r').read())
doc = parsed.getroot()
tables = doc.findall('.//table')
table = parse_options_data(tables[0])