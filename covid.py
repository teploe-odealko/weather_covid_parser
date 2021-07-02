import openpyxl
from typing import Dict, List, Tuple

from bs4 import BeautifulSoup
from requests_html import HTMLSession
import urllib3
from openpyxl.worksheet.worksheet import Worksheet
from nltk.stem.snowball import SnowballStemmer
from datetime import datetime, date, time, timedelta
from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import get_column_letter
import dateparser
urllib3.disable_warnings()


SHEET_NAME_CASES = 'Случаев'
SHEET_NAME_BASE = 'База РФ'

URL = 'https://yastat.net/s3/milab/2020/covid19-stat/data/v10/deep_data.json'
REGION_COLUMN = 2
REGION_ROWS_RANGE = (2, 87)


def normolize_string(string: str):
    string = string.strip().lower().translate({ord('.'): None})
    string = ' '.join(string.split('-'))

    # remove stop words
    stop_words = ['г', 'обл', 'область', 'республика', 'автономная', 'ао']
    try:
        splitted_string = string.split()
    except IndexError:
        return ''
    stemmer = SnowballStemmer("russian")
    string = ' '.join(stemmer.stem(w) for w in splitted_string if w not in stop_words)
    return string


def fill_сases_sheet(ws: Worksheet, regions_full_name: dict, regions_short_name, dates):
    OVERALL_CASES_FORMULA_ROW = 88

    last_available_date = datetime.strptime(dates[-1], '%Y-%m-%d')
    last_column = ws.max_column
    print("last column", last_column)
    last_table_date = ws.cell(row=1, column=last_column).value
    days_amount_to_parse = (last_available_date - last_table_date).days
    for i in range(1, days_amount_to_parse + 1):
        ws.cell(row=1, column=(last_column + i)).value = (last_table_date + timedelta(days=i))
        ws.cell(row=1, column=(last_column + i)).number_format = 'DD.MM.YYYY'

        for row in range(*REGION_ROWS_RANGE):
            name = ws.cell(row=row, column=REGION_COLUMN).value
            key = normolize_string(name)
            if key in regions_full_name:
                ws.cell(row=row, column=(last_column + i)).value = \
                    regions_full_name[key]['cases'][(i-1) - days_amount_to_parse][0]

            elif key in regions_short_name:
                ws.cell(row=row, column=(last_column + i)).value = \
                    regions_short_name[key]['cases'][(i-1) - days_amount_to_parse][0]

            else:
                raise KeyError

        formula = ws.cell(row=OVERALL_CASES_FORMULA_ROW, column=ws.max_column - 1).value
        last_cell_id = get_column_letter(ws.max_column - 1) + str(OVERALL_CASES_FORMULA_ROW)
        next_cell_id = get_column_letter(ws.max_column) + str(OVERALL_CASES_FORMULA_ROW)
        ws[next_cell_id] = Translator(formula, origin=last_cell_id).translate_formula(next_cell_id)


def get_regions_info() -> Tuple[dict, dict, list]:
    regions_full_name = {}
    regions_short_name = {}

    session = HTMLSession()
    response: dict = session.get(URL, verify=False).json()
    regions = response['russia_stat_struct']['data']
    dates = response['russia_stat_struct']['dates']
    for region in regions:
        full_name = regions[region]['info']['name']
        short_name = regions[region]['info']['short_name']
        regions[region]['info'].pop('name', None)
        regions[region]['info'].pop('short_name', None)
        regions_full_name[normolize_string(full_name)] = regions[region]
        regions_short_name[normolize_string(short_name)] = regions[region]
    return regions_full_name, regions_short_name, dates


def cases_sheet(ws: Worksheet):
    regions_info = get_regions_info()
    fill_сases_sheet(ws, *regions_info)


def continue_formula_down(ws: Worksheet, columns: list, row: int):
    for column in columns:
        formula = ws.cell(row=(row-1), column=column).value
        last_cell_id = get_column_letter(column) + str(row-1)
        next_cell_id = get_column_letter(column) + str(row)
        ws[next_cell_id] = Translator(formula, origin=last_cell_id).translate_formula(next_cell_id)

def base_rf_sheet(ws: Worksheet):
    STOP_COVID_URL = 'https://стопкоронавирус.рф/information/'
    NEW_ROW_NUM = max((c.row for c in ws['B'] if c.value is not None)) + 1
    session = HTMLSession()

    response = session.get(STOP_COVID_URL, verify=False)
    response.html.render()

    actual_date = response.html.xpath('//small/text()')[0]
    # today = datetime.today()
    # today.replace(day=int(actual_date.split()[3]))
    actual_date = dateparser.parse(''.join(actual_date.split()[3:]))
    print(actual_date)
    ws.cell(row=NEW_ROW_NUM, column=1).value = actual_date
    ws.cell(row=NEW_ROW_NUM, column=1).number_format = 'DD.MM'


    rows = response.html.xpath('//h3[@class="cv-stats-virus__item-value"]/text()')
    statistics = [int(''.join(row.strip('\n ').split())) for row in rows]
    cases = statistics[0]
    recovered = statistics[2]
    death = statistics[4]
    ws.cell(row=NEW_ROW_NUM, column=2).value = recovered
    ws.cell(row=NEW_ROW_NUM, column=3).value = death
    ws.cell(row=NEW_ROW_NUM, column=4).value = cases

    print(NEW_ROW_NUM)
    continue_formula_down(ws, list(range(5, 20)), NEW_ROW_NUM)
    # print(actual_date)
    # print(today)

def main(filename: str):
    wb = openpyxl.load_workbook(filename)
    cases_sheet(wb[SHEET_NAME_CASES])
    base_rf_sheet(wb[SHEET_NAME_BASE])


    wb.save(filename)


if __name__ == '__main__':
    with open('conf_covid', 'r') as f:
        filename = f.read().strip()
    main(filename)