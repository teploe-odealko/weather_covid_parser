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
from copy import copy
import re

urllib3.disable_warnings()

SHEET_NAME_CASES = 'Случаев'
SHEET_NAME_BASE = 'База РФ'

URL = 'https://yastat.net/s3/milab/2020/covid19-stat/data/v10/deep_data.json'
REGION_COLUMN = 2
REGION_ROWS_RANGE = (2, 87)
WEEKLY_REPORT_DAY = 4


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
    # print("last column", last_column)
    last_table_date = ws.cell(row=1, column=last_column).value
    new_days_amount = (last_available_date - last_table_date).days
    for i in range(1, new_days_amount + 1):
        ws.cell(row=1, column=(last_column + i)).value = (last_table_date + timedelta(days=i))
        ws.cell(row=1, column=(last_column + i)).number_format = 'DD.MM.YYYY'

        for row in range(*REGION_ROWS_RANGE):
            name = ws.cell(row=row, column=REGION_COLUMN).value
            key = normolize_string(name)
            if key in regions_full_name:
                ws.cell(row=row, column=(last_column + i)).value = \
                    regions_full_name[key]['cases'][(i - 1) - new_days_amount][0]

            elif key in regions_short_name:
                ws.cell(row=row, column=(last_column + i)).value = \
                    regions_short_name[key]['cases'][(i - 1) - new_days_amount][0]

            else:
                raise KeyError

        formula = ws.cell(row=OVERALL_CASES_FORMULA_ROW, column=ws.max_column - 1).value
        last_cell_id = get_column_letter(ws.max_column - 1) + str(OVERALL_CASES_FORMULA_ROW)
        next_cell_id = get_column_letter(ws.max_column) + str(OVERALL_CASES_FORMULA_ROW)
        ws[next_cell_id] = Translator(formula, origin=last_cell_id).translate_formula(next_cell_id)
    return {'new_days_amount': new_days_amount,
            'last_date': last_table_date,
            'current_date': last_table_date + timedelta(days=new_days_amount)}


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


def continue_formula_right(ws: Worksheet, rows: list, column: int):
    for row in rows:
        formula = ws.cell(row=row, column=column - 1).value
        last_cell_id = get_column_letter(column - 1) + str(row)
        next_cell_id = get_column_letter(column) + str(row)
        ws[next_cell_id] = Translator(formula, origin=last_cell_id).translate_formula(next_cell_id)
        ws[next_cell_id]._style = copy(ws[last_cell_id]._style)


def continue_date_right(ws: Worksheet, rows: list):
    for row in rows:
        try:
            column = [c.value for c in ws[row]].index(None) + 1
        except ValueError:
            column = ws.max_column + 1
        last_table_date = ws.cell(row=row, column=column - 1).value
        ws.cell(row=row, column=column).value = (last_table_date) + timedelta(days=1)
        ws.cell(row=row, column=column).number_format = 'DD.MM.YYYY'

        last_cell_id = get_column_letter(column - 1) + str(row)
        next_cell_id = get_column_letter(column) + str(row)
        ws[next_cell_id]._style = copy(ws[last_cell_id]._style)


def continue_number_right(ws: Worksheet, rows: list):
    for row in rows:
        try:
            column = [c.value for c in ws[row]].index(None) + 1
        except ValueError:
            column = ws.max_column + 1
        last_num = ws.cell(row=row, column=column - 1).value
        ws.cell(row=row, column=column).value = last_num + 1

        last_cell_id = get_column_letter(column - 1) + str(row)
        next_cell_id = get_column_letter(column) + str(row)
        ws[next_cell_id]._style = copy(ws[last_cell_id]._style)


def gain_sheet(ws: Worksheet, info_dict: dict):
    NEW_COLUMN_NUM_GAIN = [c.value for c in ws[1]].index(None) + 1

    for i in range(info_dict['new_days_amount']):
        continue_date_right(ws, [1])
        continue_formula_right(ws, list(range(2, 96)), NEW_COLUMN_NUM_GAIN + i)


def daily_gain_sheet(ws: Worksheet, info_dict: dict):
    NEW_COLUMN_NUM_DAILY_GAIN = ws.max_column + 1

    for i in range(info_dict['new_days_amount']):
        continue_date_right(ws, [1, 20, 35])
        continue_formula_right(ws,
                               list(range(2, 19)) + list(range(21, 34)) + list(range(36, 49)),
                               NEW_COLUMN_NUM_DAILY_GAIN + i)

def weekly_gain_sheet(wb, info_dict: dict):
    ws_weekly_gain = wb['Прирост нед']
    print(info_dict['current_date'].weekday() + 1, WEEKLY_REPORT_DAY)
    if (info_dict['current_date'].weekday() + 1) == WEEKLY_REPORT_DAY:
        new_col = [c.value for c in ws_weekly_gain[1]].index(None) + 1
        print(new_col)
        ws_weekly_gain.insert_cols(new_col)
        continue_number_right(ws_weekly_gain, [1])
        continue_formula_right(ws_weekly_gain, list(range(2, 94)), new_col)
        for row in range(2, 87):
            raw_formula = ws_weekly_gain.cell(row=row, column=new_col).value
            tokenized_formula = Tokenizer(raw_formula)
            last_col_in_gain_sheet = [c.value for c in wb['Прирост'][1]].index(None)

            tokenized_formula.items[1].value = re.sub(r'(?<=:\$)\D*',
                                                      get_column_letter(last_col_in_gain_sheet),
                                                      tokenized_formula.items[1].value)
            tokenized_formula.items[3].value = re.sub(r'(?<=:\$).*(?=\$)',
                                                      get_column_letter(last_col_in_gain_sheet),
                                                      tokenized_formula.items[3].value)
            new_formula = ''.join(token.value for token in tokenized_formula.items)
            ws_weekly_gain.cell(row=row, column=new_col).value = fr'={new_formula}'

def cases_sheet(wb):
    regions_info = get_regions_info()
    info_dict = fill_сases_sheet(wb[SHEET_NAME_CASES], *regions_info)

    gain_sheet(wb['Прирост'], info_dict)
    daily_gain_sheet(wb['По рег прис (сут)'], info_dict)
    weekly_gain_sheet(wb, info_dict)

        # continue_number_right(ws_weekly_gain, )

def continue_formula_n_down(ws: Worksheet, columns: list, row: int, n: int):
    for column in columns:
        formula = ws.cell(row=(row - n), column=column).value
        last_cell_id = get_column_letter(column) + str(row - n)
        next_cell_id = get_column_letter(column) + str(row)
        ws[next_cell_id] = Translator(formula, origin=last_cell_id).translate_formula(next_cell_id)
        ws[next_cell_id]._style = copy(ws[last_cell_id]._style)


def continue_formula_down(ws: Worksheet, columns: list, row: int):
    continue_formula_n_down(ws, columns, row, 1)


def base_rf_sheet(ws: Worksheet):
    STOP_COVID_URL = 'https://стопкоронавирус.рф/information/'
    NEW_ROW_NUM = max((c.row for c in ws['B'] if c.value is not None)) + 1
    session = HTMLSession()
    response = session.get(STOP_COVID_URL, verify=False)
    response.html.render()

    actual_date = response.html.xpath('//small/text()')[0]
    actual_date = dateparser.parse(''.join(actual_date.split()[3:]))
    print(actual_date.weekday())
    ws.cell(row=NEW_ROW_NUM, column=1).value = actual_date
    ws.cell(row=NEW_ROW_NUM, column=1).number_format = 'DD.MM'

    rows = response.html.xpath('//h3[@class="cv-stats-virus__item-value"]/text()')
    statistics = [int(''.join(row.strip('\n ').split())) for row in rows]
    useful_statistics = [statistics[2], statistics[4], statistics[0]]
    for i, stat in enumerate(useful_statistics):
        ws.cell(row=NEW_ROW_NUM, column=2 + i).value = useful_statistics[i]
        ws.cell(row=NEW_ROW_NUM, column=2 + i)._style = copy(ws.cell(row=NEW_ROW_NUM - 1, column=2 + i)._style)

    continue_formula_down(ws,
                          list(range(5, 20)) + [26] + list(range(28, 36)),
                          NEW_ROW_NUM)

    if actual_date.weekday() == 6:  # if it's sunday
        continue_formula_n_down(ws,
                                list(range(20, 25)),
                                NEW_ROW_NUM,
                                7)
        continue_formula_n_down(ws,
                                list(range(20, 23)),
                                NEW_ROW_NUM + 1,
                                7)
        continue_formula_n_down(ws,
                                [27],
                                NEW_ROW_NUM,
                                7)

    return actual_date


def date_week_sheet(ws: Worksheet, actual_date: datetime):
    NEW_ROW_NUM = max((c.row for c in ws['A'] if c.value is not None)) + 1
    if actual_date.weekday() == WEEKLY_REPORT_DAY:
        ws.cell(row=NEW_ROW_NUM, column=1).value = ws.cell(row=NEW_ROW_NUM - 1, column=1).value + 1
        continue_formula_down(ws,
                              list(range(1, 3)),
                              NEW_ROW_NUM)


# def daily_gain_sheet(ws: Worksheet, actual_date: datetime):


def main(filename: str):
    wb = openpyxl.load_workbook(filename)
    cases_sheet(wb)
    actual_date = base_rf_sheet(wb[SHEET_NAME_BASE])

    date_week_sheet(wb['Дата-неделя'], actual_date)

    # daily_gain_sheet(wb['По рег прис (сут)'], actual_date)

    wb.save('russia_regions.xlsx')


if __name__ == '__main__':
    with open('conf_covid', 'r') as f:
        filename = f.read().strip()
    main(filename)
