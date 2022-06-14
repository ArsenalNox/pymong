"""
Записывает результаты в Excel
"""

import xlsxwriter

def create_workSheet(name):
    xlsxWorkBook  = xlsxwriter.Workbook(f'data/{name}.xlsx')
    xlsxWorkSheet = xlsxWorkBook.add_worksheet()

    return xlsxWorkBook, xlsxWorkSheet


def write_question_header(worksheet, row, q_num, workbook):

    col = 2

    format_yellow_bkg = workbook.add_format()
    format_yellow_bkg.set_bg_color('ffc000')

    worksheet.write(row, col-1, 'Кол-во вопросов')

    for i in range(0, q_num, 1):
        worksheet.write(row, col, i+1)
        col+=1

    worksheet.write(row+1, 1, 'Всего неправильных', format_yellow_bkg)
    worksheet.write(row+2, 1, 'Всего правильных', format_yellow_bkg)

    return row+3


def write_single_answer_data(worksheet, workbook, row, nickname, answer_data):
    worksheet.write(row, 1, nickname)
    worksheet.write(row, 2, 'данные ответа')

    return row+1

