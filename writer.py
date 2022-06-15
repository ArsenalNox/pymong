"""
Записывает результаты в Excel
"""

from bson.objectid import ObjectId
import xlsxwriter

from xlsxwriter.utility   import xl_cell_to_rowcol, xl_rowcol_to_cell 

def create_workSheet(name):
    xlsxWorkBook  = xlsxwriter.Workbook(f'data/{name}.xlsx')
    xlsxWorkSheet = xlsxWorkBook.add_worksheet()

    return xlsxWorkBook, xlsxWorkSheet


def write_question_header(worksheet, row, q_num, workbook):

    col = 2

    format_yellow_bkg = workbook.add_format()
    format_yellow_bkg.set_bg_color('ffc000')

    worksheet.write(row, col-1, 'Номера вопросов')

    for i in range(0, q_num, 1):
        worksheet.write(row, col, i+1)
        col+=1

    worksheet.write(row+1, 1, '% выполнения', format_yellow_bkg)
    worksheet.write(row+2, 1, 'Всего неправильных', format_yellow_bkg)
    worksheet.write(row+3, 1, 'Всего правильных', format_yellow_bkg)

    return row+4


def write_single_answer_data(worksheet, workbook, row, nickname, answer_data, questions, q_data_orig):

    def get_q_number(index, q_data_orig, questions):
        global answer_data
        
        iterator2 = 0
        for question in questions:
            if ObjectId(q_data_orig[index]['questionId']) == ObjectId(question['_id']):
                
                return iterator2

            iterator2 += 1
        return 1
        
    #First we need to lay gray area 

    worksheet.write(row, 1, nickname)
    col = 2
    
    grey_formaty = workbook.add_format({'bg_color': 'AAAAAA'})
    green_format = workbook.add_format({'bg_color': '00dd00'})
    red_format   = workbook.add_format({'bg_color': 'dd0000'})
 
    for i in range(2, len(questions)+2, 1):
        worksheet.write(row, i, '', grey_formaty)

   
    corr_answ = 0
    iterator = 0

    for answ in answer_data:

        col2 = get_q_number(iterator, q_data_orig, questions) + 2

        if answ: 
            #Need to change row to positional number of question in module
            worksheet.write(row, col2, '1', green_format)    
            corr_answ+=1

        else: 
            worksheet.write(row, col2, '0', red_format)

        col+=1
        iterator+=1


        worksheet.conditional_format(
                f'{xl_rowcol_to_cell(row,col)}:{xl_rowcol_to_cell(row,col)}',
                {
                    "type": '3_color_scale',
                    "min_color": 'red',
                    "mid_color": 'yellow',
                    "max_color": 'green',
                    "mid_value": '0.5',
                    "max_value": '1',
                    "min_value": '0',
                    "min_type": 'num',
                    "max_type": 'num',
                    "mid_type": 'num'
                    }
            )


    return row+1
