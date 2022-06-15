
import getter as gtr
import pprint  as pp
import writer  as wrt

from xlsxwriter.utility   import xl_cell_to_rowcol, xl_rowcol_to_cell 
from datab import client

needed_modules = [ #названия модулей, по которым происходит выгрузка
        "Арифметические действия",
        ]

x = gtr.get_modules_collections(needed_modules)

for data in x: #???? Люблю монго
    for result in data:
        #Итерация через ответы 

        print('Создание Excel файла...')
        workbook_name = gtr.generate_module_file_name(result['name'])

        print(f'Название Excel файла: "{workbook_name}" ')
        workbook, worksheet = wrt.create_workSheet(workbook_name)
        

        #Получить список вопросов данного модуля 
        module_questions = []
        print('Получение информации о вопросах')
        for question in result['questionIds']:
            print(f'\nВопрос {question["_id"]}')

            q_res = gtr.get_module_question(question["_id"])
            
            for data in q_res:
                pp.pprint(data)

                module_questions.append(data)

        print(f'\nОбщее кол-во вопросов модуля: {len(module_questions)}\n')


        #Отрисовать шапку вопросов  
        row = wrt.write_question_header(worksheet, 0, len(module_questions), workbook)

        #Получить все ответы на данный модуль
        already_written_results_by_student_id = []
        all_answers = gtr.get_module_answers(result['_id'])
        
        row_start = row

        print(f'Начальная строка: {row}')
        print('\nОтветы на модуль') 
        #Начать итерацию над всеми ответами 
        for q_data in all_answers:
            print('\n')
            pp.pprint(q_data)
            answers = []
            iterator = 0
            for tdt in q_data['questions']:
                answers.append( tdt['isCorrect'])

            account = gtr.get_test_result_account_by_id(q_data["accountId"])
            print(account['nickname'])
            print(answers)

            #Проверить, ни записанн ли уже результат этого ученикеа 
            if account['nickname'] in already_written_results_by_student_id:
                continue
            
            #Записать в строку ответы одного ученика
            row = wrt.write_single_answer_data(worksheet, workbook, row, account['nickname'], answers, result['questionIds'], q_data['questions'])

            #Запомнить строки в которых находятся результаты одного класса
            already_written_results_by_student_id.append(account['nickname'])
               
        
        #Нарисовать формулу со статистикой этого класса
        col = 2
        for i in range(2, len(module_questions)+2, 1):

            
            worksheet.write(3, col, f'=COUNTIF({xl_rowcol_to_cell(row-1, col)}:{xl_rowcol_to_cell(row_start, col)}, "1")')
            worksheet.write(2, col, f'=COUNTIF({xl_rowcol_to_cell(row-1, col)}:{xl_rowcol_to_cell(row_start, col)}, "0")') 
#            worksheet.write(1, col, f'=IFERROR({xl_rowcol_to_cell(2, col)}/({xl_rowcol_to_cell(3,col)}+{xl_rowcol_to_cell(2,col)}),0)')
            worksheet.write(1, col, f'=IFERROR({xl_rowcol_to_cell(3, col)}/({xl_rowcol_to_cell(3,col)}+{xl_rowcol_to_cell(2,col)}),0)')

            worksheet.conditional_format(
                    f'C2:T2',
                    {
                        "type": '3_color_scale',
                        "min_color": 'red',
                        "mid_color": 'yellow',
                        "max_color": 'green',
                        "mid_value": '50%',
                        "max_value": '100%',
                        "min_value": '0%',
                        "min_type": 'num',
                        "max_type": 'num',
                        "mid_type": 'num'
                        }
                )
            col += 1

        #Добавить условное форматирование 
        #Добавить возможность сортировать данные внутри класса
        #Добавить дерева?

        print(f'Получение результатов модуля {result["name"]}...')
        print('\nИнформация о модуле:')
        pp.pprint(result)
        
        for qst in result['questionIds']:
            print(qst)

        print('\n')


        print('Закрытие файла...')
        workbook.close()

