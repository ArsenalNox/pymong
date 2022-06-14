
import getter as gtr
import pprint  as pp
import writer  as wrt

from datab import client

needed_modules = [ #названия модулей, по которым происходит выгрузка
        "Фонетика и графика",
        "Арифметические действия"
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
        
        
        print('\nОтветы на модуль')
        #Начать итерацию над всеми ответами 
        for q_data in all_answers:
            print('\n')
            pp.pprint(q_data)

            account = gtr.get_test_result_account_by_id(q_data["accountId"])
            print(account['nickname'])

            #Проверить, ни записанн ли уже результат этого ученикеа 
            if account['nickname'] in already_written_results_by_student_id:
                continue
            
            #Записать в строку ответы одного ученика
            row = wrt.write_single_answer_data(worksheet, workbook, row, account['nickname'], 'test')

            already_written_results_by_student_id.append(account['nickname'])

        #Запомнить строки в которых находятся результаты одного класса
        #Нарисовать формулу со статистикой этого класса
        #Добавить условное форматирование 
        #Добавить возможность сортировать данные внутри класса
        #Добавить дерева?


        print(f'Получение результатов модуля {result["name"]}...')
        print('\nИнформация о модуле:')
        pp.pprint(result)
        print('\n')


        print('Закрытие файла...')
        workbook.close()

