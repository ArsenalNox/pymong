from datab import client
from bson.objectid import ObjectId

import datetime

db = client["ApiDefaultDB"]

def get_modules_collections(module_names:list):
    global db
    modules = []


    # Collection Name
    col = db["modules"]
     
    for module_name in module_names:
        print(f'Searching for {module_name}')
        modules.append(col.find({"name":f"{module_name}"}))
    
    return modules


def get_module_answers(module_id):
    global db
    col = db['tests']
    result = col.find({
        "moduleId": ObjectId(module_id),
        "close": True
        })

    return result


def get_module_question(question_id):
    global db 

    col = db['questions']
    result = col.find({"_id":ObjectId(question_id)})

    return result

def generate_module_file_name(module_name:str)->str:
    global db
    date = datetime.datetime.now().strftime("%d_%m_%Y-%H_%M")
    module_name = module_name.replace(' ', '_')
    new_name = f"{module_name}_{date}"
    return new_name


def get_test_result_account_by_id(accound_id):
    global db
    col = db['users']
    account = col.find({"_id":ObjectId(accound_id)})
    for data in account:
        account_data = data

        return account_data


