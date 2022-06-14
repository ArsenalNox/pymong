from datab import client
import datetime

def get_modules_collections(module_names:list):
    modules = []

    db = client["ApiDefaultDB"]

    # Collection Name
    col = db["modules"]
     
    for module_name in module_names:
        print(f'Searching for {module_name}')
        modules.append(col.find({"name":f"{module_name}"}))
    
    return modules


def get_module_answers(module_id:str):
    print(module_id)
    pass


def generate_module_file_name(module_name:str)->str:
    date = datetime.datetime.now().strftime("%d_%m_%Y-%H_%M")
    module_name = module_name.replace(' ', '_')
    new_name = f"{module_name}_{date}"
    return new_name
