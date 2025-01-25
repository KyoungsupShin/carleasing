import glob
import xlwings as xw
import pandas as pd
import itertools
import copy

def bnk_brand_idx(sh, x):
    x = x.strip().replace('\xa0', ' ')
    models = sh.range('J7', 'J36').value 
    models = [str(x).strip().replace('\xa0', ' ') for x in models]
    for idx, model in enumerate(models):
        if model == x:
            model_code = model[0]
            break
    return idx+1

def bnk_car_idx(sh,x):
    x = x.strip().replace('\xa0', ' ')
    models = sh.range('AM20', 'AM99').value 
    models = [str(x).strip().replace('\xa0', ' ') for x in models]
    for idx, model in enumerate(models):
        if model == x:
            model_code = model[0]
            break
    return idx+1

def bnk_car_model_idx(sh,x):
    x = str(x).strip().replace('\xa0', ' ')
    models = sh.range('AO20', 'AO256').value 
    models = [str(x).strip().replace('\xa0', ' ') for x in models]
    for idx, model in enumerate(models):
        if model == x:
            model_code = model[0]
            break
    return idx+1


def sh_brand_idx(sheet1, x):
    if x == str:
        x = x.strip()
    if x == float:
        x = int(x)
        
    models = sheet1.range('B10', 'B20').value 
    models = [str(x).strip().replace('\xa0', ' ') for x in models]
    for idx, model in enumerate(models):
        if model == x:
            model_code = model[0]
            break
    return idx+1

def sh_car_idx(sh,x):
    models = sh.range('C10', 'C100').value 
    return models[x]


def nh_brand_idx(sh, x):
    x = x.strip().replace('\xa0', ' ')
    models = sh.range('I43', 'I79').value 
    models = [str(x).strip().replace('\xa0', ' ') for x in models]
    for idx, model in enumerate(models):
        if model == x:
            model_code = model[0]
            break
    return idx+1

def nh_car_idx(sh,x):
    x = x.strip().replace('\xa0', ' ')
    models = sh.range('BZ3', 'BZ40').value 
    models = [str(x).strip().replace('\xa0', ' ') for x in models]
    for idx, model in enumerate(models):
        if model == x:
            model_code = model[0]
            break
    return idx+1

def nh_car_model_idx(sh,x):
    x = str(x).strip().replace('\xa0', ' ')
    models = sh.range('CL4', 'CL30').value 
    models = [str(x).strip().replace('\xa0', ' ') for x in models]
    for idx, model in enumerate(models):
        if model == x:
            model_code = model[0]
            break
    return idx+1
    
def generate_parameters(provider, input_datum, deposits, downpayments, periods):
    inputs = []  
    if provider == 'bnk':
        for idx, input_data in enumerate(input_datum):
            for combo in itertools.product(deposits, downpayments, periods):
                modified_input = copy.deepcopy(input_data)
                modified_input[provider]['deposit'][-1] = combo[0]
                modified_input[provider]['downpayment'][-1] = combo[1]
                modified_input[provider]['period'][-1] = combo[2]
                inputs.append(modified_input)
    if provider == 'sh':
        for idx, input_data in enumerate(input_datum):            
            for combo in itertools.product(deposits, downpayments, periods):
                modified_input = copy.deepcopy(input_data)
                modified_input[provider]['deposit'][-1] = combo[0] * 100
                modified_input[provider]['downpayment'][-1] = combo[1] * 100
                modified_input[provider]['period'][-1] = combo[2]
                inputs.append(modified_input)
    if provider == 'mz':
        for idx, input_data in enumerate(input_datum):            
            for combo in itertools.product(deposits, downpayments, periods):
                modified_input = copy.deepcopy(input_data)
                modified_input[provider]['deposit'][-1] = combo[0]
                modified_input[provider]['downpayment'][-1] = combo[1] * input_data[provider]['car_price'][2]
                modified_input[provider]['period'][-1] = combo[2]
                inputs.append(modified_input)
    if provider == 'im':
        for idx, input_data in enumerate(input_datum):            
            for combo in itertools.product(deposits, downpayments, periods):
                modified_input = copy.deepcopy(input_data)
                modified_input[provider]['deposit'][-1] = combo[0]
                modified_input[provider]['downpayment'][-1] = combo[1]
                modified_input[provider]['period'][-1] = combo[2]
                inputs.append(modified_input)
    if provider == 'lotte':
        for idx, input_data in enumerate(input_datum):            
            for combo in itertools.product(deposits, downpayments, periods):
                modified_input = copy.deepcopy(input_data)
                modified_input[provider]['deposit'][-1] = combo[0]
                modified_input[provider]['downpayment'][-1] = combo[1] * input_data[provider]['car_price'][2]
                modified_input[provider]['period'][-1] = combo[2]
                inputs.append(modified_input)
    if provider == 'nh':
        for idx, input_data in enumerate(input_datum):            
            for combo in itertools.product(deposits, downpayments, periods):
                modified_input = copy.deepcopy(input_data)
                modified_input[provider]['deposit'][-1] = combo[0]
                modified_input[provider]['downpayment'][-1] = combo[1] * input_data[provider]['car_price'][2]
                modified_input[provider]['period'][-1] = combo[2]
                inputs.append(modified_input)
    if provider == 'woori':
        for idx, input_data in enumerate(input_datum):            
            for combo in itertools.product(deposits, downpayments, periods):
                modified_input = copy.deepcopy(input_data)
                modified_input[provider]['deposit'][-1] = combo[0]
                modified_input[provider]['downpayment'][-1] = combo[1] * input_data[provider]['car_price'][2]
                modified_input[provider]['period'][-1] = combo[2]
                inputs.append(modified_input)

    return inputs
    
def convert_params(provider, brand_name, car_name, model_name, car_price, org_brand_name, org_car_name, org_model_name, org_model_detail):
    if provider == 'bnk':
        input_data = {
            "bnk": {
                'brand_name': [0, 'B9', brand_name],
                'car_name': [0, 'B13', car_name],
                'model_name': [0, 'B15', model_name],
                'car_price': [1, 'N13', car_price],
                'deposit': [1, 'N36', 0],
                'downpayment': [1, 'N38', 0],
                'period': [0, 'B39', 0],
                'org_brand_name': org_brand_name,
                'org_car_name': org_car_name,
                'org_model_name': org_model_name,
                'org_model_detail': org_model_detail,
            },
        }
    if provider == 'sh':
        input_data = {
            "sh": {
                'brand_name': [0, 'B9', brand_name],
                'car_name': [1, 'AE7', car_name],
                'model_name': [1, 'AE7', car_name],
                'car_price': [1, 'AF8', car_price],
                'deposit': [1, 'AC35', 0],
                'downpayment': [1, 'AC36', 0],
                'period': [1, 'AC32', 0],
                'org_brand_name': org_brand_name,
                'org_car_name': org_car_name,
                'org_model_name': org_model_name,
                'org_model_detail': org_model_detail,
            },
        }
    if provider == 'mz':
        input_data = {
            "mz": {
                'brand_name': [0, 'AT2', brand_name],
                'car_name': [1, 'F5', car_name],
                'model_name': [2, 'F5', car_name],
                'car_price': [2, 'P9', car_price],
                'deposit': [2, 'AZ11', 0],
                'downpayment': [2, 'AZ16', 0],
                'period': [0, 'Z6', 0],
                'org_brand_name': org_brand_name,
                'org_car_name': org_car_name,
                'org_model_name': org_model_name,
                'org_model_detail': org_model_detail,
            },
        }
    if provider == 'lotte':
        input_data = {
            "lotte": {
                'brand_name': [0, 'N11', brand_name],
                'car_name': [0, 'P11', car_name],
                'model_name': [0, 'R11', model_name],
                'car_price': [1, 'BQ10', car_price],
                'deposit': [1, 'BK26', 0],
                'downpayment': [1, 'BD25', 0],
                'period': [1, 'BD24', 0],
                'org_brand_name': org_brand_name,
                'org_car_name': org_car_name,
                'org_model_name': org_model_name,
                'org_model_detail': org_model_detail,
            },
        }
    if provider == 'im':
        input_data = {
            "im": {
                'brand_name': [0, 'S7', brand_name],
                'car_name': [1, 'AS7', car_name],
                'model_name': [1, 'AS7', model_name],
                'car_price': [1, 'AS11', car_price],
                'deposit': [1, 'AS30', 0],
                'downpayment': [1, 'AS33', 0],
                'period': [1, 'AS28', 0],
                'org_brand_name': org_brand_name,
                'org_car_name': org_car_name,
                'org_model_name': org_model_name,
                'org_model_detail': org_model_detail,
            },
        }
    if provider == 'woori':
        input_data = {
            "woori": {
                'brand_name': [0, 'BA6', brand_name],
                'car_name': [0, 'BA7', car_name],
                'model_name': [0, 'BA7', model_name],
                'car_price': [0, 'BA12', car_price],
                'deposit': [0, 'BA43', 0],
                'downpayment': [0, 'BA41', 0],
                'period': [0, 'BA39', 0],
                'org_brand_name': org_brand_name,
                'org_car_name': org_car_name,
                'org_model_name': org_model_name,
                'org_model_detail': org_model_detail,
            },
        }
    if provider == 'nh':
        input_data = {
            "nh": {
                'brand_name' : [0, 'BT3', brand_name],
                'car_name' : [0, 'BT9', car_name],
                'model_name' : [0, 'BT11', model_name],        
                'car_price' : [0, 'AY9', car_price],
                'deposit' : [0, 'AY26', 0],
                'downpayment' : [0, 'AZ24', 0],
                'period' : [0, 'AY23', 0],
                'org_brand_name': org_brand_name,
                'org_car_name': org_car_name,
                'org_model_name': org_model_name,
                'org_model_detail': org_model_detail,
            },
        }

    return input_data  # 디버깅을 위해 input_data 출력

def merge_output(provider, output):
    records = []
    if provider == 'bnk':
        for entry in output:
            input_data = entry['input'][provider]
            output_value = entry['output']
            record = {
                'carpan_brand_name': input_data['org_brand_name'],
                'carpan_car_name': input_data['org_car_name'],
                'carpan_model_name': input_data['org_model_name'],
                'carpan_model_detail': input_data['org_model_detail'],
                'provider' : provider,
                'brand_name': input_data['brand_name'][2],
                'car_name': input_data['car_name'][2],
                'model_name': input_data['model_name'][2],
                'car_price': input_data['car_price'][2],
                'deposit': input_data['deposit'][2],
                'downpayment': input_data['downpayment'][2],
                'period': input_data['period'][2],
                'output': output_value
            }
            records.append(record)
    if provider == 'sh':
        for entry in output:
            input_data = entry['input'][provider]
            output_value = entry['output']
            record = {
            'carpan_brand_name': input_data['org_brand_name'],
            'carpan_car_name': input_data['org_car_name'],
            'carpan_model_name': input_data['org_model_name'],
            'carpan_model_detail': input_data['org_model_detail'],
            'provider' :provider,
            'brand_name': input_data['brand_name'][2],
            'car_name': input_data['car_name'][2],
            'model_name': input_data['car_name'][2],
            'car_price': input_data['car_price'][2],
            'deposit': input_data['deposit'][2] / 100,
            'downpayment': input_data['downpayment'][2] / 100,
            'period': input_data['period'][2],
            'output': output_value
            }
            records.append(record)
    if provider == 'mz':
        for entry in output:
            input_data = entry['input'][provider]
            output_value = entry['output']
            record = {
            'carpan_brand_name': input_data['org_brand_name'],
            'carpan_car_name': input_data['org_car_name'],
            'carpan_model_name': input_data['org_model_name'],
            'carpan_model_detail': input_data['org_model_detail'],
            'provider' : provider,
            'brand_name': input_data['brand_name'][2],
            'car_name': input_data['car_name'][2],
            'model_name': input_data['model_name'][2],
            'car_price': input_data['car_price'][2],
            'deposit': input_data['deposit'][2],
            'downpayment': input_data['downpayment'][2] / input_data['car_price'][2],
            'period': input_data['period'][2],
            'output': output_value
            }
            records.append(record)
    if provider == 'im':
        for entry in output:
            input_data = entry['input'][provider]
            output_value = entry['output']
            record = {
            'carpan_brand_name': input_data['org_brand_name'],
            'carpan_car_name': input_data['org_car_name'],
            'carpan_model_name': input_data['org_model_name'],
            'carpan_model_detail': input_data['org_model_detail'],
            'provider' : provider,
            'brand_name': input_data['brand_name'][2],
            'car_name': input_data['car_name'][2],
            'model_name': input_data['model_name'][2],
            'car_price': input_data['car_price'][2],
            'deposit': input_data['deposit'][2] ,
            'downpayment': input_data['downpayment'][2] ,
            'period': input_data['period'][2],
            'output': output_value
            }
            records.append(record)
    if provider == 'lotte':
        for entry in output:
            input_data = entry['input'][provider]
            output_value = entry['output']
            record = {
            'carpan_brand_name': input_data['org_brand_name'],
            'carpan_car_name': input_data['org_car_name'],
            'carpan_model_name': input_data['org_model_name'],
            'carpan_model_detail': input_data['org_model_detail'],
            'provider' : provider,
            'brand_name': input_data['brand_name'][2],
            'car_name': input_data['car_name'][2],
            'model_name': input_data['model_name'][2],
            'car_price': input_data['car_price'][2],
            'deposit': input_data['deposit'][2] ,
            'downpayment': input_data['downpayment'][2] / input_data['car_price'][2],
            'period': input_data['period'][2],
            'output': output_value
            }
            records.append(record)
    if provider == 'nh':
        for entry in output:
            input_data = entry['input'][provider]
            output_value = entry['output']
            record = {
            'carpan_brand_name': input_data['org_brand_name'].strip().replace('\xa0', ' '),
            'carpan_car_name': input_data['org_car_name'].strip().replace('\xa0', ' '),
            'carpan_model_name': input_data['org_model_name'].strip().replace('\xa0', ' '),
            'carpan_model_detail': input_data['org_model_detail'].strip().replace('\xa0', ' '),
            'provider' : provider,
            'brand_name': input_data['brand_name'][2],
            'car_name': input_data['car_name'][2],
            'model_name': input_data['model_name'][2],
            'car_price': input_data['car_price'][2],
            'deposit': input_data['deposit'][2] ,
            'downpayment': input_data['downpayment'][2] / input_data['car_price'][2],
            'period': input_data['period'][2],
            'output': output_value
            }
            records.append(record)
    if provider == 'woori':
        for entry in output:
            input_data = entry['input'][provider]
            output_value = entry['output']
            record = {
            'carpan_brand_name': input_data['org_brand_name'],
            'carpan_car_name': input_data['org_car_name'],
            'carpan_model_name': input_data['org_model_name'],
            'carpan_model_detail': input_data['org_model_detail'],
            'provider' : provider,
            'brand_name': input_data['brand_name'][2],
            'car_name': input_data['car_name'][2],
            'model_name': input_data['model_name'][2],
            'car_price': input_data['car_price'][2],
            'deposit': input_data['deposit'][2] ,
            'downpayment': input_data['downpayment'][2] / input_data[provider]['car_price'][2],
            'period': input_data['period'][2],
            'output': output_value
            }
            records.append(record)

    
    
    return records, record

def convert_period(provider, x):
    if provider == 'bnk':
        if x == 1:
            return '60개월'
        if x == 2:
            return '48개월'
        if x == 3:
            return '36개월'
    if provider == 'sh':
        if x == 60:
            return '60개월'
        if x == 48:
            return '48개월'
        if x == 36:
            return '36개월'        
    if provider == 'mz':
        if x == 4:
            return '60개월'
        if x == 3:
            return '48개월'
        if x == 2:
            return '36개월'        
    if provider == 'lotte':
        if x == 60:
            return '60개월'
        if x == 48:
            return '48개월'
        if x == 36:
            return '36개월'        
    if provider == 'im':
        if x == 60:
            return '60개월'
        if x == 48:
            return '48개월'
        if x == 36:
            return '36개월'        
    if provider == 'woori':
        if x == 60:
            return '60개월'
        if x == 48:
            return '48개월'
        if x == 36:
            return '36개월'        
    if provider == 'nh':
        if x == 60:
            return '60개월'
        if x == 48:
            return '48개월'
        if x == 36:
            return '36개월'        



