import json
import openpyxl

CONFIG_FILE_NAME = ".config.json"

# Get following parameters from GUI inputs.
config_params = {
    "receiver_email" : "user@email.com",
    "notification_days" : 5,
    "expiry_title" : "Valid till",
    "output_columns" : ["Name", "Access card no", "Vehicle no"],
}

def update_configfile(file_name, config_params):
    with open(file_name, 'w') as fp:
        json.dump(config_params, fp)

def retrive_configfile(file_name):
    with open(file_name) as fp:
        params = json.load(fp)
    return params

def parse_xlsx_header(xlsx_file, cfg_params):
    '''
    Parse xlsx file and returen column number for expiry date and other output
    columns. Also returning data start row number

    '''
    data_locations = {
        "work_sheet" : None,
        "data_start_row": None,
        "columns": None,
    }
    # First interested column is expiry date
    data_locations["columns"] = [None]
    # Subsequent columns are requested bu user
    data_locations["columns"].extend([None for x in cfg_params["output_columns"]])
    num_locations = len(data_locations["columns"])

    wb = openpyxl.load_workbook(xlsx_file)
    ws = wb.active
    data_locations["work_sheet"] = ws

    for row in range(1, ws.max_row+1):
        for column in range(1, ws.max_column+1):
            if (ws.cell(row, column).value == cfg_params["expiry_title"]):
                data_locations["columns"][0] = column
                data_locations["data_start_row"] = row + 1
                num_locations -= 1
            for column_label in cfg_params["output_columns"]:
                if ws.cell(row, column).value == column_label:
                    index = cfg_params["output_columns"].index(column_label)
                    data_locations["columns"][1+index] = column
                    num_locations -= 1
            if (num_locations == 0):
                break
        if (num_locations == 0):
            break

    return data_locations

def get_xlsx_data(xlsx_locations):
    '''
    Return expiry date and other output columns in a list from the excel file
    '''
    xlsx_data = []
    ws = xlsx_locations["work_sheet"]

    for row in range(xlsx_locations["data_start_row"], ws.max_row+1):
        row_data = []
        for column in xlsx_locations["columns"]:
            value = ws.cell(row, column).value
            row_data.append(value)
        if row_data[0] != None:
            xlsx_data.append(row_data)

    return xlsx_data

def filter_data(xlsx_data, num_expiry_days):
    '''
    Filter xlsx_data within num_expiry_days. Sort the filtered data 
    '''

def write_xlsx_file(xlsx_data):
    '''
    Write filtered data into output xlsx file
    '''
if __name__ == '__main__':
    locations = parse_xlsx_header('DMM Access cards details.xlsx', config_params)
    data = get_xlsx_data(locations)
    for item in data:
        print(item)
