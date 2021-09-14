import PySimpleGUI as sg
import os
from process_xlsx import *
from send_email import *


def load_config(config_file=CONFIG_FILE_NAME):
    if os.path.exists(config_file):
        return retrive_configfile(config_file)
    else:
        return CONFIG_PARAMS

def update_user_inputs(input_values, config, config_file=CONFIG_FILE_NAME):
    '''
    input_values format is based on the InputTexts layout

    '''
    #TODO: need to validate input values
    config["reminder_days"]  = int(input_values[0])
    config["email"]          = input_values[1]
    config["expiry_title"]   = input_values[2]
    config["output_columns"] = input_values[3].split(',')

    update_configfile(config, config_file)


if __name__ == '__main__':

    config = load_config()

    sg.theme('DarkBlue16')
    input_xlsx = sg.popup_get_file('Enter input excel file')

    sg.theme('DarkBlack')
    layout = [
        [sg.Text(f'File: "{os.path.basename(input_xlsx)}"')],
        [sg.Text('Reminder days:', size=(35,2)),
            sg.InputText(config["reminder_days"])],
        [sg.Text('Email address:', size=(35,2)),
            sg.InputText(config["email"])],
        [sg.Text('Expiry date column title:', size=(35,2)),
            sg.InputText(config["expiry_title"])],
        [sg.Text('output column titles:', size=(35,2)),
            sg.InputText(", ".join(config["output_columns"]))],
        [sg.Button('Ok'), sg.Button('Cancel')]
    ]

    # Create the Window
    window = sg.Window('Expiry Notifications', layout)

    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        if event == 'Ok':
            update_user_inputs(values, config)
            if process_xlsx(input_xlsx, config) != None:
                #TODO: Display dialog for the error message (one or more columns are not found))
            send_email(config["email"], "Car park expiry remainder", "", OUTPUT_FILE_NAME)
            #TODO: DIsplay dialog for completion
            break
        elif event == 'Cancel': # if user closes window or clicks cancel
            break

    window.close()

