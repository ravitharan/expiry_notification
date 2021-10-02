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
    input_xlsx                  = input_values[0]
    config["reminder_days"]     = int(input_values[1])
    config["email"]             = input_values[2]
    config["expiry_title"]      = input_values[3]
    config["output_columns"]    = [x.strip() for x in input_values[4].split(',')]

    return input_xlsx


if __name__ == '__main__':

    config_file = CONFIG_FILE_NAME

    config = load_config()

    sg.set_options(font = 'Courier 20')

    sg.theme('DarkBlue16')
    layout = [
        [sg.Text('Select input excel File: ', size=(35,2)),
            sg.InputText(), sg.FileBrowse(file_types = (("excel file", "*.xlsx"),))],
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
    window = sg.Window('Expiry Notifications v1.0', layout)

    # Event Loop to process "events" and get the "values" of the inputs
    return_value = "cancel"
    while True:
        event, values = window.read()
        if event == 'Ok':
            input_xlsx = update_user_inputs(values, config)
            err_msg = process_xlsx(input_xlsx, config)
            if err_msg:
                return_value = "error"
            else:
                update_configfile(config, config_file)
                send_email(config["email"], "Car park expiry remainder", "", OUTPUT_FILE_NAME)
                return_value = "pass"
            break
        elif event == 'Cancel' or event == sg.WIN_CLOSED: # if user closes window or clicks cancel
            break

    window.close()

    if return_value == "pass":
        sg.popup("\n  email sent!  \n", title="email sent")
    elif return_value == "error":
        sg.popup_error(f"\n {err_msg}   \n", title="Error")
