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
    config["output_columns"] = [x.strip() for x in input_values[3].split(',')]

    update_configfile(config, config_file)


if __name__ == '__main__':

    config = load_config()

    sg.set_options(font = 'Courier 20')
    sg.theme('DarkBlue16')
    input_xlsx = sg.popup_get_file('Enter input excel file', file_types = (("excel file", "*.xlsx"),))
    if input_xlsx == None:
        exit(1)

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
    return_value = "cancel"
    while True:
        event, values = window.read()
        if event == 'Ok':
            update_user_inputs(values, config)
            err_msg = process_xlsx(input_xlsx, config)
            if err_msg:
                return_value = "error"
            else:
                send_email(config["email"], "Car park expiry remainder", "", OUTPUT_FILE_NAME)
                return_value = "pass"
            break
        elif event == 'Cancel': # if user closes window or clicks cancel
            break

    window.close()

    if return_value == "pass":
        sg.popup("\n  email sent!  \n", title="email sent")
    elif return_value == "error":
        sg.popup_error(f"\n {err_msg}   \n", title="Error")
