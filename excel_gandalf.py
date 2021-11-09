import PySimpleGUI as sg
import pandas as pd

sg.theme('Light Blue 2')

final_list = []

def moveUp(header_list, selected_index):
    if 0 < selected_index <= len(header_list)-1:
        item = header_list[selected_index]
        header_list.pop(selected_index)
        header_list.insert(selected_index - 1, item)
    if selected_index <= 0 or selected_index >= len(header_list):
        return

def moveDown(header_list, selected_index):
    if 0 <= selected_index < len(header_list)-1:
        item = header_list[selected_index]
        header_list.pop(selected_index)
        header_list.insert(selected_index + 1, item)
    if selected_index >= len(header_list):
        return

def copy_and_save_excels(pathfrom, pathto):
    for file in pathfrom.split(";"):
        export_path = ''
        df = pd.read_excel(file)
        df = df[final_list]
        file_name = file.split('.')[-2].split('/')[-1]
        export_path = pathto + "/" + file_name + "_reordered.xlsx"
        df.to_excel(export_path, index=False)

def open_next_window():
    layout = [
        [sg.Text('Apply on', size=(10, 1)),
         sg.Input(), sg.FileBrowse(button_text='Select')],
         [sg.Text('Save to', size=(10, 1)),
         sg.Input(), sg.FolderBrowse(button_text='Browse'),
         sg.Button("Save reordered excels")]
    ]
    window = sg.Window("Second Window", layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == "Save reordered excels":
            copy_and_save_excels(values["Select"], values["Browse"])
            sg.Popup('The new copies have been saved to: ' + values['Browse'])
            window.close()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
        
    window.close()

layout = [
            [sg.Text('Select your template excel')],
            [sg.Text('Template', size=(8, 1)), sg.Input(), sg.FileBrowse(), sg.Button("Load excel")],
            [sg.Listbox(values=[], size=(30,15), key="-ITEM-", expand_x=True, expand_y=True, highlight_background_color='#6c82b8')],
            [sg.Button("Move item up"), sg.Button("Move item down")],
            [sg.Text('Current order ', key='-LIST-', expand_x=True)],
            [sg.Button("Next")]
          ]
window = sg.Window('Excel Gandalf', layout, icon='icon\gandalf.ico')

while True:  # Event Loop
    event, values = window.read()
    if event == 'Load excel':
        if values['Browse'] != '' and values['Browse'].split('.')[-1] == 'xlsx':
            df1 = pd.read_excel(values['Browse'])
            header_list = list(df1)
            listbox = window['-ITEM-']
            done = False
            listbox.Update(values=header_list)
            window['-LIST-'].update(value="Current order "+repr(header_list))
        else:
            sg.Popup('Select an excel file to continue')

    elif event == "Move item up" and not done:
        index = listbox.get_indexes()[0]
        moveUp(header_list, index)
        if index == 0:
            listbox.update(values=header_list, set_to_index=index)
        else:
            listbox.update(values=header_list, set_to_index = index-1)
        final_list = list(listbox.get_list_values()) # Get final values in sg.litbox
        window['-LIST-'].update(value="Current order "+repr(final_list))

    elif event == "Move item down" and not done:
        index = listbox.get_indexes()[0]
        moveDown(header_list, index)
        if index == len(header_list)-1:
            listbox.update(values=header_list, set_to_index=index)
        else:
            listbox.update(values=header_list, set_to_index = index+1)
        final_list = list(listbox.get_list_values()) # Get final values in sg.litbox
        window['-LIST-'].update(value="Current order "+repr(final_list))

    elif event == "Next" and not done:
        open_next_window()

    if event == sg.WIN_CLOSED or event == None or event == 'Exit':
        break

window.close()
