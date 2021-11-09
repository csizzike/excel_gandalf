import PySimpleGUI as sg
import pandas as pd

sg.theme('Light Blue 2')

def move(listbox, old_index, new_index, limit):
    print(str(listbox.get_list_values))
    if not(0<=old_index<limit and 0<=new_index<limit):
        return False
    try:
        item = listbox.get(old_index)        # Get item by index
        listbox.delete(old_index)            # Remove item by index
        listbox.insert(new_index, item)      # insert item by index
        return True
    except:
        return False

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

layout = [
            [sg.Text('Select your template excel')],
            [sg.Text('Template excel', size=(8, 1)), sg.Input(), sg.FileBrowse()],
            [sg.Button("Submit")],
            [sg.Listbox(values=[], size=(30,6), key="-ITEM-")],
            [sg.Button("Move item up")],
            [sg.Button("Move item down")],
            [sg.Text('Current column order ', size = (50,1), key='-LIST-')]
          
          ]
window = sg.Window('Excel Gandalf', layout, icon='icon\gandalf.ico')

while True:  # Event Loop
    event, values = window.read()
    #df1 = df1[["Hibá", "Meresi", "Cica", "Kutya"]]
    #df1.to_excel("test_reordered.xlsx", index=False)
    if event == 'Submit':
        df1 = pd.read_excel(values['Browse'])
        header_list = list(df1)
        print(header_list)
        listbox = window['-ITEM-']
        done = False
        listbox.Update(values=header_list)

    elif event == "Move item up" and not done:
        print("Választott index " + str(listbox.get_indexes()))
        index = listbox.get_indexes()[0]
        moveUp(header_list, index)
        if index == 0:
            listbox.update(values=header_list, set_to_index=index)
        else:
            listbox.update(values=header_list, set_to_index = index-1)
        final_list = list(listbox.get_list_values()) # Get final values in sg.litbox
        window['-LIST-'].update(value="Current order "+repr(final_list))

    elif event == "Move item down" and not done:
        print("Választott index " + str(listbox.get_indexes()))
        index = listbox.get_indexes()[0]
        moveDown(header_list, index)
        if index == len(header_list)-1:
            listbox.update(values=header_list, set_to_index=index)
        else:
            listbox.update(values=header_list, set_to_index = index+1)
        final_list = list(listbox.get_list_values()) # Get final values in sg.litbox
        window['-LIST-'].update(value="Current order "+repr(final_list))

    if event == sg.WIN_CLOSED or event == None or event == 'Exit':
        break
    if event == 'Show':
        # Update the "output" text element to be the value of "input" element
        window['-OUTPUT-'].update(values['-IN-'])
    if event == ' Cancel':
        break

window.close()