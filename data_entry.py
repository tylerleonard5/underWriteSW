from pathlib import Path
import PySimpleGUI as sg
import pandas as pd
import openpyxl

wb = openpyxl.Workbook() 

# Add some color to the window
sg.theme('DarkTeal9')

layout = [
    [sg.Text('Please name Excel File:')],
    [sg.Text('Excel Filename', size=(15,1)), sg.InputText(key='Filename')],
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Floorplan', size=(15,1)), sg.InputText(key='FP')],
    [sg.Text('# of Units', size=(15,1)), sg.InputText(key='unitNum')],
    [sg.Text('Favorite Colour', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    [sg.Text('I speak', size=(15,1)),
                            sg.Checkbox('German', key='German'),
                            sg.Checkbox('Spanish', key='Spanish'),
                            sg.Checkbox('English', key='English')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Children')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')

    window['Children'].update(0)
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        excelName = (values['Filename'])
        excelName = excelName + '.xlsx'

        current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
        EXCEL_FILE = current_dir / excelName

        # Load the data if the file exists, if not, create a new DataFrame
        if EXCEL_FILE.exists():
            df = pd.read_excel(EXCEL_FILE)
        else:
            df = pd.DataFrame()

        wb.save(filename=excelName)

        keys_to_omit = ['Filename']  # Replace with keys you want to omit
        new_record_data = {key: values[key] for key in values if key not in keys_to_omit}

        new_record_data['test'] = "TESTING ADDING VALUE"

        new_record = pd.DataFrame(new_record_data, index=[0])

        df = pd.concat([df, new_record], ignore_index=False)
        for index, row in df.iterrows():
            for column, value in row.items():
                print(f"{column}: {value}")
                
        df.to_excel(excelName, index=False)  # This will create the file if it doesn't exist
        sg.popup('Data saved!')
        clear_input()
window.close()



# import PySimpleGUI as sg
# import pandas as pd

# # Initialize an empty DataFrame to store user input
# data = pd.DataFrame(columns=["Value"])

# # Define the PySimpleGUI layout
# layout = [
#     [sg.Text("Enter a value:"), sg.InputText(key="value"), sg.Button("Submit")],
#     [sg.Text("Calculate Percentage for Submission:"), sg.Input(key="submission_index"), sg.Button("Calculate Percentage")],
#     [sg.Button("Exit")],
# ]

# # Create the window
# window = sg.Window("Data Entry and Calculation", layout)

# while True:
#     event, values = window.read()

#     if event == sg.WIN_CLOSED or event == "Exit":
#         break

#     if event == "Submit":
#         try:
#             # Add the entered value to the DataFrame
#             value = float(values["value"])
#             data = data.append({"Value": value}, ignore_index=True)
#             sg.popup(f"Value {value} added.")
#             window["value"].update("")  # Clear the input field
#         except ValueError:
#             sg.popup("Please enter a valid numeric value.")

#     if event == "Calculate Percentage":
#         try:
#             submission_index = int(values["submission_index"])
#             if 0 <= submission_index < len(data):
#                 selected_value = data.iloc[submission_index]["Value"]
#                 total = data["Value"].sum()
#                 percentage = (selected_value / total) * 100
#                 sg.popup(f"Percentage for Submission {submission_index}: {percentage:.2f}%")
#             else:
#                 sg.popup("Invalid submission index.")
#         except ValueError:
#             sg.popup("Please enter a valid submission index.")

# window.close()