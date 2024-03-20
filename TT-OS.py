import os
import win32com.client
import pandas as pd
import time

# Get the directory of the script
script_dir = os.path.dirname(os.path.abspath(__file__))
excel_file_path = os.path.join(script_dir, 'TT.xlsx')

# Initialize SAP GUI Scripting
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# Check if a connection to SAP GUI is established
if session.Children.Count < 0:
    session = application.CreateObject("Session")

# Load Excel file
df = pd.read_excel(excel_file_path)

# Set transaction code directly
transaction_code = "/nva03"

# Iterate through each number in column A
for index, number in enumerate(df['A']):
    print(f"Processing number {number} at index {index}")

    # Set transaction code directly to go back to the initial screen
    session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code

    # Press Enter to navigate to the specified transaction
    session.findById("wnd[0]").sendVKey(0)

    # Set the number as text in SAP
    try:
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = str(number)
        print(f"Number {number} set in SAP.")
    except:
        print(f"Control not found for index {index} and number {number}")
        continue

    # Press Enter
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/mbar/menu[4]/menu[7]/menu[2]").select()

    # List of controls to check
    controls = ["wnd[0]/usr/lbl[6,5]", "wnd[0]/usr/lbl[6,4]"]  # Update with actual control IDs

    found_number = None

    # Loop through the controls
    for control_id in controls:
        try:
            control = session.findById(control_id)
            control_text = control.Text
            print(f"Control ID: {control_id}, Text: {control_text}")
            if control_text.startswith('1'):
                found_number = control_text
                break
        except:
            print(f"Control not found for index {index} and number {number}")
            continue

    # If a number starting with '1' is found, save it in column B of the Excel file
    if found_number:
        df.loc[index, 'B'] = str(found_number)
        print(f"Found number {found_number} at index {index}.")

    # Add a small delay between iterations
    time.sleep(1)

# Save the updated DataFrame to Excel
df.to_excel(excel_file_path, index=False)