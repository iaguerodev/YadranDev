import os
import sys
import win32com.client
import pandas as pd
import pywintypes
import time

# Get the directory of the script // GET TT FINISH STATUS
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

# Convert column 'C' to string to avoid type issues
df['C'] = df['C'].astype(str)

# Keep track of encountered numbers and their corresponding results
encountered_numbers = set()
results_mapping = {}

# Iterate through each number in column B
for index, number in df['B'].items():
    # Skip NaN values
    if pd.isna(number):
        continue

    # Set transaction code directly
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nVI01"
    # Press Enter to navigate to VI01 transaction
    session.findById("wnd[0]").sendVKey(0)

    # Print the current number being processed
    print(f"Processing number in column B at index {index}: {number}")

    # If the number has been encountered before, copy the result from previous processing
    if number in encountered_numbers:
        result = results_mapping[number]
    else:
        try:
            # Set the number as text in SAP field VTTK-TKNUM
            session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").text = str(number)
            session.findById("wnd[0]").sendVKey(0)

            # Perform additional actions in SAP GUI
            session.findById("wnd[0]").resizeWorkingPane(129, 28, False)
            session.findById("wnd[0]/usr/tblSAPMV54ACRTL_ITEMS_VFKP/txtVFKP-POSTX[2,0]").caretPosition = 11
            session.findById("wnd[0]").sendVKey(2)
            session.findById("wnd[0]/tbar[1]/btn[18]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById("wnd[0]/usr/tabsTABSTRIP_ITEM/tabpPABR").select()
            session.findById("wnd[0]/usr/tabsTABSTRIP_ITEM/tabpPABR/ssubSCD_ITEM:SAPMV54A:0042/chkVFKPD-SLFREI").selected = True
            session.findById("wnd[0]/usr/tabsTABSTRIP_ITEM/tabpPABR/ssubSCD_ITEM:SAPMV54A:0042/chkVFKPD-SLFREI").setFocus()
            session.findById("wnd[0]").sendVKey(11)
            session.findById("wnd[0]/mbar/menu[0]/menu[2]").select()

            # Store the result in the mapping
            result = session.findById("wnd[0]/usr/ctxtVFKK-FKNUM").text
            results_mapping[number] = result
            # Update encountered numbers
            encountered_numbers.add(number)

        except pywintypes.com_error as e:
            print(f"Error during SAP interaction for number {number} at index {index}: {e}")
            sys.exit(f"Error in processing the number {number} related to column B in SAP interaction.")

    # Try to update the Excel file with the result
    try:
        df.at[index, 'C'] = str(result)
        # Verify the update was successful
        if df.at[index, 'C'] != str(result):
            raise ValueError(f"Failed to write the result {result} to column C at index {index}.")
        
        # Print what number is written in column C
        print(f"Written to column C at index {index}: {result}")

    except Exception as e:
        print(f"Error writing to column C for number {number} at index {index}: {e}")
        sys.exit(f"Error in processing the number {number} related to column B during Excel write operation.")

# Save the updated DataFrame back to the Excel file
try:
    df.to_excel(excel_file_path, index=False)
    print(f"Successfully saved the updated Excel file: {excel_file_path}")
except Exception as e:
    print(f"Error saving the Excel file: {e}")
    sys.exit("Error saving the updated Excel file.")
