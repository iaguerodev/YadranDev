import os
import win32com.client
import pandas as pd
import pywintypes
import time

# Get the directory of the script ### TT WITH DATES
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


# Set transaction code directly
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVT02N"

# Press Enter to navigate to VA03 transaction
session.findById("wnd[0]").sendVKey(0)

# Load Excel file
df = pd.read_excel(excel_file_path)

# Keep track of encountered numbers
encountered_numbers = set()

# Iterate through each number in column B
for index, number in df['B'].items():
    # If the number has been encountered before, skip it
    if number in encountered_numbers:
        continue

    # Set the number as text in SAP
    session.findById("wnd[0]/usr/ctxtVTTK-TKNUM").text = str(number)

    # Press Enter
    session.findById("wnd[0]").sendVKey(0)

    # Set text in various fields (DATE)
    preg_field = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/ctxtVTTK-DPREG")
    preg_field.setFocus()
    preg_field.caretPosition = 10
    copied_text = preg_field.text

    fields = [
        "VTTK-DTDIS",
        "VTTK-DAREG",
        "VTTK-DALBG",
        "VTTK-DALEN",
        "VTTK-DTABF",
        "VTTK-DATBG",
        "VTTK-DATEN"
    ]

    for field in fields:
        session.findById(f"wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/ctxt{field}").text = copied_text

    # Add the number to the set of encountered numbers
    encountered_numbers.add(number)

    # Set focus and caret position again
    preg_field.setFocus()
    preg_field.caretPosition = 10

    time.sleep(2)

    # Send keys
    try:
        session.findById("wnd[0]/mbar/menu[4]/menu[7]/menu[2]").select()
    except Exception as e:
        print("Menu not found, executing alternative code...")
        # Execute alternative code
        session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/btn*RV56A-ICON_STTEN").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
    else:
        # Continue with the following actions if the menu is found
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()

# Check if VTTK-DATEN is empty at the end
vttk_daten_text = session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/ctxtVTTK-DATEN").text
if not vttk_daten_text:
    # Interaction with VTTK-DATEN field only if it's empty
    session.findById("wnd[0]/usr/tabsHEADER_TABSTRIP2/tabpTABS_OV_DE/ssubG_HEADER_SUBSCREEN2:SAPMV56A:1025/btn*RV56A-ICON_STTEN").press()
    print("VTTK-DATEN field was empty. Interaction performed.")
else:
    print("VTTK-DATEN field already contains data. Skipped interaction.")