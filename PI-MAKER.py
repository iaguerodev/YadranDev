import win32com.client
import pyautogui
import pyperclip
import time
import pygetwindow as gw
import openpyxl

# Function to check if the active window's title contains a specific substring
def is_window_title_containing(substring):
    active_window = gw.getActiveWindow()
    if substring in active_window.title:
        return True
    return False

# Function to refocus on sap window
def focus_sap_window(window_title):
    sap_window = gw.getWindowsWithTitle(window_title)
    if sap_window:
        sap_window[0].activate()

# Load Excel data
excel_file = "LIST.xlsx"
workbook = openpyxl.load_workbook(excel_file, data_only=True)
worksheet = workbook.active


# Initialize SAP GUI Scripting
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# Check if a connection to SAP GUI is established
if session.Children.Count < 0:
    session = application.CreateObject("Session")

# Set the working pane dimensions
session.findById("wnd[0]").resizeWorkingPane(95, 26, False)

# Access the relevant nodes in SAP GUI
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00007"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("F00007")



# Iterate through rows in Excel
for row in worksheet.iter_rows(min_row=2, values_only=True):
    cv_value = row[0]
    client_name = row[1]

    # Set text in a field
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = f"{cv_value}"

    # Select a menu item
    session.findById("wnd[0]/mbar/menu[0]/menu[5]").select()

    # Simulate keyboard input to press a key (in this case, F86)
    session.findById("wnd[1]").sendVKey(86)

    # Set a checkbox to selected
    session.findById("wnd[2]/usr/chkSSFPP-TDIMMED").selected = True

    # Enter text in a field
    session.findById("wnd[2]/usr/ctxtSSFPP-TDDEST").text = "loca"

    # Set focus to another control
    session.findById("wnd[2]/usr/chkSSFPP-TDIMMED").setFocus()

    # Simulate keyboard input to press a key (in this case, F86)
    session.findById("wnd[2]").sendVKey(86)

    # Wait for 2 seconds (you can adjust this as needed)
    time.sleep(2)

    # Check if the active window's title contains "Guardar impresión como"
    if is_window_title_containing("Guardar impresión como"):
        pyautogui.typewrite(f"PI {cv_value} {client_name}")
        print(f"Successfully typed: PI {cv_value} {client_name}")
    else:
        print("Skipped typing because the window title doesn't match.")

    # Debugging: Capture the text from the clipboard
    copied_text = pyperclip.paste()
    if copied_text == f"PI {cv_value} {client_name}":
        print(f"Successfully typed: {copied_text}")
    else:
        print(f"Failed to type: PI {cv_value} {client_name}")

    # Press enter to save the document
    pyautogui.press('enter')

    # Specify the SAP window title to focus on
    sap_window_title = "Visualizar documentos de ventas"

    # Call the function to focus on the SAP window
    focus_sap_window(sap_window_title)

