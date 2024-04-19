import keyboard
import pyautogui
import pygetwindow as gw
import time
import pandas as pd
import os
import pyperclip  # Import the pyperclip library for clipboard operations
import xlwings as xw
import unicodedata

# Set the default location for the Excel file (relative path)
DEFAULT_EXCEL_FILE = "CVAWB.xlsx"

# Try to find the Excel file in the current working directory
excel_file_path = os.path.join(os.getcwd(), DEFAULT_EXCEL_FILE)

if not os.path.exists(excel_file_path):
    print(f"Excel file not found in the current working directory. Trying default location...")
    
    # If not found, try the default location
    default_excel_path = os.path.join(os.path.dirname(__file__), DEFAULT_EXCEL_FILE)
    
    if os.path.exists(default_excel_path):
        excel_file_path = default_excel_path
    else:
        print(f"Excel file not found in the default location.")
    print(f"Excel file path: {excel_file_path}")

# Constants
EXCEL_FILE_NAME = "CVAWB.xlsx"
INITIAL_SAP_WINDOW_TITLE = 'SAP Easy Access'
SECOND_SCREEN_TITLE = 'Crear transporte: Acceso'
THIRD_SCREEN_TITLE = 'Tran.Aéreo Exp $0001 Crear: Resumen'
EXCEL_SHEET_NAME = "Hoja1"  # Replace with the actual sheet name

AIRLINE_MAPPING = {
    'AMERICAN': 'A0001',
    'AARG': 'A0002',
    'AEROMEXICO': 'A0003',
    'AIR CANADA': 'A0004',
    'AIR CHINA': 'A0005',
    'AIR FRANCE': 'A0006',
    'ALFALOGIST': 'A0007',
    'ALITALIA': 'A0008',
    'ANZL': 'A0009',
    'ASIANA': 'A0010',
    'ATLAS': 'A0011',
    'AV-AGI': 'A0012',
    'AV-ALF': 'A0013',
    'AVIANCA': 'A0014',
    'AZTEC': 'A0015',
    'AZUL': 'A0016',
    'BRITISH': 'B0001',
    'CENTURION': 'C0001',
    'CHARTER': 'C0002',
    'COPA': 'C0003',
    'CP': 'C0004',
    'CARGOLUX': 'C0005',
    'CATHAY': 'C0006',
    'DELTA': 'D0001',
    'DHL': 'D0002',
    'EMIRATES': 'E0001',
    'EAC': 'E0002',
    'ETHIOPIAN': 'E0003',
    'ELAL': 'E0004',
    'GOL': 'G0001',
    'IBERIA': 'I0001',
    'JETSMART': 'J0001',
    'KOREAN': 'K0001',
    'KALITAIR': 'K0002',
    'KLM': 'K0003',
    'LATAM': 'L0001',
    'LATAMJFK': 'L0002',
    'LATAMMIA': 'L0003',
    'LUFTHANSA': 'L0004',
    'NULL': 'N0001',
    'OCEANCARGO': 'O0001',
    'QANTAS': 'Q0001',
    'QATAR': 'Q0002',
    'SKY': 'S0001',
    'SAA': 'S0002',
    'SW': 'S0003',
    'TURKISH': 'T0001',
    'UNITED': 'U0001',
    'UPS': 'U0002'
}

# Destination code mappings
DEST_CODE_MAPPING = {
    'TLV': 'ZCA064',
    'HND': 'ZCA343',
    'KIX': 'ZCA074',
    'NRT': 'ZCA075',
    'LAX': 'ZCA122',
    'MIA': 'ZCA125',
    'LCA': 'ZCA281',
    'FUK': 'ZCA358',

    #CHINA
    'CKG': 'ZCA023',
    'PVG': 'ZCA032',
    'CAN': 'ZCA026',
    'PEK': 'ZCA021',
    'CTU': 'ZCA022',
    'XMN': 'ZCA035'

}

# Function to find and activate a SAP window by title
def activate_sap_window(window_title):
    while True:
        sap_window = gw.getWindowsWithTitle(window_title)
        if sap_window:
            sap_window[0].activate()
            return sap_window[0]
        else:
            print(f"SAP window '{window_title}' not found. Retrying...")
            time.sleep(2)

# Function to get the FF code based on FF name
def get_ff_code(ff_name):
    ff_name = ff_name.upper()  # Convert to uppercase for consistent matching
    ff_codes = {
        'ALFA': '76272345-K',
        'AGILITY': '76408000-9',
        'ANDES': '76788050-2',
        'A&A': '78171100-4'
        # Add more mappings as needed
    }
    return ff_codes.get(ff_name, None)

# Function to get the SAP code based on DEST value
def get_sap_code(dest):
    dest = dest.upper()  # Convert to uppercase for consistent matching
    return DEST_CODE_MAPPING.get(dest, None)

# Function to check if the specified text is present in the SAP field
def check_text_in_sap_field(text):
    time.sleep(2)  # Wait for the field to be ready
    field_text = pyautogui.typewrite(text)
    return field_text is not None

# Function to copy text to the clipboard
def copy_to_clipboard(text):
    pyperclip.copy(text)


# Function to automate the SAP input
def automate_sap_input():
    # Simulate typing into the command field (top left of the SAP window)
    pyautogui.hotkey('ctrl', '/')
    pyautogui.write("/nVT01N")
    pyautogui.press('enter')

    # Wait for a moment to give time for SAP to process the input
    time.sleep(3)


    # Ensure the SAP window is still active
    sap_window = gw.getActiveWindow()
    if not sap_window or sap_window.title != SECOND_SCREEN_TITLE:
        print(f"SAP window '{SECOND_SCREEN_TITLE}' lost focus or not found. Exiting...")
        return False

    # Simulate typing '2030' into the selected field
    pyautogui.write('2030')

    # Press 'Tab' to move to the next field
    pyautogui.press('tab')

    # Simulate typing 'ZT03' into the selected field
    pyautogui.write('Tran.Aéreo')

    time.sleep(2)

    # Press 'Enter' to submit the input
    pyautogui.press('enter')

    # Wait for a moment to give time for SAP to process the input
    time.sleep(2)

    # Ensure the SAP window is still active
    sap_window = gw.getActiveWindow()
    if not sap_window or sap_window.title != THIRD_SCREEN_TITLE:
        print(f"SAP window '{THIRD_SCREEN_TITLE}' not found. Exiting...")
        return False
    return True


# Function to convert the date format
def convert_date_format(date_str, year):
    try:
        if pd.isna(date_str):  # Check for empty cells
            return ""  # Return an empty string for empty cells

        # Extract day, month, and numerical day from the date string
        parts = date_str.split()
        numerical_day = parts[-2]
        month = parts[-1].upper()  # Convert the month representation to uppercase for consistency

        # Create a dictionary to map month names to numerical values (assuming English and Spanish month names)
        month_mapping = {
            'JAN': '01',
            'FEB': '02',
            'MAR': '03',
            'APR': '04',
            'MAY': '05',
            'JUN': '06',
            'JUL': '07',
            'AUG': '08',
            'SEP': '09',
            'OCT': '10',
            'NOV': '11',
            'DEC': '12',
            'JANUARY': '01',
            'FEBRUARY': '02',
            'MARCH': '03',
            'APRIL': '04',
            'MAY': '05',
            'JUNE': '06',
            'JULY': '07',
            'AUGUST': '08',
            'SEPTEMBER': '09',
            'OCTOBER': '10',
            'NOVEMBER': '11',
            'DECEMBER': '12',
            'ENERO': '01',
            'FEBRERO': '02',
            'MARZO': '03',
            'ABRIL': '04',
            'MAYO': '05',
            'JUNIO': '06',
            'JULIO': '07',
            'AGOSTO': '08',
            'SEPTIEMBRE': '09',
            'OCTUBRE': '10',
            'NOVIEMBRE': '11',
            'DICIEMBRE': '12',
        }

        # Check if the month abbreviation exists in the mapping
        if month in month_mapping:
            numerical_month = month_mapping[month]
        else:
            raise ValueError(f"Unknown month: {month}")

        # Combine the components to create the new date string in dd.mm.yyyy format
        new_date_str = f"{numerical_day}.{numerical_month}.{year}"

        return new_date_str
    except Exception as e:
        return str(e)

        # Main script
def main():
    try:
        # Get the directory where the Python script is located
        script_directory = os.path.dirname(os.path.abspath(__file__))

        # Construct the full path to the Excel file
        excel_file_path = os.path.join(script_directory, EXCEL_FILE_NAME)

        if not os.path.exists(excel_file_path):
            raise FileNotFoundError(f"Excel file '{EXCEL_FILE_NAME}' not found in the script directory.")

        # Load the Excel file
        df = pd.read_excel(excel_file_path, sheet_name=EXCEL_SHEET_NAME)

        # Get the content of the 'AÑO' cell in the first row
        ano_cell_content = str(df.at[0, 'AÑO'])  # Convert to a string

        # Debugging line to print the content of the 'AÑO' cell
        print(f"Content of 'AÑO' cell: '{ano_cell_content}'")

        # Remove non-breaking space characters (ASCII 160) and strip leading/trailing whitespaces
        ano_cell_content = unicodedata.normalize("NFKD", ano_cell_content).strip()

        # Check if the content of the 'AÑO' cell starts with '202'
        if isinstance(ano_cell_content, str) and ano_cell_content.startswith('202'):
            year = int(float(ano_cell_content))  # Corrected conversion to handle '2023.0'
        else:
            raise ValueError("Invalid year value in cell 'AÑO'. It should start with '202'.")

        # Debugging line to print the converted year
        print(f"Year after conversion: {year}")

        # Use the 'convert_date_format' function with the 'year' value
        date_result = convert_date_format('Your Date String', year)
        # 'Your Date String' should be replaced with the actual date string you want to convert

        # Check if the 'Itin.trans' column is not empty before applying date conversion
        if 'Itin.trans' in df.columns and not df['Itin.trans'].empty:
            # Convert the date column 'Itin.trans' and store the results in 'Date conv' column 'Date conv'
            df['Date conv'] = df['Itin.trans'].apply(lambda x: convert_date_format(x, year))

            # Save the updated DataFrame back to the Excel file
            df.to_excel(excel_file_path, sheet_name=EXCEL_SHEET_NAME, index=False)

            print("Date conversion completed.")
        else:
            print("No 'Itin.trans' column found or it is empty. No date conversion performed.")

        df = pd.read_excel(excel_file_path)

        # Get the unique CV numbers from the 'CV' column in the Excel sheet
        cv_numbers = df['CV'].unique()

        if len(cv_numbers) == 0:
            print("No CV numbers found in the Excel sheet. Exiting...")
            return

        # Iterate over the unique CV numbers
        for cv_number in cv_numbers:
            # Filter the DataFrame for the current CV number
            cv_df = df[df['CV'] == cv_number]

            # Attempt to activate the SAP Easy Access window
            sap_window = activate_sap_window(INITIAL_SAP_WINDOW_TITLE)
            if not sap_window:
                print(f"Failed to find the '{INITIAL_SAP_WINDOW_TITLE}' window. Exiting...")

            # Automate SAP input
            if not automate_sap_input():
                return

            print(f"Successfully navigated to the '{THIRD_SCREEN_TITLE}' window.")

            # Read the value in column 'FF NAME' from the filtered Excel DataFrame
            ff_name = cv_df.at[cv_df.index[0], 'FF NAME']

            # Get the corresponding FF code based on FF name
            ff_code = get_ff_code(ff_name)

            if ff_code:
                # Simulate typing the FF code into the SAP field
                pyautogui.write(ff_code)
                pyautogui.press('enter')  # Press 'Enter' after typing the FF code

                # Wait for 1 second
                time.sleep(2)

                # Press the 'Tab' key six times
                pyautogui.press('tab', presses=2)  # Press 'Tab' six times

                # Read the value in column 'DEST' from the filtered Excel DataFrame
                dest = cv_df.at[cv_df.index[0], 'DEST']

                # Get the corresponding SAP code based on DEST value
                sap_code = get_sap_code(dest)

                if sap_code:
                    # Simulate typing the SAP code into the SAP field
                    pyautogui.write(sap_code)
                    pyautogui.press('enter')  # Press 'Enter' after typing the SAP code

                    # Wait for 1 second
                    time.sleep(2)

                    # Press the 'Tab' key six times
                    pyautogui.press('tab', presses=6)  # Press 'Tab' six times

                    # Read the AWB from the filtered Excel DataFrame (assuming it's in column B with a header 'AWB')
                    awb = cv_df.at[cv_df.index[0], 'AWB']

                    # Simulate typing the AWB into the SAP field
                    pyautogui.write(awb)

                    # Press 'Tab' 3 times to navigate to the appropriate SAP field
                    for _ in range(3):
                        pyautogui.press('tab')

                    # Wait for 1 second
                    time.sleep(2)

                    # Copy the 'FLIGHT' value from the filtered Excel DataFrame to the clipboard
                    flight_value = cv_df.at[cv_df.index[0], 'FLIGHT']
                    copy_to_clipboard(flight_value)

                    # Use Ctrl+V to paste the copied 'FLIGHT' value into the SAP field
                    pyautogui.hotkey('ctrl', 'v')

                    # Press 'Tab' 5 times to navigate to the appropriate SAP field
                    for _ in range(5):
                        pyautogui.press('tab')

                    # Read the value from the 'Date conv' column in Excel (header "Date conv")
                    date_from_column_date_conv = cv_df.at[cv_df.index[0], 'Date conv']

                    # Copy the date value to the clipboard
                    copy_to_clipboard(date_from_column_date_conv)

                    # Use Ctrl+V to paste the date value into the SAP field
                    pyautogui.hotkey('ctrl', 'v')

                    # After pasting the 'FLIGHT' value, press 'Tab' 5 times to navigate to the appropriate SAP field
                    for _ in range(5):
                        pyautogui.press('tab')

                    # Read the value from the 'Date conv' column in Excel (header "Date conv")
                    date_from_column_date_conv = cv_df.at[cv_df.index[0], 'Date conv']

                    # Copy the date value to the clipboard
                    copy_to_clipboard(date_from_column_date_conv)

                    # Use Ctrl+V to paste the date value into the SAP field
                    pyautogui.hotkey('ctrl', 'v')

                    # Repeat the above steps 4 more times
                    for _ in range(4):
                        # Move down again
                        pyautogui.press('down')

                        # Use Ctrl+V to paste the date value into the SAP field
                        pyautogui.hotkey('ctrl', 'v')

                    # After pasting the 'FLIGHT' value, press 'Tab' 5 times to navigate to the appropriate SAP field
                    pyautogui.hotkey('ctrl', 'down')

                    # Wait for a moment to give time for SAP to process the input
                    time.sleep(1)

                    # After
                    pyautogui.hotkey('ctrl', 'left')

                    # After
                    pyautogui.hotkey('ctrl', 'left')

                    # After
                    pyautogui.hotkey('ctrl', 'left')

                    # After
                    pyautogui.hotkey('right')

                    # After
                    pyautogui.hotkey('enter')

                    time.sleep(1)

                    for _ in range(3):
                        pyautogui.press('tab')

                    # Copy the 'CV' value from the filtered Excel DataFrame (assuming it's in column A with a header 'CV')
                    cv_value = cv_df.at[cv_df.index[0], 'CV']

                    # Check if cv_value is not empty
                    if cv_value:
                        try:
                            # Check if the CV value is already an integer
                            if isinstance(cv_value, int):
                                cv_to_copy = cv_value
                            else:
                                # Try to convert the CV value to an integer (skipping any decimal parts)
                                cv_to_copy = int(float(cv_value))

                            # Clear the clipboard
                            pyperclip.copy('')

                            # Copy the integer CV value to the clipboard
                            copy_to_clipboard(cv_to_copy)

                            # Simulate a keyboard shortcut to paste (e.g., Ctrl+V)
                            keyboard.press_and_release('ctrl + v')

                            # Wait for a moment to ensure the value is pasted
                            time.sleep(1)

                            # Check if the pasted CV value matches the expected 'CV' value
                            pasted_cv_value = pyperclip.paste()

                            if str(pasted_cv_value).strip() == str(cv_to_copy).strip():
                                print(f"Successfully pasted 'CV' value: {cv_to_copy}")
                            else:
                                print(f"Pasted 'CV' value does not match: Expected '{cv_to_copy}', Actual '{pasted_cv_value}'")
                        except Exception as e:
                            print(f"Error pasting 'CV' value: {str(e)}")
                    else:
                        print("CV value is empty or not found in the Excel sheet.")

                    # After
                    for _ in range(1):
                        pyautogui.press('tab')

                    pyautogui.write("USD")

                    # Change to under tabs
                    for _ in range(3):
                        pyautogui.press('tab')

                    # Change to under tabs to datos adicionales
                    for _ in range(8):
                        pyautogui.press('right')

                    time.sleep(1)
                    pyautogui.press('enter')
                    time.sleep(1)
                    # Datos adicionales
                    pyautogui.hotkey('tab')

                    # Code for AIRLINE

                    # Read the airline name from the filtered Excel DataFrame (assuming it's in column E with a header 'AIRLINE')
                    airline_name = cv_df.at[cv_df.index[0], 'AIRLINE']

                    # Get the corresponding SAP code based on the airline name
                    sap_code = AIRLINE_MAPPING.get(airline_name.upper())

                    if sap_code:
                        # Simulate typing the SAP code into the SAP field
                        pyautogui.write(sap_code)
                        pyautogui.press('enter')  # Press 'Enter' after typing the SAP code
                    else:
                        print(f"Airline name '{airline_name}' not found in the dictionary. Unable to determine SAP code.")

                    # Save Air Transport
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 's')
                    time.sleep(1)
                    pyautogui.press('enter')
                    pyautogui.press('enter')

                    time.sleep(1)
                    pyautogui.press('F10')
                    pyautogui.press('v')
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.hotkey('ctrl', 'c')

                    # Check if the clipboard contains a value
                    clipboard_value = pyperclip.paste()
                    if clipboard_value:
                        try:
                            # Initialize an Excel app with xlwings
                            app = xw.App(visible=False)
                            try:
                                # Open the Excel file without displaying it
                                wb = app.books.open(excel_file_path)

                                # Check if the Excel file was opened successfully
                                if not wb:
                                    raise Exception(f"Failed to open Excel file '{excel_file_path}'")

                                # Select the worksheet by name
                                ws = None
                                for sheet in wb.sheets:
                                    if sheet.name == EXCEL_SHEET_NAME:
                                        ws = sheet
                                        break

                                # Check if the worksheet was found
                                if not ws:
                                    raise Exception(f"Worksheet '{EXCEL_SHEET_NAME}' not found in Excel file")

                                # Find the next empty row in column 'I' under the header 'N° TA'
                                next_empty_row = ws.range(f'I{ws.cells(ws.cells.rows.count, "I").end("up").row + 1}').end('up').row + 1

                                # Print the next empty cell where data will be pasted
                                next_empty_cell = ws.range(f'I{next_empty_row}')
                                print(f"Next empty cell in column 'I' (N° TA) is at row {next_empty_row}")

                                # Paste the clipboard value into the next empty cell
                                next_empty_cell.value = clipboard_value

                                # Save the changes to the Excel file
                                wb.save()

                                # Close the Excel file
                                wb.close()

                                print(f"Pasted value '{clipboard_value}' into Excel in cell 'I{next_empty_row}'")

                            except Exception as e:
                                print(f"Error while working with Excel: {str(e)}")
                            finally:
                                # Quit the Excel app
                                app.quit()
                        except Exception as e:
                            print(f"Error working with Excel: {str(e)}")
                    else:
                        print("Clipboard is empty. Nothing to paste into Excel.")
                else:
                    print(f"SAP code not found for DEST '{dest}'. Exiting...")
            else:
                print(f"FF code not found for FF name '{ff_name}'. Exiting...")

            # After processing one CV, return to the main menu with F3
            time.sleep(1)
            pyautogui.press('F3')

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
    