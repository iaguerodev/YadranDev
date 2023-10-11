import pandas as pd
import os
import unicodedata

# Constants
EXCEL_FILE_NAME = "CVAWB.xlsx"
EXCEL_SHEET_NAME = "Hoja1"  # Replace with the actual sheet name

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
            year = int(ano_cell_content)
        else:
            raise ValueError("Invalid year value in cell 'AÑO'. It should start with '202'.")

        # Debugging line to print the converted year
        print(f"Year after conversion: {year}")

        # Check if the 'Itin.trans' column is not empty before applying date conversion
        if 'Itin.trans' in df.columns and not df['Itin.trans'].empty:
            # Convert the date column 'Itin.trans' and store the results in 'Date conv' column 'Date conv'
            df['Date conv'] = df['Itin.trans'].apply(lambda x: convert_date_format(x, year))

            # Save the updated DataFrame back to the Excel file
            df.to_excel(excel_file_path, sheet_name=EXCEL_SHEET_NAME, index=False)

            print("Date conversion completed.")
        else:
            print("No 'Itin.trans' column found or it is empty. No date conversion performed.")

    except FileNotFoundError as e:
        print(f"File not found error: {str(e)}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
