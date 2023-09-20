import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import time

# Get the directory where the script is located
script_directory = os.path.dirname(os.path.abspath(__file__))

# Construct the relative path to the Excel file
excel_filename = "AWLIST.xlsx"
excel_file_path = os.path.join(script_directory, excel_filename)

# Open the Excel file
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active

# Set up the web driver (make sure the Chrome WebDriver executable is in your system's PATH)
driver = webdriver.Chrome()

# URL of the login page
login_url = "http://192.168.6.233/etbox_yadran/aspnet_vb/index.aspx"

# Navigate to the login page
driver.get(login_url)

# Fill in the username and password fields using the By class
username_field = driver.find_element(By.ID, "txtUserName")
username_field.send_keys("iaguero")

password_field = driver.find_element(By.ID, "txtUserPass")
password_field.send_keys("iaguero17")

# Submit the login form
password_field.send_keys(Keys.RETURN)

# Sleep for a few seconds to allow the login to complete (you can adjust this as needed)
time.sleep(5)

# URL of the web page
base_url = "http://192.168.6.233/etbox_yadran/aspnet_vb/ComexContratoVenta.aspx?iIdContratoVenta=5555&esCierre="

# Loop through Excel rows and update the web page
for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming data starts from row 2
    order_number = row[0]  # Assuming order number is in the first column (column A)
    new_url = base_url.replace("Venta=5555", f"Venta={order_number}")

    # Navigate to the updated URL
    driver.get(new_url)

    # Find and modify the elements as described
    modify_button = driver.find_element(By.XPATH, "//span[contains(text(),'Modificar')]")
    modify_button.click()

    reserva_input = driver.find_element(By.ID, "txtReserva_I")
    reserva_input.clear()
    reserva_input.send_keys(row[1])  # B2 value

    awb_input = driver.find_element(By.ID, "txtAWBCRTBL_I")
    awb_input.clear()
    awb_input.send_keys(row[1])  # B2 value

    motonave_input = driver.find_element(By.ID, "txtMotonave_I")
    motonave_input.clear()
    motonave_input.send_keys(row[2])  # C2 value

    contrato_destino_input = driver.find_element(By.ID, "txtContratoDestino_I")
    contrato_destino_input.clear()
    contrato_destino_input.send_keys("1" + contrato_destino_input.get_attribute("value"))  # Add "1" to the existing value

    guardar_button = driver.find_element(By.XPATH, "//span[contains(text(),'Guardar')]")
    guardar_button.click()

    # Sleep for a few seconds to allow the changes to save (you can adjust this as needed)
    time.sleep(1)

# Close the web driver and Excel file when done
driver.quit()
workbook.close()
