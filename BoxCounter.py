from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime
import threading

# Setup Selenium WebDriver
chrome_options = Options()
driver = None

# Dictionary to store history of changes
contract_history = {}

# Function to initialize WebDriver
def init_driver():
    global driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Function to Scrape Data
def scrape_data():
    global driver
    if driver is None:
        init_driver()

    extracted_data = []

    try:
        url = "https://containertrack.yadran.cl/smarttv.php"
        driver.get(url)

        # Wait for the elements to load
        time.sleep(10)

        # Find all contract elements on the page
        contract_elements = driver.find_elements(By.CLASS_NAME, 'small-box')

        for contract in contract_elements:
            try:
                # Extract all visible text from each contract element
                contract_text = contract.text.strip()
                extracted_data.append(contract_text)
                print(f"Extracted Data: {contract_text}")

            except Exception as e:
                print(f"Error processing contract: {e}")

    except Exception as e:
        print(f"Error during scraping: {e}")
        if "no such window" in str(e).lower():
            print("Reinitializing WebDriver...")
            driver.quit()
            init_driver()

    finally:
        # Close the browser after scraping
        driver.quit()
        driver = None

    return extracted_data

# Function to Update History and Track Changes
def update_history(new_data):
    timestamp = datetime.now().strftime("%H:%M %d/%m/%Y")
    updated_data = []

    for item in new_data:
        lines = item.split('\n')
        if len(lines) >= 3:
            cv_number = lines[0]
            company_name = lines[1]
            box_count = lines[2]

            if cv_number not in contract_history:
                contract_history[cv_number] = []

            # Check if there's a change in box count
            if not contract_history[cv_number] or contract_history[cv_number][-1]["box_count"] != box_count:
                contract_history[cv_number].append({"box_count": box_count, "timestamp": timestamp})

            # Format the history for display
            history = "\n".join([f"{entry['box_count']} {entry['timestamp']}" for entry in contract_history[cv_number]])
            updated_data.append(f"{cv_number}\n{company_name}\n{history}")

    return updated_data

# Function to Display Data in a Window
def display_data(data, text_area):
    text_area.config(state=tk.NORMAL)
    text_area.delete('1.0', tk.END)  # Clear previous data

    for item in data:
        text_area.insert(tk.END, item + "\n\n")

    text_area.config(state=tk.DISABLED)

# Function to start the scraping process periodically
def start_scraping(text_area):
    while True:
        scraped_data = scrape_data()
        updated_data = update_history(scraped_data)
        display_data(updated_data, text_area)
        time.sleep(300)  # Wait for 5 minutes before the next check

# Main Logic
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Extracted Data History")

    text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=20, font=("Arial", 12))
    text_area.pack(padx=10, pady=10)

    # Start the scraping in a separate thread
    scraping_thread = threading.Thread(target=start_scraping, args=(text_area,))
    scraping_thread.daemon = True
    scraping_thread.start()

    root.mainloop()