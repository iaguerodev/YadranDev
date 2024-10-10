import json
import os
import time
import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import threading

# Setup Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument('--disable-gpu')  # Disable GPU acceleration
chrome_options.add_argument('--window-size=1920x1080')  # Set window size
chrome_options.add_argument('--disable-extensions')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--hide-scrollbars')
chrome_options.add_argument('--mute-audio')

# Initialize WebDriver
driver = None

# File to store history
data_file = "contract_history.json"

# Initialize Data File if not present
if not os.path.exists(data_file):
    with open(data_file, 'w') as file:
        json.dump({}, file)

# Function to initialize WebDriver
def init_driver():
    print("Initializing WebDriver with normal Chrome...")
    global driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    print("WebDriver initialized.")

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

    return extracted_data

# Function to Update History and Track Changes
def update_history(new_data):
    timestamp = datetime.now().strftime("%H:%M %d/%m/%Y")
    updated_data = []

    # Load existing history
    with open(data_file, 'r') as file:
        contract_history = json.load(file)

    for item in new_data:
        lines = item.split('\n')
        if len(lines) >= 3:
            # Extract CV number by keeping only the first part before any whitespace or other characters
            cv_number = lines[0].split(':')[-1].strip().split()[0]
            company_name = lines[1]
            box_count = lines[2]

            if cv_number not in contract_history:
                contract_history[cv_number] = []

            # Check if the box_count has changed before adding to history
            if not contract_history[cv_number] or contract_history[cv_number][-1]['box_count'] != box_count:
                contract_history[cv_number].append({"box_count": box_count, "timestamp": timestamp})

            # Format the history for display
            updated_data.append((cv_number, company_name, box_count, timestamp))

    # Save updated history
    with open(data_file, 'w') as file:
        json.dump(contract_history, file, indent=4)

    return updated_data

# Function to Display Data in a Table
def display_data(data, tree):
    # Debug: Print data to check if CV numbers are correct
    print("Displaying Data:")
    for item in data:
        print(item)

    # Clear the treeview before updating
    tree.delete(*tree.get_children())

    for item in data:
        cv_number, company_name, box_count, timestamp = item
        tag = cv_number.replace(" ", "_")  # Create a unique tag for each CV number

        # Insert a new item or update existing one
        tree.insert("", tk.END, values=(cv_number, company_name, box_count, timestamp), tags=(tag,))

        # Handle different formats of box_count
        try:
            if ':' in box_count:  # If box_count contains a time (e.g., "15:44"), mark as dispatched
                tree.tag_configure(tag, background="lightblue")
            else:
                left_count, right_count = map(int, box_count.split('/'))
                if left_count == right_count:
                    tree.tag_configure(tag, background="lightgreen")
                elif left_count > right_count:
                    tree.tag_configure(tag, background="orange")
                else:
                    tree.tag_configure(tag, background="lightgray")
        except ValueError:
            # If box_count is not in the expected format, use a default background
            tree.tag_configure(tag, background="lightgray")

# Function to start the scraping process periodically
def start_scraping(tree):
    while True:
        scraped_data = scrape_data()
        updated_data = update_history(scraped_data)
        display_data(updated_data, tree)
        time.sleep(300)  # Wait for 5 minutes before the next check

# Function to stop the scraping process
def stop_scraping():
    global driver
    if driver is not None:
        driver.quit()
    root.destroy()

# Function to handle double-click event on Treeview item
def on_tree_item_double_click(event):
    item_id = tree.selection()[0]
    cv_number = tree.item(item_id, 'values')[0]

    # Load existing history
    with open(data_file, 'r') as file:
        contract_history = json.load(file)

    if cv_number in contract_history:
        # Remove existing expanded items under the clicked item
        children = tree.get_children(item_id)
        for child in children:
            tree.delete(child)

        # Insert history items under the clicked CV
        for record in contract_history[cv_number]:
            tree.insert(item_id, tk.END, values=("", "", record["box_count"], record["timestamp"]))

# Main Logic
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Extracted Data History")

    columns = ("CV Number", "Company Name", "Box Count", "Timestamp")
    tree = ttk.Treeview(root, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)
    tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    tree.bind("<Double-1>", on_tree_item_double_click)

    stop_button = tk.Button(root, text="Stop Process", command=stop_scraping)
    stop_button.pack(pady=5)

    # Start the scraping in a separate thread
    scraping_thread = threading.Thread(target=start_scraping, args=(tree,))
    scraping_thread.daemon = True
    scraping_thread.start()

    root.mainloop()