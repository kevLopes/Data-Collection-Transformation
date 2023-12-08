import os
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import re  # Import the re module for cleaning file names


# Define a function to clean file names
def clean_file_name(name):
    # Remove invalid characters (replace with underscores)
    cleaned_name = re.sub(r'[\/:*?"<>|]', '_', name)
    return cleaned_name


# Define file paths
excel_file_path = r"C:\Users\keven.deoliveiralope\Documents\File Extraction App\Fata\Invoice - Overview 27 Nov 23 for Keven.xlsx"
download_folder_path = r"C:\Users\keven.deoliveiralope\Documents\File Extraction App\Fata\Invoices PDF"

# Set up Chrome options
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_folder_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

# Initialize WebDriver
chrome_driver_path = "C:/Users/keven.deoliveiralope/Downloads/chromedriver-win64/chromedriver.exe"
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open Excel file
workbook = openpyxl.load_workbook(excel_file_path)
worksheet = workbook.active


def wait_for_download_completion(directory, timeout=60):
    # Wait until the download is complete (no .crdownload files)
    time.sleep(8)
    start_time = time.time()
    while True:
        if all(not f.endswith('.crdownload') for f in os.listdir(directory)):
            return True
        if time.time() - start_time > timeout:
            return False
        time.sleep(1)

# Counter for file naming
file_counter = 1

# URL Prefix
url_prefix = "chrome-extension://hehijbfgiekmjfkfjpbkbammjbdenadd/nhc.htm#url="

# Loop through the rows starting from row
for row in worksheet.iter_rows(min_row=2):
    url_cell = row[14]  # URLs are in column F
    name_cell = row[15]  # 'Purchase Invoice DbCapture' in column AE
    if url_cell.hyperlink:
        url = url_cell.hyperlink.target
        print(f"Opening URL: {url}")

        # Navigate to the URL
        driver.get(url)
        time.sleep(8)  # Wait for the file to download
        # Wait for the download to complete
        if wait_for_download_completion(download_folder_path):
            # Find the latest file with the desired name
            downloaded_files = [f for f in os.listdir(download_folder_path) if f.startswith("SunInvoicesPreview")]
            if downloaded_files:
                latest_file = max(downloaded_files, key=lambda x: os.path.getctime(os.path.join(download_folder_path, x)))
                downloaded_file_path = os.path.join(download_folder_path, latest_file)
                # Clean the file name from column P
                new_name = clean_file_name(name_cell.value)
                new_file_name = f"{file_counter} - {new_name}.pdf"
                new_file_path = os.path.join(download_folder_path, new_file_name)
                os.rename(downloaded_file_path, new_file_path)
                file_counter += 1  # Increment the counter
        else:
            print(f"Download timed out for URL: {url}")

        time.sleep(2)

# Close WebDriver
driver.quit()

print("Task completed. PDFs downloaded and renamed in:", download_folder_path)
