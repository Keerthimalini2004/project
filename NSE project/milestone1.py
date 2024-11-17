from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import os
from datetime import date


url = "https://www.nseindia.com/all-reports"

# Download folder path
download_folder =r"C:\Users\R KEERTHI MALINI\Desktop\NSE project\downloaded_files"

# Create Chrome Options
chrome_options = webdriver.ChromeOptions()

user_agent ="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (HTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
chrome_options.add_argument(f"user-agent={user_agent}")

chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safe browsing.enabled": True
})

# Create Chrome Driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)


try:
    # Navigate to URL
    driver.get(url)

    # Implicit Wait
    driver.implicitly_wait(40)

    # Explicit Wait for element visibility
    WebElement = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.XPATH, "//*[@id='cr_equity_daily']"))
    )


    # Find download links
    download_links = WebElement.find_elements(By.XPATH, ".//div[contains(@class,'reportsDownloadIcon')]")

    # Get today's date
    today = date.today().strftime("%Y-%m-%d")
    for i, link in enumerate(download_links):
        link.click()
        time.sleep(2)  # Wait for download to start

        filename = f"NSE_File_{today}_{i+1}.csv"
        downloaded_file = os.path.join(download_folder, filename)
        os.replace(os.path.join(download_folder, "file.csv"), downloaded_file)
        print("Files downloaded successfully to:",download_folder )

#except Exception as e:
#print(f"Error: {e}")

finally:
       print("Downloaded Successfully")
       driver.quit()