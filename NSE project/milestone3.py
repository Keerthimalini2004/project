import csv
import os
import random
import shutil
import time
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException
from csv import Error as CSVError
import logging
import smtplib
from email.mime.text import MIMEText


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(message)s',
    filename='Log_report.txt',  # Log file name
    filemode='a'  # Append mode
)

download_stats = {"attempts": 0, "success": 0, "failure": 0,"duplicates": 0}

def check_duplicate_file(file_name, folder_path):
    if file_name in os.listdir(folder_path):
        logging.warning(f"Duplicate file {file_name} found. Skipping download.")
        download_stats["duplicates"] += 1
        return True
    return False


def send_notification(subject, files=None):
    global download_stats
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    FROM_EMAIL = "keerthimalini797@gmail.com"
    PASSWORD = "dwig neea ewes cbvo"
    TO_EMAIL = "keerthimalinir2004@gmail.com"

    body_content = f"""
Dear User,

This is an automated email notification that confirms the successful download of NSE Reports.

Report Summary:
    -Total reports attempted:{download_stats["attempts"]}
    -Successfully downloaded reports:{download_stats["success"]}
    -Failed to download:{download_stats["failure"]}
    -Duplicate files: {download_stats["duplicates"]}

Please review downloaded files and log report for any errors.Log file attached below for detailed information.

For any issues,please contact keerthimalinir2004@gmail.com.


Best regards,
Your Automation System
           """
    try:
        # Create a multipart message
        msg = MIMEMultipart()
        msg['From'] = FROM_EMAIL
        msg['To'] = TO_EMAIL
        msg['Subject'] = subject

        msg.attach(MIMEText(body_content, 'plain'))

        # Attach files if provided
        if files:
            for file in files:
                with open(file, 'rb') as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(file))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file)}"'
                    msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(FROM_EMAIL, PASSWORD)
        server.sendmail(FROM_EMAIL, TO_EMAIL, msg.as_string())
        server.quit()
        logging.info("Notification sent successfully")
    except Exception as e:
        logging.error(f"Error sending notification: {e}")

    logging.info("Download process completed.")
    logging.info(f"Total files attempted: {download_stats['attempts']}")
    logging.info(f"Total files downloaded successfully: {download_stats['success']}")
    logging.info(f"Total files failed to download: {download_stats['failure']}")
    logging.info(f"Total duplicate files: {download_stats['duplicates']}")


def validate_csv_file(file_path):
    try:
        with open(file_path, 'r') as file:
            csv_reader = csv.reader(file)
            rows = list(csv_reader)
            if len(rows) == 0:
                logging.warning(f"CSV file {file_path} is empty")
                return False
            header_row = rows[0]
            if len(header_row) == 0:
                logging.warning(f"CSV file {file_path} has no header")
                return False
            for row in rows[1:]:
                if len(row) != len(header_row):
                    logging.warning(f"CSV file {file_path} has inconsistent row length")
                    return False
            logging.info(f"CSV file {file_path} is valid")
            return True
    except CSVError as e:
        logging.error(f"Error parsing CSV file {file_path}: {e}")
        return False
    except Exception as e:
        logging.error(f"Error validating CSV file {file_path}: {e}")
        return False


def organize_files_by_date_and_extension(source_folder):
    try:
        for file_name in os.listdir(source_folder):
            file_path = os.path.join(source_folder, file_name)
            if os.path.isfile(file_path):
                timestamp = os.path.getctime(file_path)
                file_date = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d")
                date_folder = os.path.join(source_folder, file_date)
                os.makedirs(date_folder, exist_ok=True)

                csv_folder = os.path.join(date_folder, "CSV")
                dat_folder = os.path.join(date_folder, "DAT")
                zip_folder = os.path.join(date_folder, "ZIP")
                os.makedirs(csv_folder, exist_ok=True)
                os.makedirs(dat_folder, exist_ok=True)
                os.makedirs(zip_folder, exist_ok=True)

                if file_name.endswith('.csv'):
                    validation_result = validate_csv_file(file_path)
                    if validation_result:
                        shutil.move(file_path, os.path.join(csv_folder, file_name))
                        logging.info(f"Moved valid CSV file {file_name} to csv_folder")
                        download_stats["success"] += 1
                    else:
                        logging.warning(f"Did not move invalid CSV file {file_name} to csv_folder")
                        download_stats["failure"] += 1
                elif file_name.endswith('.zip'):
                    shutil.move(file_path, os.path.join(zip_folder, file_name))
                    logging.info(f"{file_name} moved zip file  to zip_folder")
                else:
                    shutil.move(file_path, os.path.join(dat_folder, file_name))
                    logging.info(f"{file_name} moved dat file to dat_folder")
                    download_stats["success"] += 1
                download_stats["attempts"] += 1
        logging.info("Files have been organized by date and extension.")
    except Exception as e:
        logging.error(f"Error organizing files: {e}")


def retry_download_file(file_url, download_folder, max_retries=3,retry_delay=1):
    file_name = file_url.split("/")[-1]

    # Check if file already exists
    if check_duplicate_file(file_name, download_folder):
        download_stats["duplicates"] += 1  # Increment duplicates counter
        logging.warning(f"Duplicate file {file_name} found. Skipping download.")
        return

    retries = 0
    while retries < max_retries:
        try:
            driver.get(file_url)
            logging.info(f"Downloading file: {file_url}")
            download_stats["attempts"] += 1
            download_stats["success"] += 1
            break
        except Exception as e:
            retries += 1
            logging.error(f"Error downloading file (attempt {retries}/{max_retries}): {e}")
            time.sleep(retry_delay + random.uniform(0, 1))
    else:
        download_stats["failure"] += 1
        logging.error(f"Failed to download file after {max_retries} retries.")



url = "https://www.nseindia.com/all-reports"
# Download folder path
download_folder = r"C:\Users\R KEERTHI MALINI\Desktop\NSE project\downloaded_files"

# Create Chrome Options
chrome_options = webdriver.ChromeOptions()

user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (HTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"
chrome_options.add_argument(f"user-agent={user_agent}")

chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safe browsing.enabled": True
})

try:
    # Initialize Chrome Driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
except WebDriverException as e:
    logging.error(f"Error initializing Chrome Driver: {e}")
    exit()

try:
    # Navigate to URL
    driver.get(url)

    # Implicit Wait
    driver.implicitly_wait(40)
    try:
        # Explicit Wait for element visibility
        WebElement = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, "//*[@id='cr_equity_daily']")))
    except TimeoutException:
        logging.error("Timed out waiting for element to be visible")
        download_stats["failure"] += 1
        exit()

    # Find download links
    download_links = WebElement.find_elements(By.XPATH, ".//div[contains(@class,'reportsDownloadIcon')]")

    # Download each file linked in the element
    for link in download_links:
        file_url = link.get_attribute('href')
        if file_url:
            retry_download_file(file_url, download_folder, max_retries=5, retry_delay=2)


    organize_files_by_date_and_extension(download_folder)

    # Send notification
    subject = "Notification - NSE Reports Download Status"
    log_file = 'Log_report.txt'
    send_notification(subject, [log_file])

except Exception as e:
    logging.error(f"Error: {e}")
    download_stats["failure"] += 1
finally:
    logging.info("Downloaded Successfully")
    driver.quit()
