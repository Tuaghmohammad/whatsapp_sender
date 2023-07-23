import time  # Import the time module for delays in the script
import pandas as pd  # Import pandas library for working with Excel data
from selenium import webdriver  # Import Selenium WebDriver for web automation
# Import Keys for keyboard actions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By  # Import By for locating elements

# Function to send a WhatsApp message


def send_whatsapp_message(driver, phone_number, message):
    url = f"https://web.whatsapp.com/send?phone={phone_number}&text={message}"
    driver.execute_script("window.open('', '_blank');")  # Open a new tab
    driver.switch_to.window(driver.window_handles[1])  # Switch to the new tab
    driver.get(url)
    time.sleep(10)  # Wait for the WhatsApp Web page to load
    try:
        driver.find_element(
            By.CSS_SELECTOR, "#main > footer > div._2lSWV._3cjY2.copyable-area > div > span:nth-child(2) > div > div._1VZX7 > div._2xy_p._3XKXx > button > span").click()
        print(f"Message sent to {phone_number}: {message}")
    except Exception as e:
        print(f"Failed to send message to {phone_number}: {str(e)}")
        driver.close()  # Close the current tab if the phone number is not found
        # Switch back to the original tab
        driver.switch_to.window(driver.window_handles[0])

    time.sleep(5)  # Wait for a moment after sending the message

# Read phone numbers and messages from the Excel file


def read_excel_data(file_path):
    df = pd.read_excel(file_path)
    numbers = df['Phone Number'].tolist()
    messages = df['Message'].tolist()
    return numbers, messages


if __name__ == "__main__":
    # Replace 'list.xlsx' with the actual path to your Excel file
    excel_file_path = 'list.xlsx'

    # Initialize the Chrome WebDriver (Make sure to provide the correct path to chromedriver.exe)
    driver = webdriver.Chrome(
        executable_path="chromedriver\chromedriver.exe")

    numbers, messages = read_excel_data(excel_file_path)

    try:
        # Open WhatsApp Web
        driver.get("https://web.whatsapp.com")
        input("Please scan the QR code and press Enter after successful login...")
        time.sleep(5)  # Wait for the page to fully load after login

        for i in range(len(numbers)):
            phone_number = numbers[i]
            message = messages[i]

            send_whatsapp_message(driver, phone_number, message)

    except Exception as e:
        print(f"An error occurred: {str(e)}")

    finally:
        # Keep the browser open after the process is finished
        input("Press Enter to close the browser...")
        # Close the Chrome WebDriver
        driver.quit()
