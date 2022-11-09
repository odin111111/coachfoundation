from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import gspread as gs
from time import sleep

def botInitialization(headless=False):
    # Initialize the Bot
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    prefs = {"profile.default_content_setting_values.notifications": 2}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("disable-infobars")
    if headless:
        options.add_argument("--headless")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )
    driver.maximize_window()
    driver.implicitly_wait(1)
    return driver

gc = gs.service_account(filename="updated_keys.json")
sheet = gc.open_by_url(
    "https://docs.google.com/spreadsheets/d/1GzCMHktuGJjpp8Eo6eNdMlWoPDrJXRQyWpNmBTrPB3Y/edit?usp=sharing"
)
worksheet = sheet.worksheet("input")

records = worksheet.get_all_records()

# write
worksheet = sheet.worksheet("Loading Status")
worksheet = worksheet.clear()
worksheet = sheet.worksheet("Loading Status")
worksheet.append_row(["Link", "Optin Monster Status"])

driver = botInitialization()

for record in records:
    link = record["Link"]
    print("--------------------------------------------")
    print("Currently on link: {}/{}".format(records.index(record) + 1, len(records)))
    print("Link: {}".format(link))
    
    driver.get(link)
    try:
        WebDriverWait(driver, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//h1[contains(text(), "Bad Request")]',
                    )
                )
            )
        driver.get(link)
    except:
        pass
    
    for i in range(4):
        try:
            driver.find_element(By.XPATH, '//body').send_keys(Keys.PAGE_DOWN)
        except:
            break
        
    optin_monster_status = "No"
    for i in range(20):
        try:
            try:
                driver.find_element(By.XPATH, '//body').send_keys(Keys.PAGE_DOWN)
                sleep(1)
                driver.find_element(By.XPATH, '//body').send_keys(Keys.PAGE_DOWN)
                driver.find_element(By.XPATH, '//body').send_keys(Keys.PAGE_UP)
            except:
                pass
            WebDriverWait(driver, 1).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//body//div//*[contains(@id, "ImageElement") or contains(@id,"TextElement")]',
                    )
                )
            )
            optin_monster_status = "Yes"
            break
        except:
            continue
        
    worksheet.append_row([link, optin_monster_status])
    print("Optin Master Status: {}".format(optin_monster_status))
    print("--------------------------------------------")

driver.close()
driver.quit()
print("Finished...")
dummy = input("Press any key to exit...")
exit()
