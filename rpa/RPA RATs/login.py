from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def login(driver):
    wait = WebDriverWait(driver, 20)

    driver.get("https://eprivacy.ecix.tech/login/")

    wait.until(EC.presence_of_element_located((By.ID, "loginUser"))).send_keys("marta.pena@asisa.es")
    wait.until(EC.presence_of_element_located((By.ID, "loginPass"))).send_keys("Prueba2026")
    wait.until(EC.element_to_be_clickable((By.ID, "submit"))).click()

    wait.until(EC.url_changes("https://eprivacy.ecix.tech/login/"))
    print("✅ Login OK")