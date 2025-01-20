from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.common.keys import Keys
import os
import dotenv
import logging
import time
from datetime import datetime
from urllib.parse import quote
from config import STORE_CONFIGS

class ShopifyAutomation:
    def __init__(self, store_type):
        if store_type not in STORE_CONFIGS:
            raise ValueError(f"Store type must be one of {list(STORE_CONFIGS.keys())}")
        
        self.store_config = STORE_CONFIGS[store_type]
        self.setup_folders()
        self.setup_logging()
        dotenv.load_dotenv()

    def setup_folders(self):
        """Crea la estructura de carpetas necesaria"""
        self.store_folder = os.path.join(os.getcwd(), self.store_config['folder_name'])
        if not os.path.exists(self.store_folder):
            os.makedirs(self.store_folder)
            logging.info(f"Created folder: {self.store_folder}")

        

    def setup_logging(self):
        """Configura el sistema de logging para la automatización"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('shopify_automation.log'),
                logging.StreamHandler()
            ]
        )

    def get_shopify_url(self):
        """Genera la URL de Shopify con los parámetros de consulta necesarios"""
        start_date = "2024-12-01"
        end_date = datetime.now().strftime('%Y-%m-%d')
        
        query = f"""FROM sales
SHOW quantity_ordered, gross_sales, discounts, total_sales, net_sales
GROUP BY order_name, day, order_payment_status, customer_id, product_title, product_variant_sku, product_variant_price, product_type, line_type, order_or_return, is_canceled_order, customer_name, customer_email, customer_first_order_date, customer_last_order_date, new_or_returning_customer, customer_email_subscription_status, customer_sms_subscription_status
TIMESERIES day
WITH TOTALS
HAVING quantity_ordered != 0
SINCE {start_date}
UNTIL {end_date}
ORDER BY day ASC
LIMIT 1000
VISUALIZE total_sales TYPE line"""
        
        encoded_query = quote(query)
        return f"{self.store_config['base_url']}?ql={encoded_query}"

    def setup_firefox_driver(self):
        """Configura y retorna una instancia del driver de Firefox con las opciones necesarias"""
        firefox_options = Options()
        firefox_options.set_preference("browser.download.folderList", 2)
        firefox_options.set_preference("browser.download.dir", self.store_folder) 
        firefox_options.set_preference("browser.download.useDownloadDir", True)
        firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/csv,text/csv")
        firefox_options.set_preference("browser.download.manager.showWhenStarting", False)
        firefox_options.set_preference("browser.download.manager.focusWhenStarting", False)
        firefox_options.set_preference("browser.download.manager.closeWhenDone", True)
        
        firefox_options.set_preference("browser.download.manager.alertOnEXEOpen", False)
        firefox_options.set_preference("browser.download.manager.useWindow", False)
        firefox_options.set_preference("browser.download.manager.addToRecentDocs", False)
        firefox_options.set_preference("browser.download.always_ask_before_handling_new_types", False)

        firefox_options.set_preference("browser.download.lastDir", self.store_folder)
        firefox_options.set_preference("browser.download.defaultFolder", self.store_folder)
        firefox_options.set_preference("browser.download.manager.defaultFileName", self.store_config['output_file'])
        
        return webdriver.Firefox(options=firefox_options)

    def wait_and_click(self, driver, wait, xpath, message, timeout=5):
        """Espera a que un elemento sea clickeable y lo clickea"""
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                logging.info(f"{message} (Attempt {attempt + 1}/{max_attempts})")
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", element)
                time.sleep(timeout)
                return True
            except Exception as e:
                if attempt == max_attempts - 1:
                    logging.error(f"Failed to {message.lower()}: {str(e)}")
                    driver.save_screenshot(f"error_{message.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
                    return False
                time.sleep(2)
        return False

    def rename_downloaded_file(self):
        """Renombra el archivo descargado al nombre deseado"""
        max_attempts = 10
        attempt = 0
        while attempt < max_attempts:
            time.sleep(2)
            for filename in os.listdir(self.store_folder):
                if filename.endswith('.csv') and filename != self.store_config['output_file']:
                    try:
                        old_path = os.path.join(self.store_folder, filename)
                        new_path = os.path.join(self.store_folder, self.store_config['output_file'])
                        
                        if os.path.exists(new_path):
                            os.remove(new_path)
                        
                        os.rename(old_path, new_path)
                        logging.info(f"Archivo renombrado exitosamente a {self.store_config['output_file']} en {self.store_folder}")
                        return True
                    except Exception as e:
                        logging.error(f"Error al renombrar archivo: {e}")
                        attempt += 1
                        continue
            attempt += 1
            time.sleep(2) 
        
        logging.error("No se pudo renombrar el archivo después de varios intentos")
        return False

    def shopify_login(self):
        """Realiza el proceso completo de login y exportación de datos"""
        driver = None
        try:
            for var in ['SHOPIFY_EMAIL', 'SHOPIFY_PASSWORD']:
                if not os.getenv(var):
                    raise ValueError(f"Missing environment variable: {var}")

 
            driver = self.setup_firefox_driver()
            driver.maximize_window()
            wait = WebDriverWait(driver, 20)


            driver.get("https://accounts.shopify.com/store-login")
            time.sleep(3)

            email_input = wait.until(EC.presence_of_element_located((By.ID, "account_email")))
            email_input.clear()
            email_input.send_keys(os.getenv('SHOPIFY_EMAIL'))

            if not self.wait_and_click(driver, wait, "//button[@type='submit']", "Clicking Next button"):
                raise Exception("Failed to click Next button")
            
            try:
                password_input = wait.until(EC.element_to_be_clickable((By.ID, "account_password")))
                time.sleep(2)
                password_input.clear()
                password_input.send_keys(os.getenv('SHOPIFY_PASSWORD'))
            except Exception as e:
                logging.error(f"Password field error: {str(e)}")
                driver.save_screenshot("password_error.png")
                raise

            if not self.wait_and_click(driver, wait, "//button[@type='submit']", "Clicking Login button"):
                raise Exception("Failed to click Login button")
            time.sleep(5)
            

            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".Polaris-Navigation, .Polaris-TopBar")))
            except TimeoutException:
                driver.save_screenshot("login_failed.png")
                raise Exception("Login verification failed")

            shopify_url = self.get_shopify_url()
            driver.get(shopify_url)
            time.sleep(8)

            try:
                more_actions_button = "//button[contains(@class, '_Button_1yxn0_1') and .//shopify-internal-icon[@type='menu-horizontal']]"
                if not self.wait_and_click(driver, wait, more_actions_button, "Clicking more actions button"):
                    raise Exception("Failed to click more actions button")
                time.sleep(2)
                export_button_xpath = "//button[contains(., 'Export')]"
                if not self.wait_and_click(driver, wait, export_button_xpath, "Clicking Export button"):
                    raise Exception("Failed to click Export button")
                time.sleep(3)
                
                try:
                    csv_radio = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='radio' and @value='csv']")))
                    if not csv_radio.is_selected():
                        csv_radio.click()
                        time.sleep(1)
                except Exception as e:
                    logging.error(f"Error selecting CSV option: {e}")
                    driver.save_screenshot("csv_selection_error.png")
                    raise

                export_final_button = "//button[contains(., 'Export')]"
                if not self.wait_and_click(driver, wait, export_final_button, "Clicking final Export button"):
                    raise Exception("Failed to click final Export button")
                
                time.sleep(5)
                if not self.rename_downloaded_file():
                    logging.error("No se pudo completar el proceso de renombrado del archivo")
                    return False
                    
                return True
                
            except Exception as e:
                logging.error(f"Error clicking more actions button: {e}")
                driver.save_screenshot("more_actions_error.png")
                raise

        except Exception as e:
            logging.error(f"An error occurred during the export process: {e}")
            return False
        finally:
            if driver:
                driver.quit()

    def run(self):
        """Ejecuta el proceso completo con reintentos"""
        max_attempts = 3
        for attempt in range(max_attempts):
            logging.info(f"Attempt {attempt + 1} of {max_attempts}")
            if self.shopify_login():
                logging.info("Script completed successfully")
                break
            else:
                if attempt < max_attempts - 1:
                    logging.info(f"Waiting 10 seconds before next attempt...")
                    time.sleep(10)
                else:
                    logging.error("All attempts failed")

