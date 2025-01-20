from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
from pathlib import Path
import logging
import time
from dataclasses import dataclass

@dataclass
class SharePointConfig:
    email: str
    password: str
    base_url: str
    mayorca_folder_url: str
    monteria_folder_url: str

class SharePointUploader:
    def __init__(self, config: SharePointConfig):
        self.config = config
        self.setup_logging()
    
    def setup_logging(self):
        """Configura el sistema de logging"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('sharepoint_upload.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

        def handle_replace_dialog(self, driver: webdriver.Firefox, wait: WebDriverWait):
            """Maneja el diálogo de reemplazo de archivo si aparece"""
            try:
                
                replace_dialog = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//div[contains(@role, 'dialog') and contains(.//span, 'replace')]")
                ))
            
          
                replace_button = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(.//span, 'Replace') or contains(.//span, 'Reemplazar')]")
                ))
            
          
                try:
                    replace_button.click()
                except:
                    driver.execute_script("arguments[0].click();", replace_button)
                
                self.logger.info("Archivo existente reemplazado exitosamente")
                return True
            
            except TimeoutException:
             
                return True
            except Exception as e:
                self.logger.error(f"Error manejando el diálogo de reemplazo: {str(e)}")
                return False

    def upload_file_to_folder(self, driver: webdriver.Firefox, wait: WebDriverWait, 
                            folder_url: str, file_path: Path) -> bool:
        """Sube un archivo a una carpeta específica de SharePoint"""
        try:
            driver.get(folder_url)
            time.sleep(5)
            
            upload_button = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//span[contains(text(), 'Cargar') or contains(text(), 'Upload')]")))
            
            driver.execute_script("arguments[0].scrollIntoView(true);", upload_button)
            time.sleep(2)
            
            try:
                upload_button.click()
            except:
                driver.execute_script("arguments[0].click();", upload_button)

            files_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(text(), 'Archivos') or contains(text(), 'Files')]")))
            
            try:
                files_button.click()
            except:
                driver.execute_script("arguments[0].click();", files_button)

            file_input = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[@type='file']")))
            
            file_input.send_keys(str(file_path.absolute()))
            

            time.sleep(5)
            

            if not self.handle_replace_dialog(driver, wait):
                return False
            
     
            time.sleep(10)
            
            try:
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//*[contains(text(), '{file_path.name}')]")))
            except TimeoutException:
                self.logger.warning(f"No se pudo verificar la subida de {file_path.name}, pero el proceso terminó sin errores")
            
            self.logger.info(f"Archivo {file_path.name} subido exitosamente a {folder_url}")
            return True

        except Exception as e:
            self.logger.error(f"Error subiendo archivo {file_path}: {str(e)}")
            driver.save_screenshot(f"error_upload_{file_path.name.replace('.', '_')}.png")
            return False

    def setup_driver(self):
        """Configura el driver de Firefox para la subida de archivos"""
        options = Options()
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream")
        return webdriver.Firefox(options=options)

    def login_to_sharepoint(self, driver: webdriver.Firefox, wait: WebDriverWait):
        """Maneja el proceso de login a SharePoint"""
        try:
            driver.get("https://login.microsoftonline.com")
            
      
            email_elem = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
            email_elem.send_keys(self.config.email)
            wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

      
            try:
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "lightbox-cover")))
            except TimeoutException:
                driver.execute_script("document.querySelector('.lightbox-cover').style.display = 'none';")

           
            password_elem = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
            password_elem.send_keys(self.config.password)
            

            for _ in range(2):
                wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                time.sleep(1)

            return True
        except Exception as e:
            self.logger.error(f"Error en login: {str(e)}")
            driver.save_screenshot("login_error.png")
            return False

    def upload_file_to_folder(self, driver: webdriver.Firefox, wait: WebDriverWait, 
                            folder_url: str, file_path: Path) -> bool:
        """Sube un archivo a una carpeta específica de SharePoint"""
        try:
           
            driver.get(folder_url)
            time.sleep(5)  
            upload_button = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//span[contains(text(), 'Cargar') or contains(text(), 'Upload')]")))
            
   
            driver.execute_script("arguments[0].scrollIntoView(true);", upload_button)
            time.sleep(2) 
            

            try:
                upload_button.click()
            except:
                driver.execute_script("arguments[0].click();", upload_button)
            
            files_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(text(), 'Archivos') or contains(text(), 'Files')]")))
            
            try:
                files_button.click()
            except:
                driver.execute_script("arguments[0].click();", files_button)
            

            file_input = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[@type='file']")))

            file_input.send_keys(str(file_path.absolute()))

            time.sleep(10)
            
            try:
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//*[contains(text(), '{file_path.name}')]")))
            except TimeoutException:
                self.logger.warning(f"No se pudo verificar la subida de {file_path.name}, pero el proceso terminó sin errores")
            
            self.logger.info(f"Archivo {file_path.name} subido exitosamente a {folder_url}")
            return True

        except Exception as e:
            self.logger.error(f"Error subiendo archivo {file_path}: {str(e)}")
            driver.save_screenshot(f"error_upload_{file_path.name.replace('.', '_')}.png")
            return False

    def upload_files(self, max_retries=3):
        """Proceso principal de subida de archivos con reintentos"""
        driver = None
        try:
            driver = self.setup_driver()
            wait = WebDriverWait(driver, 30)

            for attempt in range(max_retries):
                if self.login_to_sharepoint(driver, wait):
                    break
                if attempt < max_retries - 1:
                    self.logger.warning(f"Reintento {attempt + 1} de login...")
                    time.sleep(5)
                else:
                    raise Exception("Fallo en el login a SharePoint después de todos los reintentos")
            
            def upload_with_retry(file_path: Path, folder_url: str):
                for attempt in range(max_retries):
                    if self.upload_file_to_folder(driver, wait, folder_url, file_path):
                        return True
                    elif attempt < max_retries - 1:
                        self.logger.warning(f"Reintento {attempt + 1} de subida para {file_path}...")
                        time.sleep(5)
                return False

            mayorca_file = Path("Mayorca/2025.csv")
            if mayorca_file.exists():
                if not upload_with_retry(mayorca_file, self.config.mayorca_folder_url):
                    self.logger.error(f"Falló la subida de {mayorca_file} después de {max_retries} intentos")
            else:
                self.logger.error(f"Archivo no encontrado: {mayorca_file}")

            monteria_file = Path("Monteria/2025.csv")
            if monteria_file.exists():
                if not upload_with_retry(monteria_file, self.config.monteria_folder_url):
                    self.logger.error(f"Falló la subida de {monteria_file} después de {max_retries} intentos")
            else:
                self.logger.error(f"Archivo no encontrado: {monteria_file}")

        except Exception as e:
            self.logger.error(f"Error en el proceso de subida: {str(e)}")
            if driver:
                driver.save_screenshot("error_upload.png")
            raise
        finally:
            if driver:
                driver.quit()