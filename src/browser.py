from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

class Browser(object):
    def __init__(self, base_url: str) -> None:
        """
        Class constructor for Browser 

        Args:
            folder_path_to_download (str): folder path to download files
            base_url (str): base URL for the browser

        Attributes:
            self.driver (webdriver): selenium self.driver
        """
        self.base_url: str = base_url
        self.folder_path_to_download: str = ""

        chrome_options = webdriver.ChromeOptions()
        #chrome_options.add_argument('--headless') # Disable headless mode
        chrome_options.add_argument('--disable-gpu') # Disable GPU acceleration
        chrome_options.add_argument('--disable-software-rasterizer')  # Disable software rasterizer
        chrome_options.add_argument('--disable-dev-shm-usage') # Disable shared memory usage
        chrome_options.add_argument('--no-sandbox') # Disable sandbox
        chrome_options.add_argument('--disable-extensions') # Disable extensions
        chrome_options.add_argument('--disable-sync') # Disable syncing to a Google account
        chrome_options.add_argument('--disable-webgl') # Disable WebGL 
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--disable-gl-extensions') # Disable WebGL extensions
        chrome_options.add_argument('--disable-in-process-stack-traces') # Disable stack traces
        chrome_options.add_argument('--disable-logging') # Disable logging
        chrome_options.add_argument('--disable-cache') # Disable cache
        chrome_options.add_argument('--disable-application-cache') # Disable application cache
        chrome_options.add_argument('--disk-cache-size=1') # Set disk cache size to 1
        chrome_options.add_argument('--media-cache-size=1') # Set media cache size to 1
        chrome_options.add_argument('--kiosk-printing') # Enable kiosk printing
        chrome_options.add_argument('--kiosk-pdf-printing')

        chrome_prefs = {
            "download.prompt_for_download": False,
            "plugins.always_open_pdf_externally": True,
            "download.open_pdf_in_system_reader": False,
            "profile.default_content_settings.popups": 0,
            "printing.print_to_pdf": True,
            "download.default_directory": self.folder_path_to_download,
            "savefile.default_directory": self.folder_path_to_download
        }

        chrome_options.add_experimental_option("prefs", chrome_prefs)

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.minimize_window()

        self.wait = WebDriverWait(self.driver, 10)

    def complete_login_form(self, username: str, password: str) -> None:
        """
        Completes login form

        Args:
            driver (webdriver): selenium driver
            username (str): username
            password (str): password
        """
        self.driver.get(f"{self.base_url}/index.php")
        
        self.wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/form/input[1]")))
        self.driver.find_element(By.XPATH, "/html/body/form/input[1]").send_keys(username)
        self.driver.find_element(By.XPATH, "/html/body/form/input[2]").send_keys(password)
        self.driver.find_element(By.XPATH, "/html/body/form/button").click()

    def print_label_document(self, tracking_number: str) -> None:
        url_rotulo = f"{self.base_url}/srv.RotuloFCSPdf.emitir+id={tracking_number[:7]}&idservicio={tracking_number[:7]}"
        self.__print_webpage__(url_rotulo)

    def close(self) -> None:
        """Closes the browser."""
        try:
            self.driver.quit()
        except Exception as e:
            raise Exception(f"Error quitting driver: {e}")

    def __print_webpage__(self, url: str) -> None:
        try:
            self.driver.get(url) # this print since chrome options are set to print automatically
            self.driver.implicitly_wait(5)

        except Exception as e:
            raise Exception(f"Error printing documents: {e}")