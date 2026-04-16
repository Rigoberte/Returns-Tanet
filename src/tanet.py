import requests
import pandas as pd

from src.browser import Browser

class Tanet:
    def __init__(self) -> None:
        #self.environment: str = "development"
        self.environment: str = "testing"

        if self.environment == "production":
            self.base_url: str = "https://sgi.tanet.com.ar/sgi"
        else:
            self.base_url: str = "http://sgi.tatest.com.ar/sgi/sgi"
        
        self.rsid: str = ""
        self.username: str = ""
        self.password: str = ""

    def login(self, username: str, password: str) -> None:
        login_url = f"{self.base_url}/usr.usrJSON.login+"
        login_url += f"login_username={username}&login_password={password}"

        response = self.__do_a_http_request__(login_url)

        if response.get('result') == 'FAIL':
            raise Exception("Login failed. Please check your credentials.")
        
        self.rsid = response['data']['data'].get('RSID')
        self.username = username
        self.password = password
        
        return
    
    def build_driver_and_login(self) -> None:
        if not self.username or not self.password:
            raise Exception("Username and password must be set before logging in.")
        
        self.browser = Browser(base_url=self.base_url)
        self.browser.complete_login_form(self.username, self.password)

    def print_label_document(self, tracking_number: str) -> None:
        if not hasattr(self, 'browser'):
            raise Exception("Browser not initialized. Please call build_driver_and_login() first.")
        
        self.browser.print_label_document(tracking_number)
    
    def close_browser(self) -> None:
        if not hasattr(self, 'browser'):
            raise Exception("Browser not initialized. Please call build_driver_and_login() first.")
        
        self.browser.close()
    
    def load_site_data(self) -> pd.DataFrame:
        if self.environment == "production":
        
            if not self.rsid:
                raise Exception("You must login first to obtain an RSID.")

            site_id_url = f"{self.base_url}/srv.SrvClienteJSON.buscarUbicacion+"
            site_id_url += f"RSID={self.rsid}&nomlinea=&site=&idubicacion="

            response = self.__do_a_http_request__(site_id_url)

            if response.get('result') == 'FAIL':
                raise Exception("Failed to retrieve site ID. Please check the protocol and site number. " + "Error: " + str(response.get('data')))

            return pd.DataFrame(response.get('data').get('data')).T
        
        else:
            return pd.read_excel(r"C:\Users\inaki.costa\Downloads\site_id_data.xlsx")
    
    def create_return(self, id_ubicacion: str, referencia: str, 
            retiradde: str, retirahta: str, entregadde: str, entregahta: str,
            obsOper: str, tipomaterial: str, 
            cajas: int) -> dict:
        
        if not self.rsid:
            raise Exception("You must login first to obtain an RSID.")
        
        temperatura = 1 # Valor fijo para temperatura 'Ambiente' según la documentación de TANET

        material_mapping = {
            "Biologico No Infeccioso": 1,
            "Biologico Infeccioso": 2,
            "Medicacion": 3,
            "Otros": 4,
            "Documentacion": 5,
            "Material Clinico": 6,
            "Veterinario": 7,
            "Equipos": 8,
            "Equipo con Bateria": 9,
            "Productos para Destruccion": 10
        }

        if tipomaterial.capitalize().strip() not in material_mapping:
            raise ValueError(f"Invalid tipomaterial: {tipomaterial}. Must be one of the following: {list(material_mapping.keys())}")

        number_tipomaterial = str(material_mapping[tipomaterial.capitalize().strip()])

        order_url = f"{self.base_url}/srv.SrvClienteJSON.crearRecoleccion+"
        order_url += f"RSID={self.rsid}&idubicacion={id_ubicacion}&referencia={referencia}"
        order_url += f"&retiradde={retiradde}&retirahta={retirahta}"
        order_url +=  f"&entregadde={entregadde}&entregahta={entregahta}"
        order_url += f"&obsOper={obsOper}&tipomaterial={number_tipomaterial}&temperatura={str(temperatura)}"
        order_url += f"&cajas={str(cajas)}"

        response = self.__do_a_http_request__(order_url)

        if response.get('result') == 'FAIL':
            raise Exception("Failed to create return. Please check the protocol and site number. " + "Error: " + str(response.get('data')))

        return response.get('data')

    def __do_a_http_request__(self, url: str) -> dict:
        response = requests.get(url)
        fail_dict = {"result": "FAIL", "data": "N/A"}

        if response.status_code != 200:
            return fail_dict

        try:
            response_data = response.json()
        except ValueError:
            return fail_dict
            

        if 'err' in response_data and response_data['err']:
            return {"result": "FAIL", "data": response_data['err']}
        
        success_dict = {"result": "OK", "data": response_data}
        return success_dict

    def __standarize_contacts__(self, contacts: str) -> str:
        replacements = [" / ", "/ ", " /", "/", 
                        " ; ", "; ", " ;", ";", 
                        " , ", ", ", " ,", ",",
                        " - ", "- ", " -", "-"]
        for replacement in replacements:
            contacts = contacts.replace(replacement, ", ")

        contacts = contacts.title() # Capitalize
        contacts = contacts.strip() # Trim

        if contacts[-1:] == ",":
            contacts = contacts[:-1]

        return contacts