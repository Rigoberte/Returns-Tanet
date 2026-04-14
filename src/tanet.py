import requests
import pandas as pd

class Tanet:
    def __init__(self) -> None:
        self.base_url: str = "https://sgi.tanet.com.ar/sgi"
        #self.base_url: str = "http://sgi.tatest.com.ar/sgi/sgi"
        self.rsid: str = ""

    def login(self, username: str, password: str) -> None:
        login_url = f"{self.base_url}/usr.usrJSON.login+"
        login_url += f"login_username={username}&login_password={password}"

        response = self.__do_a_http_request__(login_url)

        if response.get('result') == 'FAIL':
            raise Exception("Login failed. Please check your credentials.")
        
        self.rsid = response['data']['data'].get('RSID')
        
        return
    
    def load_site_data(self) -> pd.DataFrame:
        if not self.rsid:
            raise Exception("You must login first to obtain an RSID.")

        site_id_url = f"{self.base_url}/srv.SrvClienteJSON.buscarUbicacion+"
        site_id_url += f"RSID={self.rsid}&nomlinea=&site=&idubicacion="

        response = self.__do_a_http_request__(site_id_url)

        if response.get('result') == 'FAIL':
            raise Exception("Failed to retrieve site ID. Please check the protocol and site number. " + "Error: " + str(response.get('data')))

        return pd.DataFrame(response.get('data').get('data')).T
    
    def create_return(self, id_ubicacion: str, referencia: str, 
            retiradde: str, retirahta: str, entregadde: str, entregahta: str,
            obsOper: str, tipomaterial: int, 
            cajas: int) -> dict:
        
        if not self.rsid:
            raise Exception("You must login first to obtain an RSID.")
        
        temperatura = 1 # Valor fijo para temperatura 'Ambiente' según la documentación de TANET

        order_url = f"{self.base_url}/srv.SrvClienteJSON.crearRecoleccion+"
        order_url += f"RSID={self.rsid}&idubicacion={id_ubicacion}&referencia={referencia}"
        order_url += f"&retiradde={retiradde}&retirahta={retirahta}"
        order_url +=  f"&entregadde={entregadde}&entregahta={entregahta}"
        order_url += f"&obsOper={obsOper}&tipomaterial={str(tipomaterial)}&temperatura={str(temperatura)}"
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