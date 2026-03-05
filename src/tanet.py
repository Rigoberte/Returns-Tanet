import requests
import pandas as pd

class Tanet:
    def __init__(self) -> None:
        self.base_url: str = "https://sgi.tanet.com.ar/sgi"
        self.rsid: str = ""

    def login(self, username: str, password: str) -> None:
        login_url = f"{self.base_url}/usr.usrJSON.login+login_username={username}&login_password={password}"

        response = self.__do_a_http_request__(login_url)

        if response.get('result') == 'FAIL':
            raise Exception("Login failed. Please check your credentials.")
        
        self.rsid = response['data']['data'].get('RSID')
        
        return
    
    def load_site_data(self) -> pd.DataFrame:
        if not self.rsid:
            raise Exception("You must login first to obtain an RSID.")

        site_id_url = f"{self.base_url}/srv.SrvClienteJSON.buscarUbicacion+RSID={self.rsid}&nomlinea=&site=&idubicacion="

        response = self.__do_a_http_request__(site_id_url)

        if response.get('result') == 'FAIL':
            raise Exception("Failed to retrieve site ID. Please check the protocol and site number.")

        return pd.DataFrame(response.get('data').get('data')).T

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
            return fail_dict
        
        success_dict = {"result": "OK", "data": response_data}
        return success_dict