from src.tanet import Tanet
import pandas as pd
from difflib import SequenceMatcher

def get_site_data() -> pd.DataFrame:
    tanet = Tanet()
    try:
        tanet.login("", "")
    except Exception as e:
        print(f"Login error: {e}")
        return pd.DataFrame()

    site_id_df = tanet.load_site_data()
    
    path = r"C:\Users\inaki.costa\Downloads\site_id_data.xlsx"
    site_id_df.to_excel(path, index=False)

    return site_id_df

def load_user_request() -> pd.DataFrame:
    user_request = pd.DataFrame([
        {"protocol": "MK-6482-011 LOCAL", "site_number": "800"},
        {"protocol": "I8F-MC-GPGN", "site_number": "921"},
        {"protocol": "MK-2870-012", "site_number": "0501"}
    ])

    user_request['ORDER_ID'] = user_request.index + 1

    return user_request

def process_user_request(site_data_df: pd.DataFrame, order_id: str, protocol: str, site_number: str) -> pd.DataFrame:
    results = pd.DataFrame()
    similarity_threshold = 0.8
    
    protocol = protocol.upper().strip()
    site_number = site_number.upper().strip()


    for _, site_row in site_data_df.iterrows():
        site_protocol = site_row['nomlinea'].upper().strip()
        site_site_number = site_row['site'].upper().strip()

        similarity = SequenceMatcher(None, protocol, site_protocol).ratio()
        
        if similarity >= similarity_threshold and site_number in site_site_number:
            site_row['ORDER_ID'] = order_id
            results = pd.concat([results, site_row.to_frame().T], ignore_index=True)

    if results is not None and not results.empty:
        return results
    else:
        return pd.DataFrame()

def main():
    site_data_df = get_site_data()
    
    user_request = load_user_request()

    result = pd.DataFrame()
    for index, row in user_request.iterrows():
        processed_data = process_user_request(site_data_df, row['ORDER_ID'], row['protocol'], row['site_number'])
        result = pd.concat([result, processed_data], ignore_index=True)

    result_path = r"C:\Users\inaki.costa\Downloads\result.xlsx"

    result.to_excel(result_path, index=False)

if __name__ == "__main__":
    main()