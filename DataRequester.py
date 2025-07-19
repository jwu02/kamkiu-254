import requests
import json
import os
from dotenv import load_dotenv

load_dotenv()

class DataRequester:

    KAMKIU_BASE_API_URL = os.getenv("KAMKIU_BASE_API_URL")

    def request_general(self, url: str, data: dict) -> list:
        endpoint = f"{self.KAMKIU_BASE_API_URL}/{url}"
        headers = {
            "Content-Type": "application/json"
        }

        try:
            response = requests.post(endpoint, data=json.dumps(data), headers=headers)

            # Try to parse the response as JSON
            try:
                json_data = response.json()  # This converts the response to Python dict
                # print("Response JSON:", json.dumps(json_data, indent=2, ensure_ascii=False))
            except ValueError as e:
                print("Response is not valid JSON:", response.text)
            
            # # Print the response details
            # print("Status Code:", response.status_code)
            # print("Response:", response.text)

            return json_data['data']
            
        except requests.exceptions.RequestException as e:
            print("Error making the request:", e)
            return []
    
    def request_shipment_details(self) -> list:
        url = "493"
        data = {
            "project": "254",
            "shipment_date_start": "2025-07-17",
            "model_code": "KAP-7461中板-A76-50,KAP-7457上U-A76-50,KAP-7487下U-A76-50"
        }

        return self.request_general(url, data)
    
    def request_ageing_qrcode(self, model_code_list: list, extrusion_batch_code_list: list) -> list:
        url = "507"
        data = {
            "model_code": ",".join(model_code_list),
            "extrusion_batch_code": ",".join(extrusion_batch_code_list),
        }

        return self.request_general(url, data)

    def request_process_card_qrcode(self, model_code_list: list, extrusion_batch_code_list: list) -> list:
        url = "pz230"
        data = {
            "model_code": ",".join(model_code_list),
            "extrusion_batch_code": ",".join(extrusion_batch_code_list),
        }

        return self.request_general(url, data)

    def request_mechanical_properties(self, model_code_list: list, smelting_furnace_code_list=[], ageing_furnace_code_list=[]) -> list:
        url = "wtdmx"
        data = {
            "model_code_list": ",".join(model_code_list),
            "billet_furnace_code": ",".join(smelting_furnace_code_list),
            "ageing_furnace_code": ",".join(ageing_furnace_code_list),
        }

        return self.request_general(url, data)

    def request_chemical_composition(self, smelt_lot_list: list[str]) -> list:
        url = "043"
        data = {
            "billet_furnace_code": ",".join(smelt_lot_list),
        }

        return self.request_general(url, data)
    
    def request_test_commission_form(self, model_code_list: list, billet_furnace_code_list=[], ageing_furnace_code_list=[]) -> list:
        url = "wtd1"
        data = {
            "model_code_list": ",".join(model_code_list),
            "billet_furnace_code": ",".join(billet_furnace_code_list),
            "ageing_furnace_code": ",".join(ageing_furnace_code_list),
        }

        return self.request_general(url, data)