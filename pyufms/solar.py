""" Helper script for downloading data from sems portal and ploting graphs """
import json
from datetime import datetime
from pathlib import Path

import requests
from pyufms.config import ARGS, INVERTERS_INFO
from pyufms.inverters import Inverter

API_URL = "https://www.semsportal.com/api/"

headers = {"Token": "{'version': 'v2.1.0', 'client': 'ios', 'language': 'en'}"}


def login() -> None:
    payload = {
        "account": ARGS.get("gw_account"),
        "pwd": ARGS.get("gw_password"),
    }
    url = API_URL + "v1/Common/CrossLogin"
    response = requests.post(url, headers=headers, json=payload)
    Token = response.json().get("data")
    headers["Token"] = json.dumps(Token)


def get_excel_payload(start: datetime, end: datetime, station_id: str, sn: str) -> dict:
    return {
        "tm_content": {
            "qry_time_start": str(start),
            "qry_time_end": str(end),
            "pws_historys": [
                {
                    "id": station_id,
                    "inverters": [{"sn": sn}],
                }
            ],
            "targets": [{"target_key": "Pac", "target_index": 18}],
        }
    }


def get_excel_qry_key(start: datetime, end: datetime, inverter: Inverter) -> str:
    inverter_info = INVERTERS_INFO.get(inverter)
    if inverter_info is None:
        raise Exception("Error getting inverter's info: " + inverter.name)

    station_id = inverter_info.get("station_id")
    if station_id is None:
        raise Exception("Error getting inverter's station_id: " + inverter.name)

    sn = inverter_info.get("sn")
    if sn is None:
        raise Exception("Error getting inverter's sn: " + inverter.name)

    url = API_URL + "HistoryData/ExportExcelStationHistoryData"
    payload = get_excel_payload(start, end, station_id, sn)
    response = requests.post(url, headers=headers, json=payload)
    data = response.json().get("data")
    qry_key = data.get("qry_key")
    return qry_key


def get_excel_file_path(start: datetime, end: datetime, inverter: Inverter) -> str:
    qry_key = get_excel_qry_key(start, end, inverter)
    url = API_URL + "HistoryData/GetStationHistoryDataFilePath"
    payload = {"file_id": qry_key}
    response = requests.post(url, headers=headers, json=payload)
    data = response.json().get("data")
    file_path = data.get("file_path")
    return file_path


def get_excel(start: datetime, end: datetime, inverter: Inverter) -> None:
    file_path = get_excel_file_path(start, end, inverter)
    response = requests.get(file_path)
    response.raise_for_status()
    path_to_save = Path(start.strftime("%d_%m_%Y-") + inverter.name + ".xls")
    path_to_save.write_bytes(response.content)


def main() -> None:
    start = datetime(2024, 1, 5)
    end = datetime(2024, 1, 6)
    inverter = Inverter.S1_BL11
    login()
    try:
        get_excel(start, end, inverter)
    except Exception as e:
        print(e)
