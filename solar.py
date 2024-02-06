""" Helper script for downloading data from sems portal and ploting graphs """
import json
import requests
from config import args

API_URL = "https://www.semsportal.com/api/"

headers = {"Token": "{'version': 'v2.1.0', 'client': 'ios', 'language': 'en'}"}


def login() -> None:
    payload = {
        "account": args.get("gw_account"),
        "pwd": args.get("gw_password"),
    }
    url = API_URL + "v1/Common/CrossLogin"
    response = requests.post(url, headers=headers, json=payload)
    Token = response.json().get("data")
    headers["Token"] = json.dumps(Token)


def main() -> None:
    login()


if __name__ == "__main__":
    main()
