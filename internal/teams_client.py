from typing import Dict

import json
from datetime import datetime, timezone
import uuid

import requests
from requests import Response
import time

from internal.exceptions import ResourceNotFoundException, UnknownAPIException

class TeamsClient:

    def __init__(self, search_token: str, message_token: str) -> None:
        self.search_token = search_token
        self.message_token = message_token
        
    
    def get_user_by_email(self, email: str) -> Dict:
        
        url = f"https://teams.microsoft.com/api/mt/apac/beta/users/{email}/externalsearchv3?includeTFLUsers=true"

        payload = {}
        headers = {
        'x-ms-client-version': '1415/24090101421', 
        'authorization': f'Bearer {self.search_token}',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'content-type': 'application/json;charset=UTF-8',
        'x-ms-migration': 'True',
        'accept': 'application/json',
        'x-ms-request-id': '',
        'Referer': 'https://teams.microsoft.com/v2/worker/precompiled-web-worker-b686ae686e2a6f80.js',
        'x-ms-client-caller': 'powerbarSearchFederatedUser'
        }

        response = self.send_request("GET", url, headers, payload)
        
        if response.status_code != 200:
            raise UnknownAPIException(f"error {response.status_code} fetching user from api: " + response.text)
        
        response_json = response.json()
        if len(response_json) == 0:
            raise ResourceNotFoundException(f"User with email {email} not found")
        
        return response_json[0]
    

    def get_chat(self, chat_id) -> Dict:

        url = f"https://teams.microsoft.com/api/chatsvc/apac/v1/users/ME/conversations/{chat_id}?view=msnp24Equivalent"

        payload = {}
        headers = {
        'behavioroverride': 'redirectAs404',
        'x-ms-test-user': 'False',
        'authorization': f'Bearer {self.message_token}',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'clientinfo': 'os=linux; osVer=undefined; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=1415/24090101421; utcOffset=+05:00; timezone=Asia/Karachi',
        'x-ms-migration': 'True',
        'x-ms-request-priority': '10',
        'Referer': 'https://teams.microsoft.com/v2/worker/precompiled-web-worker-b686ae686e2a6f80.js'
        }

        response = self.send_request("GET", url, headers, payload)

        if response.status_code == 404:
            raise ResourceNotFoundException(f"Chat with given id {chat_id} not found")

        if response.status_code != 200:
            raise UnknownAPIException(f"Error {response.status_code} getting chat: {response.text}")
        
        return response.json()




    def create_chat(self, auth_user_id: str, user_id: str) -> None:
        url = "https://teams.microsoft.com/api/chatsvc/apac/v1/threads"

        payload = json.dumps({
        "members": [
            {
            "id": f"8:orgid:{auth_user_id}",
            "role": "Admin"
            },
            {
            "id": f"8:orgid:{user_id}",
            "role": "Admin"
            }
        ],
        "properties": {
            "threadType": "chat",
            "fixedRoster": True,
            "uniquerosterthread": True
        }
        })
        headers = {
        'behavioroverride': 'redirectAs404',
        'x-ms-test-user': 'False',
        'authorization': f'Bearer {self.message_token}',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'content-type': 'application/json',
        'clientinfo': 'os=linux; osVer=undefined; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=1415/24090101421; utcOffset=+05:00; timezone=Asia/Karachi',
        'x-ms-migration': 'True',
        'Referer': 'https://teams.microsoft.com/v2/worker/precompiled-web-worker-b686ae686e2a6f80.js'
        }

        response = self.send_request("POST", url, headers, payload)

        if response.status_code != 201:
            raise UnknownAPIException(f"error {response.status_code} creating chat: {response.text}")




    def send_message(self, chat_id: str, message_text: str) -> None:

        url = f"https://teams.microsoft.com/api/chatsvc/apac/v1/users/ME/conversations/{chat_id}/messages"

        current_time = str(datetime.now(timezone.utc).isoformat())[:23] + "Z"

        clientmessageid = str(uuid.uuid4().int)[:19]

        payload = json.dumps({
        "id": "-1",
        "type": "Message",
        "conversationid": chat_id,
        "conversationLink": f"blah/{chat_id}",
        "from": "8:orgid:9ad57a72-3174-4f4d-9957-df14ff7bd471",
        "composetime": current_time,
        "originalarrivaltime": current_time,
        "content": f"<p>{message_text}</p>",
        "messagetype": "RichText/Html",
        "contenttype": "Text",
        "imdisplayname": "Marium Asif",
        "clientmessageid": clientmessageid,
        "callId": "",
        "state": 0,
        "version": "0",
        "amsreferences": [],
        "properties": {
            "importance": "",
            "subject": "",
            "title": "",
            "cards": "[]",
            "links": "[]",
            "mentions": "[]",
            "onbehalfof": None,
            "files": "[]",
            "policyViolation": None
        },
        "crossPostChannels": []
        })
        headers = {
        'behavioroverride': 'redirectAs404',
        'x-ms-test-user': 'False',
        'authorization': f'Bearer {self.message_token}',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'content-type': 'application/json',
        'clientinfo': 'os=linux; osVer=undefined; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=1415/24090101421; utcOffset=+05:00; timezone=Asia/Karachi',
        'x-ms-migration': 'True',
        'x-ms-request-priority': '0',
        'Referer': 'https://teams.microsoft.com/v2/worker/precompiled-web-worker-b686ae686e2a6f80.js'
        }

        response = self.send_request("POST", url, headers, payload)

        if response.status_code != 201:
            raise UnknownAPIException(f"error {response.status_code} sending message: {response.text}")
        

    def send_request(self, method: str, url: str, headers: Dict, payload: Dict) -> Response:
        try:
            return requests.request(method, url, headers=headers, data=payload)
        except (requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError):
            print("Connection error. Retrying request in 5 seconds...")
            time.sleep(5)
            return self.send_request(method, url, headers, payload)