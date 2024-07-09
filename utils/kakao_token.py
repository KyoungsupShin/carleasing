import requests
import json
from datetime import datetime, timedelta
import time
import schedule
from PyKakao import Message

# Constants
REST_API_KEY = "7205d2e339b1e499f44ca77518252fb7"
REDIRECT_URI = "https://localhost:5000"
TOKEN_FILE = "../data/tokens.json"  # File to store token data

# Global variables
access_token = None
token_expiry = None

def save_tokens(access_token, refresh_token):
    """ Save current access token to a JSON file """
    data = {
        "access_token": access_token,
        "expiry_time": token_expiry.isoformat(),
        "refresh_token" : refresh_token
    }
    with open(TOKEN_FILE, 'w') as f:
        json.dump(data, f)

def load_tokens():
    """ Load access token from JSON file if it exists and not expired """
    global access_token, token_expiry, refresh_token
    try:
        with open(TOKEN_FILE, 'r') as f:
            data = json.load(f)
            access_token = data.get('access_token')
            expiry_time = data.get('expiry_time')
            refresh_token= data.get('refresh_token')
            if expiry_time:
                token_expiry = datetime.fromisoformat(expiry_time)
                if token_expiry > datetime.now():
                    print(f"Loaded access token valid until: {data} {token_expiry}")
                    return True
    except FileNotFoundError:
        pass  # File doesn't exist yet, so ignore
    return False

def get_authorization_code():
    API = Message(service_key = REST_API_KEY)
    auth_url = API.get_url_for_generating_code()
    print(auth_url)
    authorization_code = input("Enter the authorization code: ")
    return authorization_code.split('=')[-1]

def exchange_code_for_tokens(authorization_code):
    url = "https://kauth.kakao.com/oauth/token"
    data = {
        "grant_type": "authorization_code",
        "client_id": REST_API_KEY,
        "redirect_uri": REDIRECT_URI,
        "code": authorization_code
    }
    response = requests.post(url, data=data)
    tokens = response.json()
    return tokens

def refresh_kakao_token():
    global access_token, refresh_token, token_expiry

    url = "https://kauth.kakao.com/oauth/token"
    data = {
        "grant_type": "refresh_token",
        "client_id": REST_API_KEY,
        "refresh_token": refresh_token
    }

    resp = requests.post(url, data=data)
    new_tokens = resp.json()

    if "access_token" in new_tokens:
        access_token = new_tokens["access_token"]
        token_expiry = datetime.now() + timedelta(seconds=new_tokens["expires_in"])
        print(f"New access token valid until: {new_tokens} {token_expiry}")
        save_tokens(access_token, refresh_token)
    else:
        print("Failed to refresh token:")
        print(new_tokens)

def main():
    global access_token, token_expiry, refresh_token
    if not load_tokens():
        authorization_code = get_authorization_code()
        tokens = exchange_code_for_tokens(authorization_code)
        if "access_token" in tokens:
            access_token = tokens["access_token"]
            refresh_token = tokens["refresh_token"]

            token_expiry = datetime.now() + timedelta(seconds=tokens["expires_in"])
            print(f"Initial access token valid until: {tokens} {token_expiry}")
            save_tokens(access_token, refresh_token)
        else:
            print("Failed to obtain tokens:")
            print(tokens)

    schedule.every(1).hour.do(refresh_kakao_token)

    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    main()
