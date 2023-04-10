import datetime
import json

import msal
import pandas as pd
import requests

cred = json.load(open("namasys_credentials.json"))

CLIENT_ID = cred['CLIENT_ID']
CLIENT_SECRET = cred['CLIENT_SECRET']
AUTHORITY_URL = cred['AUTHORITY_URL']
REDIRECT_URI = cred['REDIRECT_URI']
USERNAME = cred['USERNAME']
PASSWORD = cred['PASSWORD']
SCOPES = ["https://graph.microsoft.com/.default", ]
REFRESH_TOKEN = cred['REFRESH_TOKEN']

DATAFRAME_LIST = []


def get_start_and_end() -> str:

    today = datetime.datetime.now()-datetime.timedelta(60)
    start = today.replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + datetime.timedelta(60)

    return start.isoformat() + ".000Z", end.isoformat() + ".000Z"


def get_access_token() -> str:
    """ 
    Extracts access_token 
    Args:

    Returns: access_token(str) 
    get_access_token Example: access_token = get_access_token() 

    """

    data = {'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
            'scope': "https://graph.microsoft.com/.default", 'grant_type': 'client_credentials', }

    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    response = requests.post(url=url, data=data, headers=headers,timeout=None)
    print('access_token:', response.status_code)
    response_json = response.json()
    access_token = response_json['access_token']
    return access_token


def get_msal_access_token() -> str:
    """ 
    Extracts access_token 
    Args:

    Returns: access_token(str) 
    get_access_token Example: access_token = get_msal_access_token() 

    """
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY_URL,)

    result = app.acquire_token_by_username_password(
        username=USERNAME,
        password=PASSWORD,
        scopes=SCOPES
    )

    access_token = result.get("access_token")
    return access_token


def get_token_using_refresh_token() -> str:
    """ 
    Extracts access_token using refresh token.
    Args:

    Returns: access_token(str) 
    get_access_token Example: access_token = get_token_using_refresh_token() 

    """

    app = msal.ConfidentialClientApplication(CLIENT_ID,
                                             authority=AUTHORITY_URL,
                                             client_credential=CLIENT_SECRET)

    result = app.acquire_token_for_client(scopes=SCOPES)

#     refresh_token = result.get("refresh_token")
    token_response = app.acquire_token_by_refresh_token(refresh_token=REFRESH_TOKEN,
                                                        scopes=SCOPES,)
    print("token_response", token_response)
    # Extract new access token from token response
    access_token = token_response['access_token']
    return access_token


def get_groups_list(access_token) -> list:
    """
    Get List of Microsoft groups.
    Args:

    Return: [groups1,group2,...]
    get_groups_list Example: GROUP_LIST = get_groups_list()
    """

    url = f"https://graph.microsoft.com/v1.0/groups"
    headers = {

        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }

    response = requests.get(url=url, headers=headers,timeout=None)
    group_list = list(
        map(lambda data: [data['displayName'], data['id']], response.json()['value']))

    return group_list


def get_calender_view(group_name, group_id, access_token, start_time, end_time) -> None:

    event_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/calendarView?startDateTime={start_time}&endDateTime={end_time}"
    print(event_url)
    headers = {

        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json',
        'Prefer': 'outlook.body-content-type="text"'
    }

    event_response = requests.get(url=event_url, headers=headers)
    print(group_name, group_id, event_response.status_code)

    if event_response.status_code == 200:
        # .to_csv(f"{group_name}.csv",index=False)
        event_dataframe = pd.DataFrame(event_response.json()["value"])
        event_dataframe['Group_Name'] = group_name
        DATAFRAME_LIST.append(event_dataframe)
    else:
        pass

#     return event_response.json()


if __name__ == "__main__":

    start_time, end_time = get_start_and_end()

    access_token = get_token_using_refresh_token()

    GROUP_LIST = get_groups_list(access_token)

    results = list(map(lambda lst: get_calender_view(
        lst[0], lst[1], access_token, start_time, end_time), GROUP_LIST))

    # concatenate all data frames in the list into one
    Event_Dataframe = pd.concat(DATAFRAME_LIST, ignore_index=True)
    Event_Dataframe.to_csv("namasys_Teams_Events.csv", index=False)