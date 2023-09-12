import requests
import datetime
import pytz

# Client ID and secret from registering your app in Azure
client_id = '1125a0d3-9e0d-4850-8cdc-6306a849f7f3'
client_secret = 'Vi38Q~FHHwTxTcXrgNpuOOaEE3ZnJzqwJnliodlR'

# API endpoint to get an OAuth token
token_url = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token'

# Scope required to read and write calendars
scope = 'https://graph.microsoft.com/Calendars.ReadWrite'

# POST request body to get a token
data = {
    'grant_type': 'client_credentials',
    'client_id': client_id,
    'client_secret': client_secret,
    'scope': scope
}

access_to = ""
# Request an access token
try:
    response = requests.post(token_url, data=data)
    if 'error' in response.json():
        print(f'Errror -->', response.json()['error'])
    else:
        access_token = response.json()['access_token']
        access_to = access_token
except Exception as e:
    print(f'Error -->', e)

# Use the token to schedule an event
headers = {
    'Authorization': 'Bearer {}'.format(access_to),
    'Content-Type': 'application/json'
}

# Event details
start = datetime.datetime(2023, 9, 10, 8, 0, 0, tzinfo=pytz.timezone('Asia/Kolkata'))
end = datetime.datetime(2023, 9, 10, 10, 0, 0, tzinfo=pytz.timezone('Asia/Kolkata'))
subject = 'Meeting with Team'
attendees = ['vijay.khade@humancloud.co.in', 'vikram.chandel@humancloud.co.in']

# POST body to create event
data = {
    'subject': subject,
    'start': {'dateTime': start.isoformat(), 'timeZone': 'Asia/Kolkata'},
    'end': {'dateTime': end.isoformat(), 'timeZone': 'Asia/Kolkata'},
    'attendees': [{'emailAddress': {'address': email}} for email in attendees],
    'allowNewTimeProposals': 'true'
}

# URL to create event
url = 'https://graph.microsoft.com/v1.0/me/events'

# POST request to create event
requests.post(url, headers=headers, json=data)
