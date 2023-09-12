import requests
import datetime
import pytz

# Client ID and secret from registering your app in Azure
client_id = '5b1c5653-d52f-4b70-af57-07375efbf9f0'
client_secret = 'jHv8Q~0mcHHqy7KQy_6lHLp2lcE6hto2Eyd5ocL4'
redirect_uri = 'http://localhost:8000/oauth/callback'

# API endpoint to get an OAuth token
token_url = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token'
auth_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'

# Scope required to read and write calendars
scope = 'https://graph.microsoft.com/.default'

#scope = 'https://graph.microsoft.com/Calendars.ReadWrite offline_access'

# Redirect the user to Microsoft login for authentication
auth_params = {
    'client_id': client_id,
    'response_type': 'code',
    'redirect_uri': redirect_uri,
    'scope': scope,
}

auth_redirect_url = requests.get(auth_url, params=auth_params).url

# POST request body to get a token
data = {
    'grant_type': 'client_credentials',
    'client_id': client_id,
    'client_secret': client_secret,
    'scope': scope
}

# Request an access token
response = requests.post(token_url, data=data)
token_data = response.json()

if 'access_token' in token_data:
    access_token = token_data['access_token']

    # Use the token to get the user's calendars
    headers = {
        'Authorization': 'Bearer {}'.format(access_token),
        'Content-Type': 'application/json'
    }

    # Get the user's calendars
    calendars_url = 'https://graph.microsoft.com/v1.0/me/calendars'
    calendars_response = requests.get(calendars_url, headers=headers)

    if calendars_response.status_code == 200:
        calendars = calendars_response.json().get('value', [])

        if calendars:
            # Assuming you want to use the first calendar found
            calendar_id = calendars[0]['id']

            # Event details
            start_time = datetime.datetime(2023, 9, 10, 8, 0, 0, tzinfo=pytz.timezone('Asia/Kolkata'))
            end_time = datetime.datetime(2023, 9, 10, 10, 0, 0, tzinfo=pytz.timezone('Asia/Kolkata'))
            subject = 'Meeting with Team'
            attendees = ['vijay.khade@humancloud.co.in', 'vikram.chandel@humancloud.co.in']

            # POST body to create an event
            event_data = {
                'subject': subject,
                'start': {
                    'dateTime': start_time.isoformat(),
                    'timeZone': 'Asia/Kolkata'
                },
                'end': {
                    'dateTime': end_time.isoformat(),
                    'timeZone': 'Asia/Kolkata'
                },
                'attendees': [{'emailAddress': {'address': email}} for email in attendees],
            }

            # URL to create an event on the user's calendar
            url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events'

            # POST request to create the event
            response = requests.post(url, headers=headers, json=event_data)

            if response.status_code == 201:
                print('Meeting scheduled successfully!')
            else:
                print('Error scheduling meeting:', response.status_code, response.text)
        else:
            print('No calendars found for the user.')
    else:
        print('Error getting user calendars:', calendars_response.status_code, calendars_response.text)
else:
    print('Error getting access token:', token_data.get('error', 'Unknown error'))
