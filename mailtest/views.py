import base64
import json
import re

from bs4 import BeautifulSoup
from ics import Event, Calendar
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition, ContentId
from django.shortcuts import render, redirect
import requests
from decouple import config
from django.http import HttpResponse
import datetime
import pytz

CLIENT_ID = config('CLIENT_ID')
CLIENT_SECRET = config('CLIENT_SECRET')
REDIRECT_URI = config('REDIRECT_URI')
SENDGRID_KEY = config('SENDGRID_KEY')

AUTH_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
SCOPE = 'https://graph.microsoft.com/Calendars.ReadWrite'

access_token = None


def home(request):
    auth_url = f'{AUTH_URL}?client_id={CLIENT_ID}&response_type=code&redirect_uri={REDIRECT_URI}&scope={SCOPE}'
    return redirect(auth_url)


def callback(request):
    global access_token

    if 'code' in request.GET:
        code = request.GET['code']
        token_data = {
            'grant_type': 'authorization_code',
            'client_id': CLIENT_ID,
            'client_secret': CLIENT_SECRET,
            'code': code,
            'redirect_uri': REDIRECT_URI,
            'scope': SCOPE,
        }
        token_response = requests.post(TOKEN_URL, data=token_data)
        token_data = token_response.json()

        if 'access_token' in token_data:
            access_token = token_data['access_token']
            schedule_event_ics(request)
            return HttpResponse('Event sent successfully..!')
        else:
            return HttpResponse('Error getting access token.')
    else:
        return HttpResponse('Permission denied. Please grant access to continue.')


def schedule_event_ics(request):
    global access_token

    if access_token:
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        start_time = datetime.datetime(2023, 9, 12, 13, 0, 0, tzinfo=pytz.timezone('Asia/Kolkata'))
        end_time = datetime.datetime(2023, 9, 12, 14, 0, 0, tzinfo=pytz.timezone('Asia/Kolkata'))
        subject = 'Meeting with Team'
        attendees = [{'emailAddress': {'address': 'vijay.khade@humancloud.co.in'}},
                     {'emailAddress': {'address': 'vjykhade@gmail.com'}}]

        # Online meeting settings
        online_meeting = {
            "isOnlineMeeting": True,
            "onlineMeetingProvider": "teamsForBusiness"
        }

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
            'attendees': attendees,
            **online_meeting
        }

        json_data = json.dumps(event_data)
        cal_url = f'https://graph.microsoft.com/v1.0/me/calendar'
        cal_response = requests.get(cal_url, headers=headers)
        cal_response_data = cal_response.json()
        calendar_id = None
        if cal_response.status_code == 200:
            calendar_id = cal_response_data.get('id')

        url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events'
        try:
            response = requests.post(url, headers=headers, data=json_data)
            response_data = response.json()
            if response.status_code == 201:
                event_id = response_data.get('id')
                event_link_url = f'https://graph.microsoft.com/v1.0/me/events/{event_id}/'
                event_link_response = requests.get(event_link_url, headers=headers)
                event_link_data = event_link_response.json()
                event_link = event_link_data.get('joinUrl')
                if not event_link:
                    event_link = event_link_data.get('onlineMeeting', {}).get('joinUrl')

                event_body = event_link_data.get('body', {}).get('content')

                soup = BeautifulSoup(event_body, 'html.parser')
                text = soup.get_text()
                meeting_id = re.search(r'Meeting ID: (\d{3} \d{3} \d{3} \d{3})', text).group(1)
                passcode = re.search(r'Passcode: (\w+)', text).group(1)
                meeting_options = soup.select_one('a[href*="meetingOptions"]')['href']
                learn_more = soup.select_one('a[href*="JoinTeamsMeeting"]')['href']
                html_content = (
                    f'<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><meta name="Generator" content="Microsoft Exchange Server"></head><body><font size="2"><span style="font-size:11pt;"><h2>Microsoft Teams meeting</h2><div class="PlainText">.........................................................................................................................................<br><b>Join on your computer, mobile app or room device</b><br><a href={event_link}>'
                    f'Click here to join the meeting</a>'
                    f'<br><br>Meeting ID: {meeting_id}'
                    f'<br>Passcode: {passcode}<br><br>'
                    f'If you need a local number, get one here. And if you\'ve forgotten the dial-in PIN, you can reset it.<br><br>'
                    f'<a href={learn_more}>Learn More</a> |'
                    f'<a href={meeting_options}>Meeting options</a>'
                    f'<br>.........................................................................................................................................<br>'
                    f'</div></span></font></body></html>')
                # Generating ICS file
                cal = Calendar()
                event = Event()
                event.name = subject
                event.begin = start_time
                event.end = end_time
                event.description = html_content
                event.organizer = 'support@teamcast.ai'
                event.attendees = [attendee['emailAddress']['address'] for attendee in attendees]
                event.url = event_link
                cal.events.add(event)
                ics_data = cal.serialize()
                ics_bytes = base64.b64encode(ics_data.encode('utf-8')).decode()

                try:
                    sg = SendGridAPIClient(SENDGRID_KEY)

                    message = Mail(
                        from_email='support@teamcast.ai',  # Replace with your email address
                        subject='Meeting Invitation: ' + subject,
                        html_content=html_content,
                        to_emails=[attendee['emailAddress']['address'] for attendee in attendees]
                    )
                    attachment = Attachment()
                    attachment.file_content = FileContent(ics_bytes)
                    attachment.file_type = FileType('application/ics')
                    attachment.file_name = FileName('invite.ics')
                    attachment.disposition = Disposition('attachment')
                    attachment.content_id = ContentId('calendar_invite')
                    # message.attachment = attachment

                    pdf_file_path = 'D:\\Dhaval_Zaveri_CV.pdf'
                    pdf_attachment = Attachment()
                    with open(pdf_file_path, 'rb') as pdf_file:
                        pdf_data = pdf_file.read()
                    pdf_attachment.file_content = FileContent(base64.b64encode(pdf_data).decode())
                    pdf_attachment.file_type = FileType('application/pdf')
                    pdf_attachment.file_name = FileName('Dhaval_Zaveri_CV.pdf')
                    pdf_attachment.disposition = Disposition('attachment')
                    pdf_attachment.content_id = ContentId('cv_attachment')

                    message.attachment = [attachment, pdf_attachment]

                    res = sg.send(message)
                    if res.status_code == 202:
                        return HttpResponse(f'Meeting invitation sent successfully. Event link: {event_link}')
                    else:
                        return HttpResponse(f'Error sending meeting invitation: {res.status_code}')
                except Exception as e:
                    return HttpResponse(f'Error sending meeting invitation: {str(e)}')
            else:
                return HttpResponse(f'Error scheduling meeting: {response.status_code}, {response_data}')
        except Exception as e:
            return HttpResponse(f'Error scheduling meeting: {str(e)}')
    else:
        return HttpResponse('Access token not available. Please authenticate first.')
