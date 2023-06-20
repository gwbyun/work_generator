import datetime
import os
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/calendar']

creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('calendar', 'v3', credentials=creds)




#colorId = 1: 밝은 빨간색
#colorId = 2: 밝은 주황색
#colorId = 3: 밝은 노란색
#colorId = 4: 밝은 연두색
#colorId = 5: 밝은 초록색
#colorId = 6: 밝은 청록색
#colorId = 7: 밝은 파란색
#colorId = 8: 밝은 자주색
#colorId = 9: 밝은 분홍색
#colorId = 10: 밝은 회색
#colorId = 11: 밝은 갈색
#위와 같이 colorId 값을 설정하


def event_A(year, month, day):
    start_date = datetime.datetime(year, month, day, 9, 0, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 19, 0, 0).isoformat()
    event = {
        'summary': '오전 근무',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
        'colorId': 2,  # 색상 아이디 (1~11 사이의 정수)

    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_P(year, month, day):
    start_date = datetime.datetime(year, month, day, 12, 30, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 22, 0, 0).isoformat()
    event = {
        'summary': '오후 근무',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_M(year, month, day):
    start_date = datetime.datetime(year, month, day, 10, 0, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 20, 0, 0).isoformat()
    event = {
        'summary': 'M 근무',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_N(year, month, day):
    start_date = datetime.datetime(year, month, day, 21, 0, 0).isoformat()
    end_date = datetime.datetime(year, month, day+1, 9, 30, 0).isoformat()
    event = {
        'summary': '당직 근무',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_O(year, month, day):
    start_date = datetime.datetime(year, month, day, 0, 0, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 23, 59, 0).isoformat()
    event = {
        'summary': '휴무',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_AH(year, month, day):
    start_date = datetime.datetime(year, month, day, 9, 0, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 13, 30, 0).isoformat()
    event = {
        'summary': '오전 반차',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_HA(year, month, day):
    start_date = datetime.datetime(year, month, day, 14, 30, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 19, 00, 0).isoformat()
    event = {
        'summary': '오전 반차',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_PH(year, month, day):
    start_date = datetime.datetime(year, month, day, 12, 30, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 17, 0, 0).isoformat()
    event = {
        'summary': '오후 반차',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_HP(year, month, day):
    start_date = datetime.datetime(year, month, day, 17, 30, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 22, 0, 0).isoformat()
    event = {
        'summary': '오전 근무',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))
def event_H(year, month, day):
    start_date = datetime.datetime(year, month, day, 0, 0, 0).isoformat()
    end_date = datetime.datetime(year, month, day, 19, 0, 0).isoformat()
    event = {
        'summary': '휴가',
        'start': {
            'dateTime': start_date,
            'timeZone': 'Asia/Seoul',
        },
        'end': {
            'dateTime': end_date,
            'timeZone': 'Asia/Seoul',
        },
       
        'visibility': 'private',
    }

    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created:', event.get('htmlLink'))



def event_select(year, month,day,work):
    
    if work == 'A' or work=='a':
        event_A(year,month,day)
    elif work == 'P' or work=='p':
        event_P(year,month,day)
    elif work == 'M' or work=='m':
        event_M(year,month,day)
    elif work == 'N' or work=='n':
        event_N(year,month,day)
    elif work == 'O' or work=='o':
        event_O(year,month,day)
    elif work == 'AH' or work=='ah':
        event_AH(year,month,day)
    elif work == 'HA' or work=='ha':
        event_HA(year,month,day)
    elif work == 'PH' or work=='ph':
        event_PH(year,month,day)
    elif work == 'HP' or work=='hp':
        event_HP(year,month,day)
    elif work == 'H' or work=='h':
        event_H(year,month,day)
    else:
        print("유효한 입력이 아닙니다.")
