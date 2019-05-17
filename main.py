#!/usr/bin/env python

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import re

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1Kc5kwDV1PXI8wXA3qcvlyM5Aip5auOQP_cMicYMQBi4'
COURSES_RANGE = 'Run01-Input!B2:AP3'
MOOD_EFFECT_RANGE = 'Mood effects!A2:I16'


def get_course_plan(sheet):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=COURSES_RANGE).execute()
    values = result.get('values', [])
    assert len(values) == 2
    course_plan = []
    morning, afternoon = values[0], values[1]
    for i in range(len(morning)):
        course_plan.append((morning[i], afternoon[i]))
    return course_plan

def get_mood_effects(sheet):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=MOOD_EFFECT_RANGE).execute()
    values = result.get('values', [])

    mood_effects = {}
    moods = values[0][1:]
    for mood in moods:
        mood_effects[mood] = {}
    for row in values[1:]:
        skill = row[0]
        for i, cell in enumerate(row[1:]):
            mood = moods[i]
            mood_effects[mood][skill] = int(cell)
    return mood_effects

def write_courses(sheet, morning, afternoon):
    body = {"values": [morning, afternoon]}
    sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                          range=COURSES_RANGE,
                          body=body, valueInputOption="RAW").execute()


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    # course_plan = get_course_plan(sheet)
    # mood_effects = get_mood_effects(sheet)
    # print(course_plan)
    run01 = open("run01.html", "r", encoding = "ISO-8859-1").read()
    weeks = run01.split("<hr>")[1:]
    # print(weeks[4])
    study_pattern = (r"Studied (?P<morning>.*) in the morning\.\)<br>.*"
                     r"\(Studied (?P<afternoon>.*) in the afternoon")
    morning, afternoon = [], []
    for week in weeks[1:]:
        matches = re.search(study_pattern, week)
        if matches:
            morning.append(matches.group("morning"))
            afternoon.append(matches.group("afternoon"))
        # else:
        #     print(week)
    # print(len(morning))
    # print(len(weeks))
    write_courses(sheet, morning, afternoon)
    # print(morning)


if __name__ == '__main__':
    main()
