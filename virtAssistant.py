from __future__ import print_function
import datetime
import os.path
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

import os
import time
from httplib2 import Response
import pyttsx3
import speech_recognition as sr
import pytz
import subprocess
import wikipedia

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ["january","february","march","april","may","june","july","august","september","october","november","december"]
DAYS = ["monday","tuesday","wednesday","thursday","friday","saturday","sunday"]
DAYS_EXTENSIONS = ["rd","th","st","nd"]
GAME_NAMES = ['Firefox']
GAME_EXES = ["firefox.exe"]
APP_EXECUTIONS = ['open']
GAME_PATHS = ["C:\\Program Files\\Mozilla Firefox\\"]
def speak(text):
    engine = pyttsx3.init()
    engine.setProperty('rate',150)
    voices = engine.getProperty('voices')
    engine.setProperty('voice',voices[1].id)
    engine.setProperty('volume',1.0)
    engine.say(text)
    engine.runAndWait()
    
def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        said = ""
        try:
            said = r.recognize_google(audio)
            print(said)
        except Exception as e:
            print("Exception " + str(e))
    return said

def authenticate_google():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)

    return service


def get_events(day,service):
    # Call the Calendar API
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end = end.astimezone(utc)
    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(),
                                        timeMax=end.isoformat(), singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("-")[0])

            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12)
                start_time = start_time + "pm"

            speak(event['summary'] + " at " + start_time)


def get_date(text):
    text = text.lower()
    today = datetime.date.today()
    tomorrow = datetime.date.today() + datetime.timedelta(days=1)

    if text.count("today") > 0:
        return today
    if text.count("tomorrow") > 0:
        return tomorrow
    day = -1
    month = -1
    day_of_week = -1
    year = today.year

    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        
        elif word in DAYS:
            day_of_week = DAYS.index(word)

        elif word.isdigit():
            day = int(word)
        
        else:
            for ext in DAYS_EXTENSIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass
    if month < today.month and month != -1:
        year = year + 1
    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7
        return today + datetime.timedelta(dif)
    if day != -1:
        return datetime.date(month = month, day = day, year = year)
    if month == -1 or day == -1:
        return None 

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":","-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe",file_name])

def getPersonData(text):
    wordList = text.split()
    if len(wordList) == 4:
        for x in range(0,len(wordList)):
            if x + 3 <= len(wordList) - 1 and wordList[x].lower() == "who" and wordList[x+1].lower() == "is":
                return wordList[x+2] + '' + wordList[x+3]
    else:
        for x in range(0,len(wordList)):
            if x + 4 <= len(wordList) - 1 and wordList[x].lower() == "who" and wordList[x+1].lower() == "is":
                return wordList[x+2] + '' + wordList[x+3] + '' + wordList[x+4]
def openGame(text):
    game_index = 0
    validGame = False
    for game in GAME_NAMES:
        if game in text:
            validGame = True
            game_index = GAME_NAMES.index(game)
            break
    if validGame == True:
        try:
            os.startfile(GAME_PATHS[game_index]+GAME_EXES[game_index])
        except Exception as e:
            speak("Error has occurred in opening game")
    else:
        speak("Could not open game")


WAKE_KEY = "hello Sam"
SERVICE = authenticate_google()
print("Start")

while True:
    print("Listening")
    text = get_audio()

    if text.count(WAKE_KEY) > 0:
        speak("Hello Mahir, what would you like me to do for you?")
        text = get_audio()
        CALENDAR_STRS = ["what do i have", "do i have plans", "am i busy"]
        for phrase in CALENDAR_STRS:
            if phrase in text.lower():
                date = get_date(text)
                if date:
                    get_events(date, SERVICE)
                else:
                    speak("Please Try Again")

        NOTE_STRS = ["make a note", "write this down", "remember this", "type this"]
        for phrase in NOTE_STRS:
            if phrase in text:
                speak("What would you like me to write down? ")
                write_down = get_audio()
                note(write_down)
                speak("I've made a note of that.")

        PERSO_STRS = ["who is"]
        resp = ""
        for phrase in PERSO_STRS:
            if phrase in text:
                person = getPersonData(text)
                wiki = wikipedia.summary(person,sentences=2)
                resp = resp + " " + wiki
                print(resp)
                speak(resp)
        for phrase in APP_EXECUTIONS:
            if phrase in text:
                openGame(text)
    elif text == "goodbye":
        speak("Thank You for using me")
        break
    else:
        exit(0)