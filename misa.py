import pyttsx3
import random
import tkinter as tk
from tkinter import *
from tkinter import ttk
import hashlib
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib
import requests, json
import socket
from comtypes.client import CreateObject
root = tk.Tk()
engine = CreateObject("SAPI.SpVoice")
stream = CreateObject("SAPI.SpFileStream")
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def WishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak('good morning')
    
    elif hour>=12 and hour<17:
        speak('good afternoon')
    
    else:
        speak('good evening')

    speak('I am L, how may I help you?')

def taketask():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Listening....')
        r.pause_threshold = 0.8
        r.energy_threshold = 4000
        audio = r.listen(source)

    try:
        print('Recoganizing')
        query = r.recognize_google(audio, language='en-in')
        print("User said: ", query)

    except Exception as e:
        print(e)
        print('say that again please....')
        return "none"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremailid', 'yourpassword')
    server.sendmail('the id to which you want to send the email', to, content)
    server.close()



if __name__ == "__main__":
    WishMe()
    while True:
        query = taketask().lower()

        if 'wikipedia' in query:
            speak('searching wikipedia....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak('according to wikipedia')
            speak(results)
 
        elif 'open youtube' in query:
             webbrowser.open("youtube.com")

        
        elif 'open google' in query:
             webbrowser.open("google.com")

        elif 'open stack overflow' in query:
             webbrowser.open("stackoverflow.com")
            
        elif 'play music' in query:
            music_dir = '' #enter your music_dir
            songs =  os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")

        elif 'open code' in query:
            code_path = "" #enter code path
            os.startfile(code_path)

        elif 'email to' in query:
            x = list(query)
            y = x.split()
            for i in y:
                if i == 'to':
                    z = y.index(i)
                    a = int(y) + 1
                    b = str(y[a])
            try:
                speak('what should i say?')
                content = taketask()
                to =  b #can pass a string here too
                sendEmail(to, content)
                speak('Email has been sent!')
            except Exception as e:
                print(e)
                speak('Sorry, not able to send')

        elif 'exit' in query:
            exit(taketask)

        elif 'how are you' in query:
            speak('I am fine! How are you?')
        
        elif 'i am also fine' in query:
            speak('glad to know that!')

        elif 'open youtube' in query:
              webbrowser.open('youtube.com')

        elif 'search' in query:
            speak('searching wikipedia....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak('according to wikipedia')
            speak(results)

        elif 'tell me about' in query:
            speak('searching wikipedia....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak('according to wikipedia')
            speak(results)

        elif 'who is' in query:
            speak('searching wikipedia....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak('according to wikipedia')
            speak(results)


if __name__ == "__main__":
    taketask()

