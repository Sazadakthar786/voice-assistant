import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

# Initialize text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour < 12:
        speak("Good Morning Sir!")
    elif hour < 18:
        speak("Good Afternoon Sir!")
    else:
        speak("Good Evening Sir!")
    speak("I am your Assistant. Jarvis 1 point o")

def username():
    speak("What should I call you, Sir?")
    uname = takeCommand()
    speak(f"Welcome Mister {uname}")
    columns = shutil.get_terminal_size().columns
    print("#####################".center(columns))
    print(f"Welcome Mr. {uname}".center(columns))
    print("#####################".center(columns))
    speak("How can I help you, Sir?")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except sr.UnknownValueError:
        print("Sorry, I did not understand that.")
        return "None"
    except sr.RequestError:
        print("Sorry, my speech service is down.")
        return "None"
    except Exception as e:
        print(e)
        return "None"
    return query

def sendEmail(to, content):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login('your_email@gmail.com', 'your_email_password')
        server.sendmail('your_email@gmail.com', to, content)
        server.close()
        speak("Email has been sent!")
    except Exception as e:
        print(e)
        speak("I am not able to send this email")

def getWeather(city_name):
    api_key = "YOUR_API_KEY"
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    complete_url = base_url + "appid=" + api_key + "&q=" + city_name
    response = requests.get(complete_url)
    x = response.json()
    
    if x.get("cod") == 200:
        y = x.get("main", {})
        current_temperature = y.get("temp", "N/A")
        current_pressure = y.get("pressure", "N/A")
        current_humidity = y.get("humidity", "N/A")
        z = x.get("weather", [{}])
        weather_description = z[0].get("description", "No description available")
        
        speak(f"Temperature is {current_temperature - 273.15:.2f} Celsius, pressure is {current_pressure} hPa, humidity is {current_humidity} percent, and weather is {weather_description}")
    else:
        speak("City Not Found")

def main():
    clear = lambda: os.system('cls')
    clear()
    wishMe()
    username()
    
    while True:
        query = takeCommand().lower()
        
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to Youtube")
            webbrowser.open("https://youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google")
            webbrowser.open("https://google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Overflow. Happy coding!")
            webbrowser.open("https://stackoverflow.com")

        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            music_dir = "C:\\Users\\GAURAV\\Music"
            songs = os.listdir(music_dir)
            if songs:
                print(songs)
                os.startfile(os.path.join(music_dir, songs[0]))
            else:
                speak("No music files found")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open opera' in query:
            codePath = r"C:\\Users\\GAURAV\\AppData\\Local\\Programs\\Opera\\launcher.exe"
            os.startfile(codePath)

        elif 'email to gaurav' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "receiver@example.com"  # Replace with the actual receiver email
                sendEmail(to, content)
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("Whom should I send it to?")
                to = takeCommand()
                sendEmail(to, content)
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you. How are you, Sir?")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that you're fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "").strip()
            assname = query
            speak(f"Name changed to {assname}")

        elif "change name" in query:
            speak("What would you like to call me, Sir?")
            assname = takeCommand()
            speak(f"Thanks for naming me {assname}")

        elif "what's your name" in query or "What is your name" in query:
            speak(f"My friends call me {assname}")
            print(f"My friends call me {assname}")

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Gaurav.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:
            app_id = "Wolframalpha API ID"  # Replace with your WolframAlpha API key
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text
            speak("The answer is " + answer)

        elif 'search' in query or 'play' in query:
            query = query.replace("search", "").replace("play", "").strip()
            webbrowser.open(query)

        elif "who i am" in query:
            speak("If you talk then definitely you're human.")

        elif "why you came to world" in query:
            speak("Thanks to Gaurav. Further, it's a secret")

        elif 'power point presentation' in query:
            speak("Opening PowerPoint presentation")
            power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
            os.startfile(power)

        elif 'is love' in query:
            speak("It is the 7th sense that destroys all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Gaurav")

        elif 'reason for you' in query:
            speak("I was created as a minor project by Mister Gaurav")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 0, "C:\\path\\to\\wallpaper.jpg", 0)
            speak("Background changed successfully")

        elif 'open bluestack' in query:
            appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            os.startfile(appli)

        elif 'news' in query:
            try:
                api_key = 'YOUR_API_KEY'
                jsonObj = urlopen(f'https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey={api_key}')
                data = json.load(jsonObj)
                speak('Here are some top news from the Times of India')
                i = 1
                for item in data['articles']:
                    print(f"{i}. {item['title']}\n{item['description']}\n")
                    speak(f"{i}. {item['title']}")
                    i += 1
            except Exception as e:
                print(e)
                speak("I am not able to fetch news")

        elif 'lock window' in query:
            speak("Locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold on a sec! Your system is on its way to shut down")
            subprocess.call('shutdown /p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Emptied")

        elif "don't listen" in query or "stop listening" in query:
            speak("For how long do you want to stop Jarvis from listening to commands?")
            try:
                a = int(takeCommand())
                time.sleep(a)
            except ValueError:
                speak("Please provide a valid number")

        elif "where is" in query:
            query = query.replace("where is", "").strip()
            speak(f"User asked to locate {query}")
            webbrowser.open(f"https://www.google.nl/maps/place/{query}")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown /h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the applications are closed before signing out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should I write, Sir?")
            note = takeCommand()
            with open('jarvis.txt', 'w') as file:
                speak("Should I include the date and time?")
                snfm = takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    file.write(f"{strTime} :- {note}")
                else:
                    file.write(note)
            speak("Note written successfully")

        elif "show note" in query:
            speak("Showing Notes")
            try:
                with open("jarvis.txt", "r") as file:
                    content = file.read()
                print(content)
                speak(content)
            except FileNotFoundError:
                speak("No notes found")

        elif "update assistant" in query:
            speak("After downloading the file, please replace this file with the downloaded one")
            url = 'YOUR_UPDATE_URL'
            r = requests.get(url, stream=True)
            with open("Voice.py", "wb") as Pypdf:
                total_length = int(r.headers.get('content-length', 0))
                for ch in progress.bar(r.iter_content(chunk_size=2391975), expected_size=(total_length / 1024) + 1):
                    if ch:
                        Pypdf.write(ch)
            speak("Update completed")

        elif "jarvis" in query:
            wishMe()
            speak("Jarvis 1 point o in your service, Mister")
            speak(assname)

        elif "weather" in query:
            speak("City name")
            print("City name: ")
            city_name = takeCommand()
            getWeather(city_name)

        elif "send message" in query:
            account_sid = 'YOUR_ACCOUNT_SID'
            auth_token = 'YOUR_AUTH_TOKEN'
            client = Client(account_sid, auth_token)
            speak("What should I send in the message?")
            message_body = takeCommand()
            speak("Who should I send it to?")
            to_number = takeCommand()
            message = client.messages.create(
                body=message_body,
                from_='YOUR_TWILIO_NUMBER',
                to=to_number
            )
            print(message.sid)
            speak("Message sent")

        elif "wikipedia" in query:
            webbrowser.open("https://wikipedia.com")

        elif "Good Morning" in query:
            speak("A warm Good Morning! How are you Mister?")
            speak(assname)

        elif "will you be my gf" in query or "will you be my bf" in query:
            speak("I'm not sure about that. Maybe you should give me some time")

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what is" in query or "who is" in query:
            app_id = "Wolframalpha API ID"  # Replace with your WolframAlpha API key
            client = wolframalpha.Client(app_id)
            res = client.query(query)
            try:
                answer = next(res.results).text
                speak(answer)
            except StopIteration:
                speak("No results")

if __name__ == '__main__':
    main()