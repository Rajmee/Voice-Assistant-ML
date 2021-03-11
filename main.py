from win32com.client import Dispatch
from cdifflib import CSequenceMatcher
from googletrans import Translator

import pyttsx3
import datetime
import pytz
import fractions
import sys
import speech_recognition as sr

engine = pyttsx3.init()


def speak(audio):
    speaking = Dispatch('SAPI.Spvoice')
    speaking.speak(audio)


def Time():
    time = datetime.datetime.now().strftime("%H:%M:%S")
    speak(time)


def Date():
    date = datetime.datetime.now(tz=pytz.timezone('Asia/Dhaka'))
    speak(date.strftime('%d %B, of %Y'))


def wishme():
    hour = datetime.datetime.now().hour
    if 5 <= hour < 12:
        speak("Good morning sir! Please tell me how can i help you?")
    elif 12 <= hour < 15:
        speak("Good noon sir! Please tell me how can i help you?")
    elif 15 <= hour < 18:
        speak("Good afternoon sir! Please tell me how can i help you?")
    elif 18 <= hour < 24:
        speak("Good evening sir! Please tell me how can i help you?")
    else:
        speak("Good night sir! Hope you had a good day")


def welcomevoice():
    speak("Hello sir....I am Jarvis. Version 1.0. Satellite number 2 0 3 8 7 5. ")
    speak("the current time is")
    Time()
    speak("the current date is")
    Date()
    wishme()
    speak("Jarvis at your service.")


# speak("Hello.....this is Jarvis. Version 1.0. Satellite number 2 0 3 8 7 5. You are ready to go.")
# speak("Now the time is : ")
# Time()

def hwvoice():
    speak("I am good. How are you sir?")


def getcommand():
    minAccuracy = float(0.600000000000000000000000)
    maxAccuracy = 1
    t = "What time is it now"
    dt = "What date is today"
    wl = "Tell me about yourself"
    hw = "How are you Jarvis"
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # print("Jarvis at your service. Command me sir: ")
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        voice = r.recognize_google(audio, language='bn')
        print(voice)
        translator = Translator()
        result = translator.translate(voice, src='bn', dest='en')
        print(result.text)
        command = result.text
        # speak(command)
        timeratio = CSequenceMatcher(None, command, t).ratio()
        dateratio = CSequenceMatcher(None, command, dt).ratio()
        detailsratio = CSequenceMatcher(None, command, wl).ratio()
        hwratio = CSequenceMatcher(None, command, hw).ratio()

        print(timeratio)
        print(dateratio)

        if minAccuracy <= timeratio <= maxAccuracy:
            Time()
            wishme()
            getcommand()
        elif minAccuracy <= dateratio <= maxAccuracy:
            Date()
            wishme()
            getcommand()
        elif minAccuracy <= detailsratio <= maxAccuracy:
            welcomevoice()
            getcommand()
        elif minAccuracy <= hwratio <= maxAccuracy:
            hwvoice()
            getcommand()
        else:
            speak("If you need any help just knock me sir. Jarvis is always at your service. Thank you sir!")


    except Exception as e:
        print(e)
        statement = "I can't understand sir. Please Say that again! "
        print(statement)
        speak(statement)
        return getcommand()
    return voice


getcommand()
