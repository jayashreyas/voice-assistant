import pyttsx3
import speech_recognition as sr
import webbrowser
import datetime
import wikipedia
import os
import re
import screen_brightness_control as sbc
import psutil
from plyer import notification
import urllib.request
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


# this method is for taking the commands
# and recognizing the command from the
# speech_Recognition module we will use
# the recognizer method for recognizing


def takeCommand():
    c = sr.Recognizer()

    # from the speech_Recognition module
    # we will use the Microphone module
    # for listening the command
    with sr.Microphone() as source:
        print('Please Speak')

        # seconds of non-speaking audio before
        # a phrase is considered complete
        c.pause_threshold = 0.7
        audio = c.listen(source)

        # Now we will be using the try and catch
        # method so that if sound is recognized
        # it is good else we will have exception
        # handling
        try:
            print("Recognizing")

            # for Listening the command in indian
            # english we can also use 'hi-In'
            # for hindi recognizing
            Query = c.recognize_google(audio, language='en-in')
            print("the command is printed=", Query)

        except Exception as e:
            print(e)
            print("Please Say that again sir")
            return "None"

        return Query


def speak(audio):
    engine = pyttsx3.init()
    # getter method(gets the current value
    # of engine property)
    voices = engine.getProperty('voices')

    # setter method .[0]=male voice and
    # [1]=female voice in set Property.
    engine.setProperty('voice', voices[1].id)

    # Method for the speaking of the assistant
    engine.say(audio)

    # Blocks while processing all the currently
    # queued commands
    engine.runAndWait()


def tellDay_Date():
    # This function is for telling the
    # day of the week
    day = datetime.datetime.today().weekday() + 1

    # this line tells us about the number
    # that will help us in telling the day
    now = datetime.datetime.today()
    print("today= ", now)
    speak(now)
    Day_dict = {1: 'Monday', 2: 'Tuesday',
                3: 'Wednesday', 4: 'Thursday',
                5: 'Friday', 6: 'Saturday',
                7: 'Sunday'}

    if day in Day_dict.keys():
        day_of_the_week = Day_dict[day]
        print(day_of_the_week)
        speak("The day is " + day_of_the_week)

def volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    # Get current volume
    currentVolume = volume.GetMasterVolumeLevel()
    volume.SetMasterVolumeLevel(currentVolume - 6.0, None)
    speak("volume is set to half")
    print("volume is set to half")


def screenBrightness():
    #This method will give the brightness of the PC
    current_brightness = sbc.get_brightness(display=0)
    print(current_brightness)
    speak(current_brightness)

def showBattery():
    #this method will give the battery percentage remaining for the PC
    battery = psutil.sensors_battery()
    percent = battery.percent

    notification.notify(
        title="Battery Percentage",
        message=str(percent) + "% Battery remaining",
        timeout=10
    )

def connect(host='http://google.com'):
    #this method will give the connectivity of the PC
    try:
        urllib.request.urlopen(host) #Python 3.x
        return True
    except:
        return False


def Hello():
    # This function is for when the assistant
    # is called it will say hello and then
    # take query
    speak("hello sir I am your  PC personal assistant Tell me how may I help you ")


def restart_pc():
    #this method will restart the PC
    os.system("shutdown /r /t 1")


def shutdown_pc():
    #this method will shutdown the PC
    os.system("shutdown /s /t 1")


def Take_query():
    # calling the Hello function for
    # making it more interactive
    Hello()

    # This loop is infinite as it will take
    # our queries continuously until and unless
    # we do not say bye to exit or terminate
    # the program
    while (True):

        # taking the query and making it into
        # lower case so that most of the times
        # query matches and we get the perfect
        # output
        query = takeCommand().lower()

        if "open google" in query:
            speak("Opening Google ")
            webbrowser.open("www.google.com")
            continue


        elif 'open youtube' in query:
            speak("Opening Youtube ")
            webbrowser.open("youtube.com")
            continue

        elif 'open instagram' in query:
            speak("Opening insta ")
            webbrowser.open("https://www.instagram.com/")
            continue


        elif 'open amazon' in query:
            speak("Opening amazon ")
            webbrowser.open("https://www.amazon.in/")
            continue

        elif "today" in query:
            tellDay_Date()
            continue

        elif query in ['Battery','battery','BATTERY']:
            showBattery()
            continue

        elif query in ['Brightness','brightness','BRIGHTNESS']:
            screenBrightness()
            continue

        elif query in ['Connectivity','connectivity','CONNECTIVITY']:
            connect()
            speak("connected" if connect() else "no internet!")
            print("connected" if connect() else "no internet!")
            continue

        elif query in ['Volume', 'volume', 'VOLUME']:
            volume()
            continue

        elif  query in ['Time','time','TIME']:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
            print(strTime)
            continue

        elif "restart" in query:
            speak("your PC will restart now")
            restart_pc()
            continue

        elif "shutdown" in query:
            speak("your PC is shutting down")
            shutdown_pc()
            continue

        elif 'code editor' in query:
            codePath = "C:/Users/boora/Desktop/Visual Studio Code.lnk"
            os.startfile(codePath)
            speak("your code editor is getting opened")
            continue

        elif 'play music' in query:
            music = "C:/Users/boora/Downloads/Kesariya - Brahmastra_HD-(Hd9video).mp4"
            os.startfile(music)
            speak("enjoy listening to music")
            continue


        elif query in ['News', 'news', 'NEWS']:
            reg_ex = re.search('open news (.*)', query)
            url = 'https://www.indiatoday.in/'
            if reg_ex:
                subreddit = reg_ex.group(1)
                url = url + 'r/' + subreddit
            webbrowser.open(url)
            speak("opening news for you")
            print('Done!')
            continue

        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
            continue

        elif query in ['Geeksforgeeks', 'geeksforgeeks', 'GEEKSFORGEEKS']:
            speak("Opening geeks for geeks ")
            webbrowser.open("https://www.geeksforgeeks.org/")
            continue

        elif "tell me your name" in query:
            speak("I am Three. Your desktop Assistant")
            continue

# this will exit and terminate the program
        elif "see you" in query:
            speak("Thank you this is Three signing off")
            exit()

if __name__ == '__main__':
    # main method for executing
    # the functions
    Take_query()



