import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import wikipedia
import pyjokes
import os
import webbrowser
import urllib.parse
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import subprocess
# Creating a recognizer who will understand your voice
listener = sr.Recognizer()
engine = pyttsx3.init()  # Initialize the engine

# If you want to change the voice (0 => male voice) (1 => female voice)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

# Asking the engine to say something
def talk(text):
    engine.say(text)
    engine.runAndWait()

def take_command():
    command = ""
    try:
        # Using microphone
        with sr.Microphone() as source:
            print("Listening...")
            talk("Listening...")
            listener.adjust_for_ambient_noise(source, duration=1)  # Adjust for ambient noise
            voice = listener.listen(source, timeout=5, phrase_time_limit=10)
            command = listener.recognize_google(voice)
            # if 'alexa' in command:
            print(command)
    except Exception as e:
        print(f"Error: {e}")
    return command
def set_brightness(brightness_level):
    try:
        subprocess.run(['powershell.exe', f"(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, {brightness_level})"])
        print(f"Brightness set to {brightness_level}%")
        talk(f"Brightness set to {brightness_level}%")
    except Exception as e:
        print(f"Failed to set brightness: {e}")
        talk(f"Failed to set brightness: {e}")
        
# Function to set volume
def set_volume(volume_level):
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(float(volume_level) / 100.0, None)
        print(f"Volume set to {volume_level}%")
        talk(f"Volume set to {volume_level}%")
    except Exception as e:
        print(f"Failed to set volume: {e}")
        talk("Failed to set volume")

def run_assistant():
    command = take_command()
    command = command.lower()

    if 'play' in command:
        song = command.replace('play', '')
        talk('playing' + song)
        pywhatkit.playonyt(song)
    
    elif 'time' in command:
        time = datetime.datetime.now().strftime('%I:%M %p')
        talk(time)
    
    elif 'tell me about' in command:
        person = command.replace('tell me about', '')
        info = wikipedia.summary(person, 1)
        talk(info)
    
    elif 'are you single' in command:
        talk('I am in a relationship with wifi')
    
    elif 'jokes' in command:
        joke = pyjokes.get_joke()
        print(joke)
        talk(joke)
    
    elif 'open code' in command or 'open VS Code' in command:
        talk('Opening Visual Studio Code')
        os.system("start code")
    
    elif 'open Google ' in command:
        print('Chrome')
        talk('Opening Google Chrome')
        os.system("start chrome")
    
    elif 'search website' in command:
        talk('Sure, what do you want to search for?')
        query = take_command()
        if query:
            talk(f'Searching for {query} in Google Chrome')
            url = f'https://www.google.com/search?q={query}'
            webbrowser.open(url)
        else:
            talk('I did not catch that. Please say it again.')
    
    elif 'open website' in command:
        talk('Sure, please tell me the website URL')
        website = take_command().lower()
        if 'http' not in website:
            website = 'https://www.' + urllib.parse.quote(website)  # Encode spaces correctly
        print(website)
        talk(f'Opening {website}')
        webbrowser.open(website)
    
    elif 'set brightness' in command:
        talk('Sure, what level do you want to set?')
        level = take_command()
        if level:
            print(level)
            set_brightness(level)
        else:
            talk('I did not catch that. Please say it again.')
    
    elif 'set volume' in command:
        talk('Sure, what level do you want to set?')
        volume_level = take_command()
        if volume_level:
            print(volume_level)
            set_volume(volume_level)
        else:
            print('I did not catch that. Please say it again.')
            talk('I did not catch that. Please say it again.')
    
    elif 'open file explorer' in command:
        talk('Opening File Explorer')
        os.system("explorer")
    
    elif 'close the terminal' in command:
        talk('Closing the terminal')
        return False
    
    else:
        talk('Please say it again')
    
    return True

active = True
talk("Hey, your assistant is ready to listen")
while active:
    active = run_assistant()
