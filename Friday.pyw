import tkinter as tk
from tkinter import PhotoImage
import speech_recognition as sr
import threading
import subprocess
import os
from gtts import gTTS
import pygame
import time
import pystray
from PIL import Image
import tempfile
import pyttsx3
import ctypes
import sys
import datetime
import webbrowser
from win32com.client import Dispatch
from pathlib import Path
import pythoncom
import json

tray_icon = None
if sys.platform == "win32":

    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# Initialize pygame mixer
pygame.mixer.init()
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def shutdown_computer():
    if os.name == 'nt':
        # For Windows operating system
        os.system('shutdown /s /t 0')
    elif os.name == 'posix':
        # For Unix/Linux/Mac operating systems
        os.system('sudo shutdown now')
    else:
        print('Unsupported operating system.')
def add_to_startup():
    try:
        file_path = os.path.realpath(sys.argv[0])

        # Get path to Startup folder
        startup_folder = os.path.join(os.environ['APPDATA'], r"Microsoft\Windows\Start Menu\Programs\Startup")
        shortcut_path = os.path.join(startup_folder, "Friday.lnk")

        # Don't add duplicate shortcut
        if os.path.exists(shortcut_path):
            speak("Friday is already set to run at startup.")
            return

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = file_path
        shortcut.WorkingDirectory = os.path.dirname(file_path)
        shortcut.IconLocation = file_path
        shortcut.save()

        speak("Friday has been added to Windows startup.")
    except Exception as e:
        print(f"Error adding to startup: {e}")
        speak("Failed to add Friday to startup.")
def remove_from_startup():
    try:
        startup_folder = os.path.join(os.environ['APPDATA'], r"Microsoft\Windows\Start Menu\Programs\Startup")
        shortcut_path = os.path.join(startup_folder, "Friday.lnk")

        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            speak("Friday has been removed from startup.")
        else:
            speak("Friday was not set to start with Windows.")
    except Exception as e:
        print(f"Error removing from startup: {e}")
        speak("Failed to remove Friday from startup.")
        
# Calling the shutdown_computer() function to initiate shutdown
def play_sound(filename): 
    try:
        pygame.mixer.Sound(filename).play()
    except Exception as e:
        print(f"Failed to play sound {filename}: {e}") # Grabs Temp Play file and Plays
#try show function
def on_show():
    root.after(0, window.deiconify)
#tray hide function
def on_hide():
    root.after(0, window.withdraw)
#exit app function
def exit_app(icon, item):
   on_exit()
def on_exit(icon=None, item=None):
    if icon:
        icon.stop()
    root.after(0, root.destroy)
    os._exit(0)  # Forcefully exit entire app including threads

#set up tray for windows
def setup_tray_icon():
    icon_path = resource_path("jarvis_icon.png")
    icon_image = Image.open(icon_path)
    menu = pystray.Menu(
        pystray.MenuItem('Show', lambda icon, item: on_show()),
        pystray.MenuItem('Hide', lambda icon, item: on_hide()),
        pystray.MenuItem('Exit', exit_app)
    )
    tray_icon = pystray.Icon("FridayAI", icon_image, "Friday AI Assistant", menu)
    tray_icon.run()

def speak(text):
    temp_file = tempfile.mktemp(suffix=".mp3")
    try:
        tts = gTTS(text=text, lang='en', tld='co.uk')
        tts.save(temp_file)
        pygame.mixer.music.load(temp_file) # when Speak function is called makes a Sound file using GTTS On what the speak function is called to do then plays the temp file
        pygame.mixer.music.play()

        while pygame.mixer.music.get_busy():
            time.sleep(0.1)

        # Stop mixer and unload music so file is released
        pygame.mixer.music.stop()
        pygame.mixer.music.unload() # unloads the temp file 

    finally:
        # Try deleting file, retry a few times if locked
        for i in range(5):
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file) # removes temp file
                break
            except PermissionError:
                time.sleep(0.2)
# Create GUI
root = tk.Tk()
root.withdraw() 
window = tk.Toplevel()
window.overrideredirect(1)
window.attributes('-topmost', True)
window.withdraw()  
window.config(bg='#000000')
window.geometry("250x250")
window.overrideredirect(1)

img = PhotoImage(file=resource_path("logout.png"))
img2 = PhotoImage(file=resource_path("text.gif"))
button = tk.Button(window, image=img, command=lambda: on_exit())
button.config(background='#000000', foreground='#ffffff') # Exit Button
label1 = tk.Label(window, image=img2)
label1.config(background='#000000', foreground='#ffffff')
label1.place(x=20, y=1)
button.place(relx=0.5, rely=1.0, anchor='s')
#Checks for Jarvis ICON For Tray Use
if hasattr(sys, '_MEIPASS'):
    icon_path = os.path.join(sys._MEIPASS, 'jarvis_icon.png')
else:
    icon_path =resource_path('jarvis_icon.png')

icon_image = Image.open(icon_path)

recognizer = sr.Recognizer()
mic = sr.Microphone()

# Load commands from JSON
with open("commands.json", "r") as f:
    COMMANDS = json.load(f)

def execute_command(command_text):
    command_text = command_text.lower()
    for entry in COMMANDS:
        for phrase in entry["phrases"]:
            if phrase in command_text:
                action = entry["action"]
                if action == "open_app":
                    speak(f"Opening {entry['target'].split('/')[-1].split('.')[0].capitalize()}")
                    try:
                        os.startfile(entry["target"])
                    except AttributeError:
                        subprocess.Popen(entry["target"])
                elif action == "add_to_startup":
                    speak("Adding Myself to Windows startup.")
                    add_to_startup()
                elif action == "remove_from_startup":
                    speak("Removing Friday from Windows startup.")
                    remove_from_startup()
                elif action == "speak":
                    speak(entry["text"])
                elif action == "tell_time":
                    current_time = datetime.datetime.now().strftime("%I:%M %p")
                    speak(f"The current time is {current_time}")
                elif action == "tell_date":
                    today = datetime.date.today().strftime("%A, %B %d, %Y")
                    speak(f"Today's date is {today}")
                elif action == "search_google":
                    search_term = command_text.replace(phrase, "").strip()
                    url = f"https://www.google.com/search?q={search_term}"
                    webbrowser.open(url)
                    speak(f"Searching Google for {search_term}")
                elif action == "close_app":
                    speak(f"Closing {entry['target'].split('.')[0].capitalize()}")
                    subprocess.call(["taskkill","/F","/IM",entry["target"]])
                elif action == "shutdown_computer":
                    speak("Shutting down the Computer")
                    shutdown_computer()
                return
    speak("Sorry, I did not understand that command.")

def listen_for_wake_word():
    with mic as source:
        recognizer.adjust_for_ambient_noise(source) # Constantly Monitors Voice From Microphone Till Wake Word said
    print("Listening for wake word...")

    while True:
        with mic as source:
            audio = recognizer.listen(source)
        try:
            phrase = recognizer.recognize_google(audio).lower()
            print(f"Heard: {phrase}")
            if "hey friday" in phrase or "friday" in phrase:
                print("Wake word detected! Listening for command...") # If wake word said Command function is called
                speak("Hello Sir. What can I do for you?")
                listen_for_command()
            elif "wakey wakey daddy's hair" in phrase:
                os.system("powercfg -h off && shutdown /f /r /t 0")
        except sr.UnknownValueError:
            pass
        except sr.RequestError as e:
            print(f"API error: {e}")

def listen_for_command():
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)
        print("Say your command...") # Waits till Command is said
        audio = recognizer.listen(source)
    try:
        command = recognizer.recognize_google(audio).lower()
        print(f"Command recognized: {command}") # if Command called It Executes Command
        
        # Process confirmed command
        if "open notepad" in command: 
            speak("Opening Notepad")
            subprocess.Popen("notepad.exe")
        elif "open calculator" in command:
            speak("Opening Calculator")
            subprocess.Popen("calc.exe")
        elif "enable startup" in command or "add to startup" in command:
            speak("Adding Myself to Windows startup.")
            add_to_startup()
        elif "disable startup" in command or "remove from startup" in command:
            speak("Removing Friday from Windows startup.")
            remove_from_startup()
        elif "open roblox" in command:
            speak("Opening Roblox")
            os.startfile(r"C:\Users\tidge\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Roblox\Roblox Player.lnk")
        elif "open command list" in command or "commands" in command:
            speak("Here are your commands.")
            os.startfile("command.txt") # Commands
        elif "open spotify" in command:
            speak("Opening Spotify")
            os.startfile(r"C:\Users\tidge\AppData\Roaming\Spotify\Spotify.exe")
        elif "youtube" in command:
            speak("Opening YouTube")
            os.startfile(r"C:\Users\tidge\Downloads\YouTube.html")
        elif "open steam" in command:
            speak("Opening Steam")
            os.startfile(r"C:\Program Files (x86)\Steam\steam.exe")
        elif "coding time" in command:
            speak("Opening Visual Studio Code")
            os.startfile(r"C:\Users\tidge\AppData\Local\Programs\Microsoft VS Code\Code.exe")
        elif "open minecraft" in command:
            speak("Launching Minecraft")
            os.startfile(r"C:\XboxGames\Minecraft Launcher\Content\Minecraft.exe")
        elif "thank you" in command:
                speak("It's My Pleasure Sir")
        elif "what's the time" in command or "what is the time" in command:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The current time is {current_time}")
        elif "what's the date" in command or "what is the date" in command:
            today = datetime.date.today().strftime("%A, %B %d, %Y")
            speak(f"Today's date is {today}")
        elif "search for" in command:
            search_term = command.replace("search for", "").strip()
            url = f"https://www.google.com/search?q={search_term}"
            webbrowser.open(url)
            speak(f"Searching Google for {search_term}")
        elif "close spotify" in command:
            speak("Closing Spotify")
            subprocess.call(["taskkill","/F","/IM","spotify.exe"])
        elif "close notepad" in command:
            speak("Closing Notepad")
            subprocess.call(["taskkill","/F","/IM","notepad.exe"])
        elif "close calculator" in command:
            speak("Closing calculator")
            subprocess.call(["taskkill","/F","/IM","calc.exe"])
        elif "close roblox" in command:
            speak("Closing Roblox")
            subprocess.call(["taskkill","/F","/IM","Roblox Player.lnk"])
        elif "close commandlist" in command:
            speak("Closing Commandlist")
            subprocess.call(["taskkill","/F","/IM","notepad.exe"])
        elif "close steam" in command:
            speak("Closing Steam")
            subprocess.call(["taskkill","/F","/IM","steam.exe"])
        elif "close minecraft" in command:
            speak("Closing minecraft")
            subprocess.call(["taskkill","/F","/IM","Minecraft.exe"])
        elif "nothing" in command:
            speak("Okay Sir, Let Me Know if You Need Something")
        elif "shut down" in command:
            speak("Shutting down the Computer")
            shutdown_computer()
        else:
            speak("Sorry, I did not understand that command.")
    except sr.UnknownValueError:
        speak("Sorry, I did not understand that.")
    except sr.RequestError as e:
        print(f"API error: {e}")

# Example placeholder for your listener function
listener_thread = threading.Thread(target=listen_for_wake_word, daemon=True)
listener_thread.start()

tray_thread = threading.Thread(target=setup_tray_icon)
tray_thread.start()
window.mainloop()
#BUILD EXE WITH py -3 -m PyInstaller --onefile --icon=jarvis_icon.ico --add-data "jarvis_icon.png;." --add-data "logout.png;." --add-data "text.gif;." Friday.pyw
