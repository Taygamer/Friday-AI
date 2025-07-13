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

if sys.platform == "win32":
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

# Initialize pygame mixer
pygame.mixer.init()

def shutdown_computer():
    if os.name == 'nt':
        # For Windows operating system
        os.system('shutdown /s /t 0')
    elif os.name == 'posix':
        # For Unix/Linux/Mac operating systems
        os.system('sudo shutdown now')
    else:
        print('Unsupported operating system.')

# Calling the shutdown_computer() function to initiate shutdown
def play_sound(filename): 
    try:
        pygame.mixer.Sound(filename).play()
    except Exception as e:
        print(f"Failed to play sound {filename}: {e}") # Grabs Temp Play file and Plays

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
root.withdraw()  # Hide root completely so it doesn't show in taskbar
window = tk.Toplevel()
window.overrideredirect(1)
window.attributes('-topmost', True)
window.withdraw()  # Optional: start hidden if you want tray control only
window.config(bg='#000000')
window.geometry("250x250")
window.overrideredirect(1)

img = PhotoImage(file=r"logout.png")
img2 = PhotoImage(file=r"text.gif")
button = tk.Button(window, image=img, command=lambda: exit_app())
button.config(background='#000000', foreground='#ffffff') # Exit Button
label1 = tk.Label(window, image=img2)
label1.config(background='#000000', foreground='#ffffff')
label1.place(x=20, y=1)
button.place(relx=0.5, rely=1.0, anchor='s')

recognizer = sr.Recognizer()
mic = sr.Microphone()

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

        if "open notepad" in command:
            speak("Opening Notepad")
            subprocess.Popen("notepad.exe")
        elif "open calculator" in command:
            speak("Opening Calculator")
            subprocess.Popen("calc.exe")
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

# System tray setup
def exit_app(icon=None, item=None):
    print("Exiting application...")
    if icon:
        icon.stop()   # Stop the tray icon properly
    root.quit()       # Close Tkinter window cleanly
    # No exit() here
    

def setup_tray_icon():
    # Load your tray icon image
    icon_image = Image.open("jarvis_icon.png")

    menu = pystray.Menu(
        pystray.MenuItem('Show', lambda icon, item: window.deiconify()),
        pystray.MenuItem('Hide', lambda icon, item: window.withdraw()),  # Makes a Item for the Tray to do certain functions
        pystray.MenuItem('Exit', exit_app)
    )

    tray_icon = pystray.Icon("JarvisAI", icon_image, "Friday AI Assistant", menu)
    tray_icon.run()

# Start listener thread
listener_thread = threading.Thread(target=listen_for_wake_word, daemon=True)
listener_thread.start()

# Run tray icon in its own thread so it doesn't block mainloop
tray_thread = threading.Thread(target=setup_tray_icon, daemon=True)
tray_thread.start()

window.mainloop()
