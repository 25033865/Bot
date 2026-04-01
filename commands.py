import os
import webbrowser
from speak import speak


def execute_command(command):
    if "open youtube" in command:
        speak("Opening YouTube")
        webbrowser.open("https://youtube.com")

    elif "open google" in command:
        speak("Opening Google")
        webbrowser.open("https://google.com")

    elif "open chrome" in command:
        speak("Opening Chrome")
        os.system("start chrome")

    elif "your name" in command:
        speak("I am Jarvis, your assistant")

    else:
        speak("I don't understand that yet")