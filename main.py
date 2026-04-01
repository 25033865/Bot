from listen import listen
from commands import execute_command
from speak import speak
import time


speak("Max is online")

while True:
    command = listen()

    if command == "":
        time.sleep(0.1)
        continue

    if "exit" in command or "stop" in command:
        speak("Shutting down")
        break

    execute_command(command)