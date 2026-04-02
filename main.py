from listen import listen
from commands import execute_command
from speak import speak
import time


speak("Max is online. Say help to hear what I can do.")

try:
    while True:
        command = listen()

        if command == "":
            time.sleep(0.1)
            continue

        should_continue = execute_command(command)
        if not should_continue:
            break
except KeyboardInterrupt:
    print("\nAssistant stopped by user.")