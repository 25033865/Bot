import speech_recognition as sr
import sys


def listen():
    r = sr.Recognizer()

    try:
        with sr.Microphone() as source:
            print("Listening...")
            r.adjust_for_ambient_noise(source, duration=0.5)
            audio = r.listen(source, timeout=5, phrase_time_limit=8)
    except sr.WaitTimeoutError:
        print("Listening timed out. No speech detected.")
        return ""
    except Exception as exc:
        print(f"Microphone unavailable ({exc}). Type your command instead.")
        if sys.stdin and sys.stdin.isatty():
            try:
                return input("You (type): ").strip()
            except EOFError:
                return ""
        return ""

    try:
        command = r.recognize_google(audio)
        print("You:", command)
        return command.strip()
    except sr.UnknownValueError:
        return ""
    except sr.RequestError as exc:
        print(f"Speech service error: {exc}")
        return ""