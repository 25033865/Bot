import pyttsx3


try:
    engine = pyttsx3.init()
except Exception as exc:
    print(f"Text-to-speech unavailable: {exc}")
    engine = None


def speak(text):
    print("Jarvis:", text)
    if engine is None:
        return

    try:
        engine.say(text)
        engine.runAndWait()
    except Exception as exc:
        print(f"Text-to-speech error: {exc}")