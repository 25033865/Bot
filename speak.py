import pyttsx3


try:
    engine = pyttsx3.init()
    engine.setProperty("rate", 175)
    voices = engine.getProperty("voices") or []
    preferred_voice = None

    for voice in voices:
        voice_name = (voice.name or "").lower()
        if "zira" in voice_name or "female" in voice_name:
            preferred_voice = voice.id
            break

    if preferred_voice is not None:
        engine.setProperty("voice", preferred_voice)
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