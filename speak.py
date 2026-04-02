import platform
import subprocess

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


def _speak_with_windows_sapi(text):
    escaped_text = text.replace("'", "''")
    script = (
        "$voice = New-Object -ComObject SAPI.SpVoice;"
        "$voice.Rate = 1;"
        f"[void]$voice.Speak('{escaped_text}')"
    )
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        error_details = (result.stderr or result.stdout or "").strip()
        if not error_details:
            error_details = f"PowerShell exited with code {result.returncode}"
        raise RuntimeError(error_details)


def speak(text):
    print("Jarvis:", text)

    if platform.system().lower() == "windows":
        try:
            _speak_with_windows_sapi(text)
            return
        except Exception as exc:
            print(f"Windows speech error: {exc}")

    if engine is not None:
        try:
            engine.say(text)
            engine.runAndWait()
            return
        except Exception as exc:
            print(f"Text-to-speech error: {exc}")