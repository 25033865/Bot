import speech_recognition as sr
import os
import re
import sys


recognizer = sr.Recognizer()
recognizer.dynamic_energy_threshold = True
recognizer.pause_threshold = 0.8
recognizer.non_speaking_duration = 0.4

_calibrated = False
_source_announced = False

_mic_index = None
_resolved_mic_index = None
mic_env_value = (os.getenv("BOT_MIC_INDEX") or "").strip()
if mic_env_value:
    try:
        _mic_index = int(mic_env_value)
    except ValueError:
        print(f"Invalid BOT_MIC_INDEX value: {mic_env_value}. Using default microphone.")
        _mic_index = None


def _resolve_microphone_index():
    if _mic_index is not None:
        return _mic_index

    try:
        mic_names = sr.Microphone.list_microphone_names()
    except Exception:
        return None

    for index, name in enumerate(mic_names):
        lowered = (name or "").lower()
        tokens = set(re.findall(r"[a-z0-9]+", lowered))
        if "stereo mix" in lowered or "speaker" in lowered or "output" in lowered:
            continue
        if "microphone" in tokens or "mic" in tokens or "array" in tokens:
            return index

    for index, name in enumerate(mic_names):
        lowered = (name or "").lower()
        if "stereo mix" in lowered or "speaker" in lowered or "output" in lowered:
            continue
        if "input" in lowered:
            return index

    return None


def _typed_fallback(prompt="You (type): "):
    if sys.stdin and sys.stdin.isatty():
        try:
            return input(prompt).strip()
        except EOFError:
            return ""
    return ""


def _capture_audio():
    global _calibrated, _source_announced, _resolved_mic_index

    if _resolved_mic_index is None:
        _resolved_mic_index = _resolve_microphone_index()

    if not _source_announced:
        try:
            mic_names = sr.Microphone.list_microphone_names()
            if _resolved_mic_index is None:
                print("Using default microphone.")
            elif 0 <= _resolved_mic_index < len(mic_names):
                print(f"Using microphone [{_resolved_mic_index}]: {mic_names[_resolved_mic_index]}")
            else:
                print(f"BOT_MIC_INDEX {_resolved_mic_index} is out of range. Using default microphone.")
        except Exception:
            print("Using default microphone.")
        _source_announced = True

    with sr.Microphone(device_index=_resolved_mic_index) as source:
        if not _calibrated:
            print("Calibrating microphone for background noise...")
            recognizer.adjust_for_ambient_noise(source, duration=1.0)
            _calibrated = True

        print("Listening...")
        return recognizer.listen(source, timeout=8, phrase_time_limit=12)


def listen():
    try:
        audio = _capture_audio()
    except sr.WaitTimeoutError:
        print("I didn't hear anything. Please try again.")
        return ""
    except Exception as exc:
        print(f"Microphone unavailable ({exc}). Type your command instead.")
        return _typed_fallback()

    try:
        command = recognizer.recognize_google(audio, language="en-US").strip()
        if command:
            print("You:", command)
        return command
    except sr.UnknownValueError:
        try:
            command = recognizer.recognize_google(audio, language="en-GB").strip()
            if command:
                print("You:", command)
            return command
        except sr.UnknownValueError:
            print("I heard audio, but couldn't understand the words. Please speak clearly and try again.")
            return ""
        except sr.RequestError as exc:
            print(f"Speech service error: {exc}. Type your command instead.")
            return _typed_fallback()
    except sr.RequestError as exc:
        print(f"Speech service error: {exc}. Type your command instead.")
        return _typed_fallback()