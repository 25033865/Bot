import os
import re
import subprocess
import webbrowser
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus

from speak import speak

try:
    import pyautogui
except Exception:
    pyautogui = None


class AssistantBrain:
    def __init__(self):
        self.pending_confirmation = None
        self.home = Path.home()
        self.folders = {
            "desktop": self.home / "Desktop",
            "documents": self.home / "Documents",
            "downloads": self.home / "Downloads",
            "pictures": self.home / "Pictures",
            "music": self.home / "Music",
            "videos": self.home / "Videos",
        }
        self.apps = {
            "chrome": "start chrome",
            "edge": "start msedge",
            "notepad": "start notepad",
            "calculator": "start calc",
            "paint": "start mspaint",
            "explorer": "start explorer",
            "file explorer": "start explorer",
            "task manager": "start taskmgr",
            "command prompt": "start cmd",
            "powershell": "start powershell",
            "settings": "start ms-settings:",
        }
        self.websites = {
            "youtube": "https://youtube.com",
            "google": "https://google.com",
            "gmail": "https://mail.google.com",
            "github": "https://github.com",
            "linkedin": "https://www.linkedin.com",
            "chatgpt": "https://chat.openai.com",
        }

        if pyautogui is not None:
            pyautogui.PAUSE = 0.05

    def handle(self, raw_command):
        if not raw_command:
            return True

        command = raw_command.strip()
        normalized = " ".join(command.lower().split())

        if self._handle_confirmation(normalized):
            return True

        if self._is_exit_intent(normalized):
            speak("Shutting down. Talk to you soon.")
            return False

        if self._handle_small_talk(normalized):
            return True

        if self._handle_time_date(normalized):
            return True

        if self._handle_help(normalized):
            return True

        if self._handle_system_actions(normalized):
            return True

        if self._handle_web(command, normalized):
            return True

        if self._handle_apps(normalized):
            return True

        if self._handle_folders(command, normalized):
            return True

        if self._handle_file_open(command, normalized):
            return True

        if self._handle_keyboard_actions(command, normalized):
            return True

        speak(
            "I got that, but I don't have an action for it yet. "
            "Say help to hear what I can do."
        )
        return True

    def _handle_confirmation(self, normalized):
        if self.pending_confirmation is None:
            return False

        if any(word in normalized for word in ["confirm", "yes", "go ahead", "do it", "proceed"]):
            description, action = self.pending_confirmation
            self.pending_confirmation = None
            try:
                action()
                speak(f"Done. {description}.")
            except Exception as exc:
                speak(f"I couldn't finish that action: {exc}")
            return True

        if any(word in normalized for word in ["cancel", "no", "never mind"]):
            self.pending_confirmation = None
            speak("Okay, canceled.")
            return True

        speak("Please say confirm to continue or cancel to stop.")
        return True

    def _is_exit_intent(self, normalized):
        return any(
            phrase in normalized
            for phrase in ["exit", "quit", "goodbye", "stop listening", "close assistant"]
        )

    def _handle_small_talk(self, normalized):
        if any(phrase in normalized for phrase in ["hello", "hi ", "hey", "what's up"]):
            speak("Hey. I'm ready. What do you want me to do on your computer?")
            return True

        if "how are you" in normalized:
            speak("I'm running smoothly and ready to help.")
            return True

        if "your name" in normalized or "who are you" in normalized:
            speak("I'm Max, your desktop assistant.")
            return True

        if "thank you" in normalized or "thanks" in normalized:
            speak("Anytime.")
            return True

        return False

    def _handle_time_date(self, normalized):
        now = datetime.now()

        if "time" in normalized and any(word in normalized for word in ["what", "current", "tell"]):
            speak(f"It is {now.strftime('%I:%M %p')}.")
            return True

        if "date" in normalized and any(word in normalized for word in ["what", "today", "tell"]):
            speak(f"Today is {now.strftime('%A, %B %d, %Y')}.")
            return True

        return False

    def _handle_help(self, normalized):
        if "help" not in normalized and "what can you do" not in normalized:
            return False

        print("Max can help with:")
        print("- Open apps: open chrome, open notepad, open calculator")
        print("- Open folders: open downloads, open documents")
        print("- Open file by name: open file budget.xlsx")
        print("- Web: open youtube, search for python tutorials")
        print("- Keyboard control: type hello world, press ctrl+s")
        print("- Utility: take screenshot, what time is it")
        print("- Power: shutdown computer, restart computer (with confirmation)")

        speak("I listed my capabilities in the terminal. Tell me what you want done.")
        return True

    def _handle_system_actions(self, normalized):
        if "shutdown" in normalized:
            self._request_confirmation("Shutting down your computer", lambda: os.system("shutdown /s /t 0"))
            return True

        if "restart" in normalized:
            self._request_confirmation("Restarting your computer", lambda: os.system("shutdown /r /t 0"))
            return True

        if "lock computer" in normalized or "lock screen" in normalized:
            self._request_confirmation(
                "Locking your computer",
                lambda: os.system("rundll32.exe user32.dll,LockWorkStation"),
            )
            return True

        return False

    def _handle_web(self, command, normalized):
        for keyword, url in self.websites.items():
            if f"open {keyword}" in normalized:
                speak(f"Opening {keyword}.")
                webbrowser.open(url)
                return True

        if normalized.startswith("search for "):
            query = command[11:].strip()
            if query:
                speak(f"Searching for {query}.")
                webbrowser.open(f"https://www.google.com/search?q={quote_plus(query)}")
                return True

        website_match = re.search(r"open website (.+)", normalized)
        if website_match:
            domain = website_match.group(1).strip().replace(" ", "")
            if not domain.startswith("http"):
                domain = f"https://{domain}"
            speak("Opening website.")
            webbrowser.open(domain)
            return True

        return False

    def _handle_apps(self, normalized):
        for app_name, app_command in self.apps.items():
            if f"open {app_name}" in normalized:
                speak(f"Opening {app_name}.")
                subprocess.Popen(app_command, shell=True)
                return True

        return False

    def _handle_folders(self, command, normalized):
        for folder_name, folder_path in self.folders.items():
            if f"open {folder_name}" in normalized or f"go to {folder_name}" in normalized:
                self._open_path(folder_path, folder_name)
                return True

        match = re.search(r"(?:open|go to) folder (.+)", command, re.IGNORECASE)
        if match:
            raw_target = match.group(1).strip().strip('"')
            folder_path = Path(os.path.expandvars(os.path.expanduser(raw_target)))
            if not folder_path.is_absolute():
                folder_path = Path.cwd() / folder_path
            self._open_path(folder_path, str(folder_path))
            return True

        return False

    def _handle_file_open(self, command, normalized):
        match = re.search(r"open file (.+)", command, re.IGNORECASE)
        if not match:
            return False

        filename_query = match.group(1).strip().lower()
        found_file = self._find_file(filename_query)

        if found_file is None:
            speak("I couldn't find that file in common folders.")
            return True

        try:
            os.startfile(str(found_file))
            speak(f"Opening {found_file.name}.")
        except Exception as exc:
            speak(f"I found the file but couldn't open it: {exc}")
        return True

    def _handle_keyboard_actions(self, command, normalized):
        if "take screenshot" in normalized:
            if not self._automation_available():
                return True
            image_path = Path.cwd() / f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            pyautogui.screenshot(str(image_path))
            speak(f"Screenshot saved as {image_path.name}.")
            return True

        if normalized.startswith("type "):
            if not self._automation_available():
                return True
            text_to_type = command[5:]
            pyautogui.write(text_to_type, interval=0.02)
            speak("Typed it.")
            return True

        if normalized.startswith("press "):
            if not self._automation_available():
                return True
            key_phrase = normalized[6:].replace(" plus ", "+").replace(" and ", "+")
            keys = [self._normalize_key_name(part.strip()) for part in key_phrase.split("+") if part.strip()]
            if not keys:
                speak("Tell me which key to press.")
                return True
            if len(keys) == 1:
                pyautogui.press(keys[0])
            else:
                pyautogui.hotkey(*keys)
            speak("Done.")
            return True

        if "switch window" in normalized or "next window" in normalized:
            if not self._automation_available():
                return True
            pyautogui.hotkey("alt", "tab")
            speak("Switched window.")
            return True

        return False

    def _request_confirmation(self, description, action):
        self.pending_confirmation = (description, action)
        speak(f"This is a sensitive action. Say confirm to continue: {description}.")

    def _open_path(self, path_value, label):
        path = Path(path_value)
        if not path.exists() or not path.is_dir():
            speak(f"I couldn't find that folder: {label}.")
            return
        os.startfile(str(path))
        speak(f"Opening {label}.")

    def _find_file(self, query):
        search_roots = [Path.cwd(), self.home / "Desktop", self.home / "Documents", self.home / "Downloads"]
        exact_match = None
        partial_match = None

        for root in search_roots:
            if not root.exists():
                continue

            scanned = 0
            for current_root, _, filenames in os.walk(root):
                for filename in filenames:
                    scanned += 1
                    if scanned > 12000:
                        break

                    candidate = Path(current_root) / filename
                    lowered_name = filename.lower()

                    if lowered_name == query:
                        return candidate

                    if query in lowered_name and partial_match is None:
                        partial_match = candidate
                if scanned > 12000:
                    break

            if exact_match is not None:
                return exact_match

        return partial_match

    def _automation_available(self):
        if pyautogui is None:
            speak("Desktop automation is unavailable because pyautogui could not load.")
            return False
        return True

    def _normalize_key_name(self, key_name):
        mapping = {
            "control": "ctrl",
            "return": "enter",
            "escape": "esc",
            "spacebar": "space",
            "del": "delete",
        }
        return mapping.get(key_name, key_name)


assistant = AssistantBrain()


def execute_command(command):
    return assistant.handle(command)