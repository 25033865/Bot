import os
import re
import time
import subprocess
import webbrowser
import json
import requests
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus
from urllib.request import urlopen

from speak import speak

try:
    import pyautogui
except Exception:
    pyautogui = None

try:
    import pygetwindow as gw
except Exception:
    gw = None


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
            "chrome": "chrome.exe",
            "edge": "msedge.exe",
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "calc": "calc.exe",
            "paint": "mspaint.exe",
            "explorer": "explorer.exe",
            "file explorer": "explorer.exe",
            "task manager": "taskmgr.exe",
            "command prompt": "cmd.exe",
            "powershell": "powershell.exe",
            "terminal": "wt.exe",
            "windows terminal": "wt.exe",
            "control panel": "control.exe",
            "settings": "ms-settings:",
            "camera": "microsoft.windows.camera:",
            "photos": "ms-photos:",
            "snipping tool": "snippingtool.exe",
            "word": "winword.exe",
            "excel": "excel.exe",
            "powerpoint": "powerpnt.exe",
            "visual studio code": "code.cmd",
            "vscode": "code.cmd",
        }
        self.websites = {
            "youtube": "https://youtube.com",
            "google": "https://google.com",
            "gmail": "https://mail.google.com",
            "github": "https://github.com",
            "linkedin": "https://www.linkedin.com",
            "chatgpt": "https://chat.openai.com",
        }
        self.ollama_url = os.getenv("BOT_OLLAMA_URL", "http://localhost:11434/api/generate")
        self.ollama_model = os.getenv("BOT_OLLAMA_MODEL", "llama3")

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

        if self._handle_window_navigation(command, normalized):
            return True

        if self._handle_keyboard_actions(command, normalized):
            return True

        if self._handle_general_question(command, normalized):
            return True

        if self._handle_ollama_chat(command):
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

        if any(
            phrase in normalized
            for phrase in [
                "i'm good",
                "im good",
                "i am good",
                "i'm fine",
                "im fine",
                "i am fine",
                "i'm doing well",
                "im doing well",
                "i am doing well",
            ]
        ):
            speak("That's great to hear. How can I assist you today?")
            return True

        if any(phrase in normalized for phrase in ["who created you", "who made you"]):
            speak("I was created by a talented developer named Rotondwa.")
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
        print("- Open apps by name: open spotify, open vscode, open calculator")
        print("- Navigate windows: switch to chrome, next window, previous window")
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
        for app_name, app_target in self.apps.items():
            if f"open {app_name}" in normalized:
                if self._launch_app(app_target, window_hint=app_name):
                    speak(f"Opening {app_name}.")
                else:
                    speak(f"I couldn't open {app_name}.")
                return True

        generic_open = re.search(r"^(?:please\s+)?open\s+(.+)$", normalized)
        if generic_open:
            requested_app = generic_open.group(1).strip()
            if requested_app.startswith(("website ", "folder ", "file ")):
                return False
            if requested_app in self.websites:
                return False

            requested_app = re.sub(r"^(the\s+)?(app\s+|application\s+)?", "", requested_app).strip()
            requested_app = re.sub(r"\s+(for me|please)$", "", requested_app).strip()
            if not requested_app:
                return False

            if self._launch_any_app(requested_app):
                speak(f"Opening {requested_app}.")
            else:
                speak(f"I couldn't open {requested_app}.")
            return True

        return False

    def _launch_app(self, app_target, window_hint=None):
        launched = False

        try:
            os.startfile(app_target)
            launched = True
        except Exception:
            pass

        if not launched:
            # Fallback through cmd `start` so PATH-resolved executables still launch.
            try:
                subprocess.run(
                    ["cmd", "/c", f'start "" "{app_target}"'],
                    check=True,
                    capture_output=True,
                    text=True,
                )
                launched = True
            except Exception:
                pass

        if not launched:
            # Last fallback for shell aliases and URI-like app identifiers.
            escaped_target = app_target.replace("'", "''")
            script = f"Start-Process '{escaped_target}'"
            try:
                subprocess.run(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                    check=True,
                    capture_output=True,
                    text=True,
                )
                launched = True
            except Exception:
                pass

        if launched and window_hint:
            time.sleep(0.6)
            self._activate_window(window_hint)

        return launched

    def _launch_any_app(self, app_name):
        normalized_name = app_name.lower().strip()
        mapped_target = self.apps.get(normalized_name)
        if mapped_target and self._launch_app(mapped_target, window_hint=normalized_name):
            return True

        raw_target = app_name.strip().strip('"').strip("'")
        if not raw_target:
            return False

        launch_candidates = [raw_target]
        if not raw_target.lower().endswith((".exe", ".cmd", ".bat", ".lnk")):
            launch_candidates.append(f"{raw_target}.exe")

        compact_candidate = raw_target.replace(" ", "")
        if compact_candidate and compact_candidate.lower() != raw_target.lower():
            launch_candidates.append(f"{compact_candidate}.exe")

        for candidate in launch_candidates:
            if self._launch_app(candidate, window_hint=raw_target):
                return True

        return self._launch_start_menu_app(raw_target)

    def _launch_start_menu_app(self, app_name):
        escaped_name = app_name.replace("'", "''")
        script = (
            "$ErrorActionPreference='Stop';"
            f"$target='{escaped_name}';"
            "$app=Get-StartApps | Where-Object { $_.Name -like ('*' + $target + '*') } | Select-Object -First 1;"
            "if ($null -eq $app) { exit 1 };"
            "Start-Process ('shell:AppsFolder\\' + $app.AppID)"
        )
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                check=True,
                capture_output=True,
                text=True,
            )
            return True
        except Exception:
            return False

    def _handle_window_navigation(self, command, normalized):
        if "next window" in normalized or "switch window" in normalized:
            if not self._automation_available():
                return True
            pyautogui.hotkey("alt", "tab")
            speak("Switched window.")
            return True

        if "previous window" in normalized or "last window" in normalized:
            if not self._automation_available():
                return True
            pyautogui.hotkey("alt", "shift", "tab")
            speak("Moved to previous window.")
            return True

        if "show desktop" in normalized:
            if not self._automation_available():
                return True
            pyautogui.hotkey("win", "d")
            speak("Showing desktop.")
            return True

        if any(phrase in normalized for phrase in ["close current window", "close this window", "close current app"]):
            if not self._automation_available():
                return True
            pyautogui.hotkey("alt", "f4")
            speak("Closed current window.")
            return True

        if any(phrase in normalized for phrase in ["minimize window", "minimize app", "minimize this"]):
            if not self._automation_available():
                return True
            pyautogui.hotkey("win", "down")
            speak("Minimized.")
            return True

        if any(phrase in normalized for phrase in ["maximize window", "maximize app", "maximize this"]):
            if not self._automation_available():
                return True
            pyautogui.hotkey("win", "up")
            speak("Maximized.")
            return True

        target_app = self._extract_navigation_target(command, normalized)
        if not target_app:
            return False

        if self._activate_window(target_app):
            speak(f"Switched to {target_app}.")
            return True

        if self._launch_any_app(target_app):
            speak(f"Opening {target_app}.")
            return True

        speak(f"I couldn't find or open {target_app}.")
        return True

    def _extract_navigation_target(self, command, normalized):
        navigation_match = re.search(
            r"(?:switch|navigate|focus|move|go)\s+(?:to|on)?\s+(.+)",
            normalized,
        )
        if not navigation_match:
            return None

        target = navigation_match.group(1).strip()
        if target in {"next", "previous", "window", "desktop"}:
            return None

        target = re.sub(r"\b(app|application|window)\b$", "", target).strip()
        if not target:
            return None

        return target

    def _activate_window(self, target_app):
        if gw is None:
            return False

        try:
            titles = [title for title in gw.getAllTitles() if title and title.strip()]
        except Exception:
            return False

        if not titles:
            return False

        target_lower = target_app.lower().strip()
        direct_matches = [title for title in titles if target_lower in title.lower()]

        if not direct_matches:
            target_tokens = [token for token in re.findall(r"[a-z0-9]+", target_lower) if len(token) > 1]
            ranked_matches = []
            for title in titles:
                title_lower = title.lower()
                score = sum(1 for token in target_tokens if token in title_lower)
                if score:
                    ranked_matches.append((score, len(title), title))

            ranked_matches.sort(key=lambda item: (-item[0], item[1]))
            direct_matches = [item[2] for item in ranked_matches[:3]]

        for matched_title in direct_matches[:5]:
            try:
                windows = gw.getWindowsWithTitle(matched_title)
                if not windows:
                    continue

                window = windows[0]
                if getattr(window, "isMinimized", False):
                    window.restore()
                    time.sleep(0.2)
                window.activate()
                return True
            except Exception:
                continue

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

        return False

    def _handle_general_question(self, command, normalized):
        question_starters = (
            "who",
            "what",
            "when",
            "where",
            "why",
            "how",
            "which",
            "is",
            "are",
            "can",
            "could",
            "should",
            "would",
            "do",
            "does",
            "did",
        )

        is_question = normalized.endswith("?") or any(
            normalized.startswith(f"{starter} ") for starter in question_starters
        )

        if not is_question:
            return False

        ollama_answer = self._fetch_ollama_answer(command.strip())
        if ollama_answer:
            speak(ollama_answer)
            return True

        answer = self._fetch_web_answer(command.strip())
        if answer:
            speak(answer)
            return True

        speak("I could not find a direct answer, so I opened a web search.")
        webbrowser.open(f"https://www.google.com/search?q={quote_plus(command)}")
        return True

    def _handle_ollama_chat(self, command):
        answer = self._fetch_ollama_answer(command.strip())
        if not answer:
            return False

        speak(answer)
        return True

    def _fetch_ollama_answer(self, query):
        if not query:
            return None

        payload = {
            "model": self.ollama_model,
            "prompt": (
                "You are Max, a desktop voice assistant. "
                "Respond briefly, clearly, and helpfully.\n"
                f"User: {query}"
            ),
            "stream": False,
        }

        try:
            response = requests.post(self.ollama_url, json=payload, timeout=45)
            response.raise_for_status()
            data = response.json()
            answer = (data.get("response") or "").strip()
            if not answer:
                return None
            return self._trim_spoken_answer(answer, max_chars=320)
        except Exception:
            return None

    def _fetch_web_answer(self, query):
        endpoint = (
            "https://api.duckduckgo.com/"
            f"?q={quote_plus(query)}&format=json&no_html=1&skip_disambig=1"
        )

        try:
            with urlopen(endpoint, timeout=6) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except Exception:
            return None

        for key in ["AbstractText", "Definition", "Answer"]:
            value = (payload.get(key) or "").strip()
            if value:
                return self._trim_spoken_answer(value)

        for topic in payload.get("RelatedTopics") or []:
            text = self._extract_related_text(topic)
            if text:
                return self._trim_spoken_answer(text)

        return None

    def _extract_related_text(self, topic):
        if not isinstance(topic, dict):
            return None

        direct_text = (topic.get("Text") or "").strip()
        if direct_text:
            return direct_text

        for nested_topic in topic.get("Topics") or []:
            nested_text = self._extract_related_text(nested_topic)
            if nested_text:
                return nested_text

        return None

    def _trim_spoken_answer(self, text, max_chars=260):
        clean_text = " ".join(text.split())
        if len(clean_text) <= max_chars:
            return clean_text

        shortened = clean_text[:max_chars].rsplit(" ", 1)[0]
        return f"{shortened}."

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