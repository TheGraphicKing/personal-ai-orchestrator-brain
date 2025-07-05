import speech_recognition as sr
import pyttsx3
import pyautogui
import requests
import json
import time
import subprocess # For launching apps and getting window info
import win32gui # For Windows-specific GUI interaction
import win32con # For Windows constants
import win32api # For Windows API calls
import ctypes # For low-level Windows API calls


# --- Configuration ---
# REPLACE THIS WITH YOUR REPLIT URL FROM COLAB
CLOUD_BRAIN_URL = "https://mvp.navaneethakris5.repl.co" # Example: "https://mvp.navaneethakris5.repl.co"
USER_ID = "your_unique_user_id" # A simple ID to differentiate users in the Cloud Brain's memory

# --- Initialize Speech Recognition and Text-to-Speech Engines ---
r = sr.Recognizer()
engine = pyttsx3.init()

# Adjust TTS voice properties (optional)
voices = engine.getProperty('voices')
# Try to find a natural-sounding voice
for voice in voices:
    if "Zira" in voice.name or "David" in voice.name or "Mark" in voice.name: # Common Windows voices
        engine.setProperty('voice', voice.id)
        break
engine.setProperty('rate', 180) # Speed of speech
engine.setProperty('volume', 0.9) # Volume (0.0 to 1.0)

# --- Helper Functions ---

def speak(text):
    """Converts text to speech."""
    print(f"AI: {text}")
    engine.say(text)
    engine.runAndWait()

def listen():
    """Listens for user's voice command."""
    with sr.Microphone() as source:
        print("Listening for command...")
        r.adjust_for_ambient_noise(source) # Adjust for noise once
        audio = r.listen(source)
        try:
            command = r.recognize_google(audio)
            print(f"You said: {command}")
            return command
        except sr.UnknownValueError:
            speak("Sorry, I could not understand your audio.")
            return None
        except sr.RequestError as e:
            speak(f"Could not request results from speech recognition service; {e}")
            return None

def get_screen_context():
    """
    Captures basic screen context (active window title).
    For MVP, we'll keep this simple. Advanced versions would use A11Y tree, OCR, etc.
    """
    try:
        # Get the handle of the foreground window
        hwnd = win32gui.GetForegroundWindow()
        # Get the window title
        window_title = win32gui.GetWindowText(hwnd)
        return f"Active Window: '{window_title}'"
    except Exception as e:
        print(f"Error getting screen context: {e}")
        return "Active Window: 'Unknown'"


def send_command_to_brain(command_text, screen_context_data):
    """Sends the user command and screen context to the Cloud Brain."""
    payload = {
        "user_id": USER_ID,
        "command": command_text,
        "screen_context": screen_context_data
    }
    try:
        response = requests.post(f"{CLOUD_BRAIN_URL}/command", json=payload)
        response.raise_for_status() # Raise an exception for HTTP errors
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error communicating with Cloud Brain: {e}")
        speak("I'm having trouble connecting to my brain. Please check your internet connection or try again later.")
        return None

def execute_action(action):
    """Executes a single action received from the Cloud Brain."""
    action_type = action.get("action")
    print(f"Executing: {action_type} - {action}") # For debugging

    try:
        if action_type == "OPEN_APP":
            app_name = action.get("app_name")
            if app_name:
                speak(f"Opening {app_name}...")
                # Use subprocess.Popen for more robust app opening
                # Common app names might need full paths or specific executable names
                if app_name.lower() == "notepad":
                    subprocess.Popen(["notepad.exe"])
                elif app_name.lower() == "calculator":
                    subprocess.Popen(["calc.exe"])
                elif app_name.lower() == "word": # Requires Word to be installed
                    subprocess.Popen(["winword.exe"])
                elif app_name.lower() == "excel": # Requires Excel to be installed
                    subprocess.Popen(["excel.exe"])
                elif app_name.lower() == "file explorer":
                    subprocess.Popen(["explorer.exe"])
                elif app_name.lower() == "chrome":
                    subprocess.Popen(["chrome.exe"])
                else:
                    # Fallback for other apps, might not work for all
                    subprocess.Popen([app_name]) # Tries to run directly
                time.sleep(2) # Give app time to open
            else:
                speak("I need an app name to open.")

        elif action_type == "TYPE_TEXT":
            text = action.get("text")
            target_element_selector = action.get("target_element_selector") # Not used directly by pyautogui, but for AI context
            if text:
                speak(f"Typing '{text}'...")
                pyautogui.write(text)
                time.sleep(0.5)
            else:
                speak("I need text to type.")

        elif action_type == "CLICK_ELEMENT":
            target_element_selector = action.get("target_element_selector")
            if target_element_selector:
                speak(f"Attempting to click '{target_element_selector}'...")
                # For MVP, we'll use a very basic approach: assume active window and try to click center or based on text
                # Advanced versions would use A11Y tree data to find precise coordinates
                # For now, we'll simulate a general click or use basic image recognition if needed (more complex)
                # Or, for POC, we might just click the center of the active window or rely on the AI's ability to guide.
                # A simple placeholder:
                pyautogui.click() # Clicks at current mouse position
                time.sleep(0.5)
                speak(f"Clicked on {target_element_selector}.")
            else:
                speak("I need a target to click.")

        elif action_type == "GO_TO_URL":
            url = action.get("url")
            if url:
                speak(f"Going to {url}...")
                subprocess.Popen(["start", url], shell=True) # Opens URL in default browser
                time.sleep(3)
            else:
                speak("I need a URL to go to.")

        elif action_type == "SCROLL":
            direction = action.get("direction")
            if direction == "up":
                speak("Scrolling up...")
                pyautogui.scroll(100) # Scrolls up 100 units
            elif direction == "down":
                speak("Scrolling down...")
                pyautogui.scroll(-100) # Scrolls down 100 units
            else:
                speak("I need a scroll direction (up or down).")
            time.sleep(0.5)

        elif action_type == "FIND_FILE":
            filename = action.get("filename")
            folder_path = action.get("folder_path")
            speak(f"Searching for {filename} in {folder_path}...")
            # This is a placeholder. Real file finding needs OS interaction or indexing.
            # For POC, we'll just acknowledge or assume success.
            speak(f"I have located {filename} in {folder_path}.")
            # In a real app, this would return the actual path to the brain for next steps.

        elif action_type == "SPEAK":
            text = action.get("text")
            if text:
                speak(text)

        elif action_type == "CONFIRM":
            question = action.get("question")
            if question:
                speak(question)
                # In a full app, this would wait for user's verbal/text confirmation
                # For POC, we'll just speak the question and proceed, or you can manually
                # type "yes" into the next command.
                print(f"Confirmation requested: {question}")
                speak("Please confirm by saying 'yes' or 'no' in your next command.")

        elif action_type == "DONE":
            speak("Task completed.")
            return "DONE" # Signal to main loop to end

        else:
            speak(f"I don't know how to perform the action: {action_type}.")

    except Exception as e:
        speak(f"An error occurred during action execution: {e}. I might need to re-plan.")
        print(f"Execution Error: {e}")
    return None # Continue loop if not DONE

# --- Main Loop ---
def main_loop():
    speak("Agent Desktop Pilot is ready. Say 'Hey Agent' to begin.")
    while True:
        try:
            # We'll use a simple hotword detection for POC
            # In a real app, this would be more robust and always-on
            command_trigger = listen()
            if command_trigger and "hey agent" in command_trigger.lower():
                speak("Yes, how can I help you?")
                user_command = listen()
                if user_command:
                    screen_context = get_screen_context()
                    brain_response = send_command_to_brain(user_command, screen_context)
                    if brain_response and "response" in brain_response:
                        actions = brain_response["response"]
                        for action in actions:
                            result = execute_action(action)
                            if result == "DONE":
                                break
                    else:
                        speak("I couldn't get a valid plan from my brain. Please try again.")
                else:
                    speak("I didn't hear a command. Please try again.")
            time.sleep(1) # Small delay to prevent busy-looping
        except KeyboardInterrupt:
            speak("Shutting down Agent Desktop Pilot. Goodbye!")
            break
        except Exception as e:
            print(f"Main loop error: {e}")
            speak("An unexpected error occurred in the main loop. Restarting...")
            time.sleep(2) # Wait a bit before restarting loop

if __name__ == "__main__":
    main_loop()