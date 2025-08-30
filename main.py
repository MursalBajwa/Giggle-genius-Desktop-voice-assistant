import ctypes
import win32com.client
import psutil
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import os
import webbrowser
import pyjokes
import pywhatkit as kit
import tkinter as tk
from tkinter import ttk
from tkinter import LEFT, BOTH, SUNKEN
from PIL import Image, ImageTk,ImageGrab
from threading import Thread
import pytesseract
import pyautogui
import subprocess
import shutil
import pyaudio
from PIL import ImageOps
import time



shell = win32com.client.Dispatch("Shell.Application")
TAB_BAR_HEIGHT = 88



# This tracks the last folder path opened in Explorer so we don’t spawn multiple windows for the same path.
explorer_folder_path = None

# Constants for custom styling
BG_COLOR = "#D2C6E2"
BUTTON_COLOR = "#F9F4F2"
BUTTON_FONT = ("Arial", 14, "bold")
BUTTON_FOREGROUND = "black"
HEADING_FONT = ("white", 24, "bold")
INSTRUCTION_FONT = ("Helvetica", 14)

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update path if needed

def open_or_navigate(path):
    """
    Look for any already-open File Explorer window.
      • If found, call window.Navigate(...) to the new path.
      • Otherwise, Shell.Explore(...) opens a new window.
    """
    # Convert e.g. "M:\data" → "file:///M:/data"
    url = "file:///" + path.replace("\\", "/")
    for window in shell.Windows():
        if window.Name.lower().startswith("file explorer"):
            try:
                window.Navigate(url)
                return
            except Exception:
                pass

    # If no File Explorer window was found, open a fresh one:
    shell.Explore(path)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


entry = None
stop_flag = False  # Define the stop_flag variable at the top of the script


def wish_time():
    global entry
    x = entry.get()
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 6:
        speak('Good night! Sleep tight.')
    elif 6 <= hour < 12:
        speak('Good morning!')
    elif 12 <= hour < 18:
        speak('Good afternoon!')
    else:
        speak('Good evening!')
    speak("My name is GiggleGenius")
    speak(f"{x} How can I help you?")


def take_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Say something:")
        speak("say something")
        recognizer.pause_threshold = 0.8
        recognizer.adjust_for_ambient_noise(source, duration=0.5)  # Adjust for 1 second of ambient noise
        audio = recognizer.listen(source)

    try:
        print("Recognizing...")
        speak("recognizing")
        query = recognizer.recognize_google(audio, language='en-in')
        print(f"You said: {query}")
    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query


def report_battery_status():
    battery = psutil.sensors_battery()
    percent = battery.percent
    speak(f"The battery is at {percent} percent.")
    if battery.power_plugged:
        speak("Charger is plugged in.")
    else:
        speak("Charger is not plugged in.")
    if percent < 20 and not battery.power_plugged:
        speak("Warning: Battery is low. Please plug in the charger.")

def report_system_performance():
    cpu_usage = psutil.cpu_percent()
    memory = psutil.virtual_memory()
    speak(f"Current CPU usage is {cpu_usage} percent.")
    speak(f"RAM usage is {memory.percent} percent.")

def lock_computer():
        speak("Lock the computer.")
        ctypes.windll.user32.LockWorkStation()

def enable_privacy_mode():
        speak("Enabling privacy mode. Minimizing all windows and muting the volume.")
        pyautogui.hotkey('win', 'd')
        pyautogui.press('volumemute')

def disable_privacy_mode():
        speak("Turning off privacy mode. Unmuting the volume.")
        pyautogui.press('volumemute')
def perform_task():
    global stop_flag
    global explorer_folder_path
    while not stop_flag:
        query = take_command().lower()  # Converting user query into lower case
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            try:
                results = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            except wikipedia.exceptions.DisambiguationError as e:
                # Handle disambiguation error (when the search term has multiple possible meanings)
                print(f"There are multiple meanings for '{query}'. Please be more specific.")
                speak(f"There are multiple meanings for '{query}'. Please be more specific.")
            except wikipedia.exceptions.PageError as e:
                # Handle page not found error (when the search term does not match any Wikipedia page)
                print(f"'{query}' does not match any Wikipedia page. Please try again.")
                speak(f"'{query}' does not match any Wikipedia page. Please try again.")
        elif 'play' in query:
            song = query.replace('play', "")
            speak("Playing " + song)
            kit.playonyt(song)

        elif 'open youtube' in query:
            webbrowser.open("https://www.youtube.com/")
        elif 'open google' in query:
            webbrowser.open("https://www.google.com/")
        elif 'open project' in query:
            webbrowser.open("http://localhost/Project/Home.html")
        elif 'search' in query:
            s = query.replace('search', '')
            kit.search(s)
        elif 'the time' in query:
            str_time = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {str_time}")
        elif 'open code' in query:
            code_path = r"D:\Visual studio code\Microsoft VS Code\Code.exe"


        elif 'joke' in query:
            speak(pyjokes.get_joke())


        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location.replace(" ", "+"))








        elif 'click' in query:

            word_to_find = query.replace('click', '').strip()  # Extract the word to find, e.g., "login"

            speak(f"Looking for the word '{word_to_find}' on the screen")

            # Define a scale factor to enlarge the image for better OCR accuracy

            scale_factor = 2.0  # Resizing to 2x the original size

            # Capture the screen

            screen = ImageGrab.grab()

            # Resize the screen capture to improve text detection

            large_screen = screen.resize((int(screen.width * scale_factor), int(screen.height * scale_factor)),
                                         Image.LANCZOS)

            # Convert to grayscale

            gray = large_screen.convert('L')

            gray_enhanced = ImageOps.autocontrast(gray)  # Enhance contrast

            # Try OCR on the enhanced grayscale image first

            ocr_data = pytesseract.image_to_data(gray_enhanced, output_type=pytesseract.Output.DICT, config='--psm 11')

            found = False

            for i, text in enumerate(ocr_data['text']):

                y = ocr_data['top'][i]

                # Check if the text is below the tab bar (adjusted for scale) and contains the target word

                if y >= TAB_BAR_HEIGHT * scale_factor and word_to_find.lower() in text.lower():
                    found = True

                    # Get coordinates in the resized image

                    x, w, h = ocr_data['left'][i], ocr_data['width'][i], ocr_data['height'][i]

                    # Calculate the click position in original screen coordinates

                    click_x = (x + w // 2) / scale_factor

                    click_y = (y + h // 2) / scale_factor

                    # Move to the position and click

                    pyautogui.moveTo(click_x, click_y)

                    pyautogui.click()

                    speak(f"Clicked on the word '{word_to_find}'.")

                    break

            # If not found, try inverting the image to handle light text on dark backgrounds

            if not found:

                inverted_gray = ImageOps.invert(gray)

                inverted_enhanced = ImageOps.autocontrast(inverted_gray)

                ocr_data_inv = pytesseract.image_to_data(inverted_enhanced, output_type=pytesseract.Output.DICT,
                                                         config='--psm 11')

                for i, text in enumerate(ocr_data_inv['text']):

                    y = ocr_data_inv['top'][i]

                    if y >= TAB_BAR_HEIGHT * scale_factor and word_to_find.lower() in text.lower():
                        found = True

                        x, w, h = ocr_data_inv['left'][i], ocr_data_inv['width'][i], ocr_data_inv['height'][i]

                        click_x = (x + w // 2) / scale_factor

                        click_y = (y + h // 2) / scale_factor

                        pyautogui.moveTo(click_x, click_y)

                        pyautogui.click()

                        speak(f"Clicked on the word '{word_to_find}'.")

                        break

            # If still not found, inform the user

            if not found:
                speak(
                    f"Could not find the word '{word_to_find}' on the screen")

        elif 'previous tab' in query:
            # Tell the user we’re switching tabs
            speak("Switching to the previous tab")
            # Simulate pressing Ctrl+Shift+Tab
            pyautogui.hotkey('ctrl', 'shift', 'tab')

        elif 'navigate back' in query:
            # Tell the user we’re going back in the browser
            speak("Navigating back")
            # Simulate pressing Alt+Left Arrow, which makes the active browser go back one page
            pyautogui.hotkey('alt', 'left')

        elif 'navigate forward' in query:
            speak("Navigating forward")
            pyautogui.hotkey('alt', 'right')


        elif 'scroll down' in query:
            speak("Scrolling down")
            pyautogui.scroll(-400)  # Negative value scrolls down

        elif 'scroll up' in query:
            speak("Scrolling up")
            pyautogui.scroll(400)  # Positive value scrolls up

        elif 'clear' in query:
            # Functionality to clear all typed content
            speak("Clearing the content from the active input field.")

            # Simulate pressing Ctrl+A (Select All)
            pyautogui.hotkey('ctrl', 'a')

            # Simulate pressing Backspace (Clear content)
            pyautogui.press('backspace')

        elif 'type' in query:
            speak("What should I type?")
            # Listen for the text to type
            text_to_type = take_command()
            if text_to_type:
                speak(f"typing: {text_to_type}")
                text_to_type = text_to_type.replace(" ", "")
                pyautogui.typewrite(text_to_type)  # Type the text
            else:
                speak("I didn't catch that. Please say it again.")

        elif 'write that' in query:  # New condition for logging admin commands
            speak("Please speak the content you want to log.")
            admin_command = take_command()
            if admin_command != "None":
                with open("admin_log.txt", "a") as file:
                    file.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {admin_command}\n")
                speak("Your command has been written successfully.")
            else:
                speak("No valid input detected. Please try again.")

        elif 'backspace' in query:
            speak("Typing backspace.")
            pyautogui.press('backspace')  # Simulates pressing the backspace key

            # Continuing in perform_task function
        elif 'minimize window' in query:
            pyautogui.hotkey('alt', 'space')
            pyautogui.press('n')
            speak("Minimizing the current window")

        elif 'maximize window' in query:
            pyautogui.hotkey('alt', 'space')
            pyautogui.press('x')
            speak("Maximizing the current window")

        elif 'switch window' in query:
            pyautogui.hotkey('alt', 'tab')
            speak("Switching to the next window")

        elif 'close window' in query:
            pyautogui.hotkey('alt', 'f4')
            speak("Closing the current window")

        elif 'what is my current directory' in query or 'where am i' in query:
            try:
                cwd = os.getcwd()
                # Speak the absolute path
                speak(f"Your current working directory is: {cwd}")
            except Exception as e:
                speak(f"Sorry, I could not retrieve the current directory. {e}")

            
        elif 'open file explorer' in query:

            try:
                cwd = os.getcwd()
                if explorer_folder_path != cwd:
                    os.startfile(cwd)
                    explorer_folder_path = cwd
                speak(f"Opening File Explorer at {cwd}")
            except Exception as e:
                speak(f"Could not open File Explorer. {e}")




        elif 'change directory to' in query:

            try:

                remainder = query.replace('change directory to', '').strip().upper()

                if remainder.endswith(':'):

                    drive_letter = remainder[:-1]

                else:

                    drive_letter = remainder

                drive_path = f"{drive_letter}:\\"

                if os.path.isdir(drive_path):

                    os.chdir(drive_path)

                    open_or_navigate(drive_path)

                    speak(f"Changed directory to {drive_path}")

                else:

                    speak(f"Drive {drive_path} not found on this PC.")

            except Exception as e:

                speak(f"Error changing directory. {e}")


        elif 'go to' in query:

            try:

                subfolder_name = query.replace('go to', '').strip()

                current = os.getcwd()

                candidate = os.path.join(current, subfolder_name)

                if os.path.isdir(candidate):

                    os.chdir(candidate)

                    open_or_navigate(candidate)

                    speak(f"Navigated to {candidate}")

                else:

                    speak(f"Folder {subfolder_name} does not exist under {current}")

            except Exception as e:

                speak(f"Error going to {subfolder_name}. {e}")


        elif 'go back' in query or 'go up' in query:

            try:

                current = os.getcwd()

                parent = os.path.dirname(current)

                if parent and os.path.isdir(parent):

                    os.chdir(parent)

                    open_or_navigate(parent)

                    speak(f"Moved up to {parent}")

                else:

                    speak("No parent directory found.")

            except Exception as e:

                speak(f"Error going back. {e}")

        elif 'list files' in query:
            try:
                current = os.getcwd()
                files = os.listdir('.')
                if files:
                    speak(f"Files and folders in {current}:")
                    for f in files:
                        speak(f)
                else:
                    speak(f"{current} is empty.")
            except Exception as e:
                speak(f"Error listing files. {e}")



        elif 'rename file' in query:

            try:

                # Strip off the trigger phrase, leaving something like "oldname to newname"

                remainder = query.replace('rename file', '').strip()

                # Expect the format: "<old_name> to <new_name>"

                if ' to ' in remainder:

                    old_name_part, new_name_part = remainder.split(' to ', 1)

                    old_name = old_name_part.strip()

                    new_name = new_name_part.strip()

                    cwd = os.getcwd()

                    source_path = os.path.join(cwd, old_name)

                    dest_path = os.path.join(cwd, new_name)

                    # Check if source exists (file or folder)

                    if os.path.exists(source_path):

                        # Perform the rename (works for both files and directories)

                        os.rename(source_path, dest_path)

                        speak(f"Renamed '{old_name}' to '{new_name}'.")

                    else:

                        speak(f"No file or folder named '{old_name}' in the current directory ({cwd}).")

                else:

                    speak("To rename, say: rename file oldname to newname.")

            except Exception as e:

                speak(f"Could not rename '{old_name}'. Error: {e}")


        elif 'find file' in query:

            try:

                # 1) Extract the raw search term. If the user said exactly "find file" with no term,

                #    we prompt them and skip the rest of this block.

                search_term = query.replace('find file', '').strip()

                if not search_term:
                    speak("Please say something after 'find file'.")

                    # Since we're inside the while‐loop, use continue to go back to listening again.

                    continue

                # 2) Find which window is currently in front.

                fg_hwnd = ctypes.windll.user32.GetForegroundWindow()

                explorer_hwnd = None

                explorer_path = None

                for window in shell.Windows():

                    try:

                        # If this is a File Explorer window AND it matches the front HWND:

                        if window.HWND == fg_hwnd and window.Name.lower().startswith("file explorer"):
                            explorer_hwnd = window.HWND

                            explorer_path = window.Document.Folder.Self.Path

                            break

                    except Exception:

                        # In some rare cases, window.Document might not be accessible.

                        continue

                # 3) If no Explorer was in front, open a new one at the script’s CWD:

                if explorer_hwnd is None:

                    explorer_path = os.getcwd()

                    subprocess.Popen(f'explorer "{explorer_path}"')

                    speak(f"No File Explorer was in front. Opening Explorer at {explorer_path}.")

                    # Wait a short moment for the new window to appear.

                    time.sleep(1.0)

                    # Now try again to find the HWND for the newly opened Explorer.

                    new_hwnd = None

                    for window in shell.Windows():

                        try:

                            if window.Name.lower().startswith("file explorer"):

                                # If this Explorer’s path matches explorer_path, capture its HWND.

                                if (hasattr(window.Document.Folder.Self, 'Path') and

                                        window.Document.Folder.Self.Path.lower() == explorer_path.lower()):
                                    new_hwnd = window.HWND

                                    break

                        except Exception:

                            continue

                    # If we still didn’t find the new window, just use whatever is frontmost now:

                    if new_hwnd:

                        explorer_hwnd = new_hwnd

                    else:

                        explorer_hwnd = ctypes.windll.user32.GetForegroundWindow()


                else:

                    # We already had an Explorer in front. Just bring it forward again in case it was

                    # partially hidden.

                    ctypes.windll.user32.SetForegroundWindow(explorer_hwnd)

                    speak(f"Searching inside File Explorer folder: {explorer_path}")

                # 4) By now, `explorer_hwnd` should be the handle to the Explorer window we want.

                #    Switch focus to it (in case something else grabbed focus during our wait).

                ctypes.windll.user32.SetForegroundWindow(explorer_hwnd)

                time.sleep(0.2)  # Give it a moment to actually become active

                # 5) Send the keystrokes: Ctrl+E (focus Search box), type the term, then Enter.

                #    This mimics exactly what you do manually when you click in the top‐right search field.

                pyautogui.hotkey('ctrl', 'e')

                time.sleep(0.2)

                pyautogui.typewrite(search_term)  # Type e.g. "data"

                time.sleep(0.2)

                pyautogui.press('enter')

                speak(f"Typing '{search_term}' into Explorer’s search and launching the search.")


            except Exception as e:

                # If anything above fails, we speak the error and keep running the assistant.

                speak(f"Sorry, something went wrong while searching: {e}")

        elif 'battery status' in query:
            report_battery_status()

        elif 'system performance' in query:
            report_system_performance()
        elif 'lock the computer' in query:
            lock_computer()

        elif 'privacy mode' in query:
            enable_privacy_mode()

        elif 'turn off privacy mode' in query:
            disable_privacy_mode()


        elif 'shutdown system' in query:
            speak("Shutting down now.")
            os.system("shutdown /s /t 1")


        elif 'restart system' in query:
            speak("Restarting now.")
            os.system("shutdown /r /t 1")
   

        elif 'exit' in query:
            speak("thanks for giving your time")
            stop_voice_assistant()


def stop_voice_assistant():
    global stop_flag
    speak("Stopping the Voice Assistant.")
    stop_flag = True


def start_voice_assistant():
    global stop_flag
    wish_time()
    perform_task()
    stop_flag = False  # Reset the flag to False when starting the voice assistant


def main():
    # Create the main GUI window
    root = tk.Tk()
    root.title("Voice Assistant")
    root.geometry("500x700")
    root.configure(bg=BG_COLOR)

    def on_button_click():
        global stop_flag
        if not stop_flag:
            stop_flag = False  # Reset the flag to False when starting the voice assistant
            Thread(target=start_voice_assistant).start()
        else:
            stop_voice_assistant()

    # Load and set the background image
    background_image = Image.open(
        "wallpaperflare.com_wallpaper.jpg")  # Replace "path/to/your/background_image.jpg" with the actual image file path
    background_photo = ImageTk.PhotoImage(background_image)
    background_label = ttk.Label(root, image=background_photo)
    background_label.place(x=-1, y=0, relwidth=1, relheight=1)

    f1 = ttk.Frame(root)
    f1.pack(pady=100)  # Add some padding to the frame to center it vertically



    # Heading
    heading_label = ttk.Label(root, text="Voice Assistant", font=HEADING_FONT, background=BG_COLOR)
    heading_label.pack(pady=20)

    global entry
    f1 = ttk.Frame(root)
    f1.pack()
    l1 = ttk.Label(f1, text="Enter Your Name", font=INSTRUCTION_FONT, background=BG_COLOR)
    l1.pack(side=LEFT, fill=BOTH)
    entry = ttk.Entry(f1, width=30)
    entry.pack(pady=10)

    # Instruction
    instruction_label = ttk.Label(root, text="Click the button below to start the Voice Assistant.",
                                  font=INSTRUCTION_FONT, background=BG_COLOR)
    instruction_label.pack(pady=10)

    # Create and place a button on the GUI
    button = ttk.Button(root, text="Start Voice Assistant", command=on_button_click,
                        style="VoiceAssistant.TButton")
    button.pack(pady=20)

    # Style the button
    style = ttk.Style(root)
    style.configure("VoiceAssistant.TButton", font=BUTTON_FONT, background=BUTTON_COLOR, foreground=BUTTON_FOREGROUND)

    # Run the GUI main loop
    root.mainloop()


if __name__ == "__main__":
    main()