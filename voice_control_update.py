#python voice_control_update.py
import speech_recognition as sr
import os
import pyautogui
import glob
import webbrowser
import pygetwindow as gw
import psutil
import subprocess
import time
import ctypes
import datetime
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Function to decrease the system volume
def decrease_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = max(0, current_volume - 0.1)  # Adjust as needed
    volume.SetMasterVolumeLevelScalar(new_volume, None)

# Function to increase the system volume
def increase_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = min(1, current_volume + 0.1)  # Adjust as needed
    volume.SetMasterVolumeLevelScalar(new_volume, None)

def sleep_system():
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

def lock_screen():
    ctypes.windll.user32.LockWorkStation()

def restart_system():
    os.system("shutdown /r /t 1")

# Function to execute commands based on recognized speech
def execute_command(command):
    if "open app" in command:
        open_app()
    elif "unmute volume" in command:
        unmute_volume()
    elif command.startswith("open folder"):
        open_folder(command)
    elif command.startswith("close folder"):
        close_folder(command)
    elif command.startswith("open file"):
        open_file(command)
    elif command.startswith("close file"):
        close_file(command)
    elif "open website " in command:
        open_website(command)
    elif command.startswith("open "):
        open_command(command)
    elif "sleep system" in command:
        sleep_system()
    elif "turn on hotspot" in command:
        turn_on_hotspot()
    elif "turn off hotspot" in command:
        turn_off_hotspot()
    elif "turn on wifi" in command:
        turn_on_wifi()
    elif "turn off wifi" in command:
        turn_off_wifi()
    elif "close app" in command:
        close_app()
    elif "restart system" in command:
        restart_system()
    elif "lock screen" in command:
        lock_screen()
    elif "maximize window" in command or "maximise window" in command: 
        maximize_window()
    elif "minimize window" in command or "minimise window" in command:  
        minimize_window()
    elif "switch window to" in command:
        switch_window(command)
    elif "take screenshot" in command:
        take_screenshot()
    elif "increase volume" in command:
        increase_volume()
    elif "decrease volume" in command:
        decrease_volume()
    elif "next slide" in command:
        next_slide()
    elif "previous slide" in command:
        previous_slide()
    elif "close window" in command:
        close_window()
    elif "mute volume" in command:
        mute_volume()
    elif "set timer" in command:
        set_timer(command)
    elif "search web" in command:
        search_web(command)
    elif "shut down" in command or "shutdown" in command:  # Update condition for shutting down
        shut_down()

def close_window():
    time.sleep(1)
    active_window = gw.getActiveWindow()
    if active_window is not None:
        active_window.close()
    else:
        print("No active window found.")

def turn_on_hotspot():
    cmd = "netsh wlan set hostednetwork mode=allow ssid=Vinay's vivo key=12345678"
    
    try:
        os.system(f'runas /user:vinay bodem "{cmd}"')
        print("Hotspot turned on successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error turning on hotspot: {e}")

def turn_off_hotspot():
    cmd = 'netsh wlan stop hostednetwork'
    
    try:
        subprocess.run(['runas', '/user:vinay bodem', cmd], check=True)
        print("Hotspot turned off successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error turning off hotspot: {e}")


def search_folder(root_folder, target_folder):
    for root, dirs, files in os.walk(root_folder):
        if target_folder in dirs:
            return os.path.join(root, target_folder)
    return None

# Function to open a specified folder
def open_folder(command):
    folder_name = command.split("open folder ")[-1]
    folder_path = search_folder("C:\\", folder_name)  # Start searching from the root folder C:\
    if folder_path:
        try:
            os.startfile(folder_path)
            print(f"Opened folder: {folder_path}")
            folder_window = gw.getWindowsWithTitle(folder_name)
            if folder_window:
                folder_window[0].activate()
        except Exception as e:
            print(f"Error opening folder: {e}")
    else:
        print(f"Folder '{folder_name}' not found.")

# Function to close a specified folder
def close_folder(command):
    folder_name = command.split("close folder ")[-1]
    folder_path = search_folder("C:\\", folder_name)  # Start searching from the root folder C:\
    if folder_path:
        try:
            # Get all open windows with the title containing the folder path
            windows = gw.getWindowsWithTitle(folder_name)
            for window in windows:
                window.activate()  # Activate the window before closing
                pyautogui.hotkey("alt", "f4")  # Simulate Alt+F4 to close the active window
                print(f"Closed folder: {folder_path}")
        except Exception as e:
            print(f"Error closing folder: {e}")
    else:
        print(f"Folder '{folder_name}' not found.")

# Function to open a specified application
def open_app():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    with mic as source:
        print("Which application do you want to open?")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        print("Processing...")
        app_name = recognizer.recognize_google(audio)
        print("Opening application:", app_name)
        
        # Press the Windows key to open the Start Menu
        pyautogui.press("win")
        time.sleep(0.1)  # Add a delay to ensure the Start Menu is fully opened
        
        # Type the name of the application
        pyautogui.typewrite(app_name)
        
        # Wait for a short delay before pressing Enter
        time.sleep(0.1)
        
        # Press Enter to open the application
        pyautogui.press("enter")
        
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that.")
    except sr.RequestError:
        print("Sorry, there was an error processing the request.")

def open_command(command):
    try:
        # Remove everything up to and including "open" from the command
        command_index = command.lower().index("open") + len("open")
        command = command[command_index:].strip()
        
        # Press the Windows key to open the Start Menu
        pyautogui.press("win")
        time.sleep(0.1)  # Add a delay to ensure the Start Menu is fully opened
        
        # Type the name of the application
        pyautogui.typewrite(command)
        
        # Wait for a short delay before pressing Enter
        time.sleep(0.1)
        
        # Press Enter to open the application
        pyautogui.press("enter")
        
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that.")
    except sr.RequestError:
        print("Sorry, there was an error processing the request.")

def next_slide():
    pyautogui.press('right')
    time.sleep(1)

def previous_slide():
    pyautogui.press('left')
    time.sleep(1)

# Mapping of common application names to their corresponding executable names
APP_MAPPING = {
    "calculator": "calc.exe",
    "notepad": "notepad.exe",
    "file explorer": "explorer.exe",
    "word": "winword.exe",
    "excel": "excel.exe",
    "powerpoint": "powerpnt.exe",
    "chrome": "chrome.exe",
    "firefox": "firefox.exe",
    "edge": "msedge.exe",
    "acrobat reader": "AcroRd32.exe",
    "photoshop": "photoshop.exe",
    "illustrator": "illustrator.exe",
    "premiere pro": "premiere.exe",
    "after effects": "afterfx.exe",
    "visual studio code": "code.exe",
    "sublime text": "sublime_text.exe",
    "atom": "atom.exe",
    "notepad++": "notepad++.exe",
    "eclipse": "eclipse.exe",
    "intellij idea": "idea.exe",
    "android studio": "studio.exe",
    "spotify": "spotify.exe",
    "itunes": "itunes.exe",
    "vlc media player": "vlc.exe",
    "windows media player": "wmplayer.exe",
    "zoom": "zoom.exe",
    "skype": "skype.exe",
    "teams": "teams.exe",
    "slack": "slack.exe",
    "discord": "discord.exe",
    "outlook": "outlook.exe",
    "oneDrive": "OneDrive.exe",
    "dropbox": "Dropbox.exe",
    "google drive": "Backup and Sync.exe",
    "oneNote": "onenote.exe",
    "access": "msaccess.exe",
    "publisher": "mspub.exe",
    "visio": "visio.exe",
    "project": "winproj.exe",
    "sharepoint designer": "spdesign.exe",
    "sql server management studio": "Ssms.exe",
    "mysql workbench": "MySQLWorkbench.exe",
    "oracle sql developer": "sqldeveloper.exe",
    "paint": "mspaint.exe",
    "snipping tool": "SnippingTool.exe",
    "winrar": "WinRAR.exe",
    "virtualbox": "VirtualBox.exe",
    "docker desktop": "Docker Desktop.exe",
    "telegram": "Telegram.exe",
    "evernote": "evernote.exe",
    "office lens": "OfficeLens.exe",
    "remote desktop": "mstsc.exe",
    "filezilla": "filezilla.exe",
    "visual studio": "devenv.exe",
    "pycharm": "pycharm.exe",
    "unity": "Unity.exe",
    "blender": "blender.exe",
    "maya": "maya.exe",
    "autocad": "acad.exe",
    "matlab": "matlab.exe",
    "whatsapp" : "whatsapp.exe",
    "anaconda navigator": "anaconda-navigator.exe",
    "jupyter notebook": "jupyter-notebook.exe",
}


def close_app():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    with mic as source:
        print("Which application do you want to close?")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        print("Processing...")
        app_name = recognizer.recognize_google(audio).lower()
        print("Closing application:", app_name)
        
        # Check if the provided app_name is a key in the APP_MAPPING dictionary
        if app_name in APP_MAPPING:
            app_exe = APP_MAPPING[app_name]
            os.system("taskkill /f /im " + app_exe)
        else:
            print("Application not found in mapping:", app_name)
        
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that.")
    except sr.RequestError:
        print("Sorry, there was an error processing the request.")


# Function to maximize the currently focused window
def maximize_window():
    active_window = gw.getActiveWindow()
    if active_window:
       active_window.maximize()

# Function to minimize the currently focused window
def minimize_window():
    active_window = gw.getActiveWindow()
    if active_window:
       active_window.minimize()

# Function to take a screenshot
def take_screenshot():
    try:
        # Specify the directory where you want to save the screenshot
        screenshot_directory = r"C:\Users\vinay bodem\OneDrive\Pictures\Screenshots"
        
        # Ensure the directory exists, if not, create it
        if not os.path.exists(screenshot_directory):
            os.makedirs(screenshot_directory)

        # Generate a unique filename using current timestamp
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        screenshot_filename = f"screenshot_{timestamp}.png"

        # Construct the full path to save the screenshot
        screenshot_path = os.path.join(screenshot_directory, screenshot_filename)

        # Take the screenshot and save it with the unique filename
        pyautogui.screenshot(screenshot_path)
        print("Screenshot saved successfully:", screenshot_filename)

    except Exception as e:
        print("Error taking screenshot:", e)


def turn_on_wifi():
    # Command to turn on Wi-Fi (Windows)
    cmd = 'netsh interface set interface "Wi-Fi" admin=enabled'
    
    try:
        subprocess.run(cmd, shell=True, check=True)
        print("Wi-Fi turned on successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error turning on Wi-Fi: {e}")

def turn_off_wifi():
    # Command to turn off Wi-Fi (Windows)
    cmd = 'netsh interface set interface "Wi-Fi" admin=disabled'
    
    try:
        subprocess.run(cmd, shell=True, check=True)
        print("Wi-Fi turned off successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error turning off Wi-Fi: {e}")


# Function to switch to a specified window
def switch_window(command):
    # Extract window name from command and switch to it
    window_name = command.split("switch window to ")[-1]
    open_windows = gw.getWindowsWithTitle(window_name)
    if open_windows:
        open_windows[0].activate()
    else:
        print(f"Window '{window_name}' is not open.")
    time.sleep(1)  # Add a delay to ensure the window switch is successful

# Function to open a specified file
def search_files(file_name):
    matching_files = []
    for root, dirs, files in os.walk("C:\\"):  # Start searching from the root directory C:\
        for file in files:
            if file_name.lower() in file.lower():  # Case-insensitive match
                matching_files.append(os.path.join(root, file))
    return matching_files

# Function to prompt the user to choose a folder from the list of matching files
def select_folder(matching_files):
    print("Multiple files found with the same name:")
    for index, file_path in enumerate(matching_files, start=1):
        print(f"{index}. {file_path}")
    choice = input("Enter the number of the folder to open (0 to cancel): ")
    try:
        choice_index = int(choice) - 1
        if 0 <= choice_index < len(matching_files):
            return matching_files[choice_index]
        elif choice_index == -1:
            return None  # User chose to cancel
        else:
            print("Invalid choice. Please enter a valid number.")
            return select_folder(matching_files)  # Recursive call to re-prompt for choice
    except ValueError:
        print("Invalid input. Please enter a number.")
        return select_folder(matching_files)  # Recursive call to re-prompt for choice

# Function to open a specified file (without extension) in File Explorer if found
def open_file(command):
    file_name = command.split("open file ")[-1]
    matching_files = search_files(file_name)
    if matching_files:
        if len(matching_files) == 1:
            file_path = matching_files[0]
        else:
            file_path = select_folder(matching_files)
            if file_path is None:
                print("Operation canceled.")
                return
        subprocess.Popen(["explorer", "/select,", file_path])
        print(f"Opened file in File Explorer: {file_name}")
    else:
        print(f"File '{file_name}' not found in any directory.")


def is_prefix_present(prefix):
    prefix = prefix.lower()  # Convert the prefix to lowercase for case-insensitive comparison
    windows = [window.title.lower() for window in gw.getWindowsWithTitle("")]
    for window_title in windows:
        if prefix in window_title:
            return True
    return False

# Function to close a window with a specified prefix in its title
def close_file(command):
    prefix = command.split("close file ")[-1].strip()
    if is_prefix_present(prefix):
        for window in gw.getWindowsWithTitle(prefix):
            window.close()
        print(f"Closed file named with '{prefix}'")
    else:
        print(f"No file found with '{prefix}'")

# Function to mute the system volume
def mute_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(True, None)

def unmute_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(False, None)

# Function to set a timer
import time

# Function to set a timer
def set_timer(command):
    try:
        # Split the command into words
        words = command.split()
        
        # Find the index of the word "timer" in the command
        timer_index = words.index("timer")
        
        # Initialize timer duration
        timer_duration = None
        
        # Check if the duration is specified after "timer" directly
        if timer_index + 1 < len(words):
            next_word = words[timer_index + 1]
            
            # Attempt to convert the next word to an integer (duration)
            try:
                timer_duration = int(next_word)
            except ValueError:
                pass
        
        # Check if the duration is specified with "seconds" after "timer"
        if timer_index + 2 < len(words) and words[timer_index + 2] == "seconds":
            next_word = words[timer_index + 1]
            
            # Attempt to convert the next word to an integer (duration)
            try:
                timer_duration = int(next_word)
            except ValueError:
                pass
        
        # Check if a valid duration is found
        if timer_duration is not None:
            # Sleep for the specified duration
            time.sleep(timer_duration)
            print("Timer expired!")
        else:
            print("Invalid timer duration. Please provide a valid duration in seconds.")
    except ValueError:
        print("Invalid command format. Please use 'set timer [duration]' or 'set timer [duration] seconds'.")

# Function to open a specified website
def open_website(command):
    # Extract website URL from command and open it
    website_url = command.split("open website ")[-1]
    webbrowser.open_new_tab(website_url)

# Function to search the web using a specified search engine
def search_web(command):
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    with mic as source:
        print("What do you want to search for?")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        print("Processing...")
        search_query = recognizer.recognize_google(audio)
        print("Search Query:", search_query)
        webbrowser.open_new_tab("https://www.google.com/search?q=" + search_query)
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that.")
    except sr.RequestError:
        print("Sorry, there was an error processing the request.")


# Function to shut down the system
def shut_down():
    os.system("shutdown /s /t 1")

def listen_for_commands():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    with mic as source:
        print("Listening for commands...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        print("Processing...")
        command = recognizer.recognize_google(audio).lower()
        print("Command:", command)
        if "search web" in command:
            search_web(command)  # Call search_web() function when the command is to search the web
        else:
            execute_command(command)  # Call execute_command() for other commands
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that.")
    except sr.RequestError:
        print("Sorry, there was an error processing the request.")

# Main loop to continuously listen for commands
while True:
    listen_for_commands()
