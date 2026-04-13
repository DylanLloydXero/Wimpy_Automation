import os
import json
import time
import urllib.parse
import webbrowser
import pyautogui
import win32clipboard
import win32gui

def copy_files_to_clipboard(filepaths):
    """Copies file paths to the Windows clipboard as a 'File Drop' (CF_HDROP)."""
    abs_paths = [os.path.abspath(f) for f in filepaths]
    for attempt in range(5):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            # CF_HDROP format
            win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, abs_paths)
            win32clipboard.CloseClipboard()
            break
        except Exception as e:
            time.sleep(0.2)

def check_whatsapp_ready():
    """Checks if the WhatsApp Desktop window is open and gives it focus."""
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd).lower()
            if "whatsapp" in title:
                windows.append(hwnd)
    
    windows = []
    win32gui.EnumWindows(callback, windows)
    
    if not windows:
        return False, "WhatsApp Desktop is not open. Please open it first."
    
    try:
        win32gui.SetForegroundWindow(windows[0])
        time.sleep(1)
        return True, "WhatsApp is ready."
    except Exception as e:
        return False, f"Could not focus WhatsApp: {e}"

def start_bot(payload_path):
    if not os.path.exists(payload_path):
        print(f"Payload not found: {payload_path}")
        return

    ready, msg = check_whatsapp_ready()
    if not ready:
        pyautogui.alert(msg, "WhatsApp Bot Error")
        return

    with open(payload_path, 'r') as f:
        data = json.load(f)
        
    print("--- WHATSAPP DESKTOP BOT STARTED ---")
    print("!!! DO NOT TOUCH YOUR MOUSE OR KEYBOARD UNTIL FINISHED !!!")
    time.sleep(3)

    for item in data:
        phone = item.get('phone', '')
        phone = ''.join(filter(str.isdigit, phone))
        msg = item.get('message', '')
        filepath = item.get('file', '')
        
        if not phone:
            print(f"Skipping entry with no phone number.")
            continue
            
        print(f"Targeting: {phone}")
        url = f"whatsapp://send?phone={phone}"
        webbrowser.open(url)
        time.sleep(4)
        
        if filepath and os.path.exists(filepath):
            print(f"Attaching: {filepath}")
            copy_files_to_clipboard([filepath])
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1.5) 
            
        if msg:
            pyautogui.write(msg)
            time.sleep(0.5)
            
        pyautogui.press('enter')
        print(f"Sent to {phone}")
        time.sleep(2)

    print("--- ALL MESSAGES SENT ---")
    pyautogui.alert("WhatsApp Automation Complete!", "Wimpy De Ville Manager")

if __name__ == '__main__':
    start_bot("output/whatsapp_payload.json")
