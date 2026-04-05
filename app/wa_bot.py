import os
import json
import time
import urllib.parse
import webbrowser
import pyautogui
import win32clipboard

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

def start_bot(payload_path):
    if not os.path.exists(payload_path):
        print(f"Payload not found: {payload_path}")
        return

    with open(payload_path, 'r') as f:
        data = json.load(f)
        
    print("--- WHATSAPP DESKTOP BOT STARTED ---")
    print("!!! DO NOT TOUCH YOUR MOUSE OR KEYBOARD UNTIL FINISHED !!!")
    time.sleep(3)

    for item in data:
        phone = item.get('phone', '')
        # Remove any non-digits if necessary, but protocol usually needs digits
        phone = ''.join(filter(str.isdigit, phone))
        msg = item.get('message', '')
        filepath = item.get('file', '')
        
        if not phone:
            print(f"Skipping entry with no phone number.")
            continue
            
        print(f"Targeting: {phone}")
        
        # 1. Open the Desktop App via protocol
        # We don't include text in URL to avoid URL length issues and focus on the paste
        url = f"whatsapp://send?phone={phone}"
        webbrowser.open(url)
        
        # 2. Wait for App to focus and Load chat
        # Usually 3-4 seconds is safe for the desktop app to switch chats
        time.sleep(4)
        
        # 3. Handle File Attachment (PDF)
        if filepath and os.path.exists(filepath):
            print(f"Attaching: {filepath}")
            copy_files_to_clipboard([filepath])
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1.5) # Wait for preview to appear
            
        # 4. Type/Paste Message
        if msg:
            # We use typewrite or paste. Typewrite is safer for emojis/formatting sometimes
            # But pasting is faster for long messages.
            # Let's just type it if it's short, or assume it's in the text box if we used text in URL
            # Since we didn't use text in URL, we type it now.
            pyautogui.write(msg)
            time.sleep(0.5)
            
        # 5. Send
        pyautogui.press('enter')
        print(f"Sent to {phone}")
        
        # Wait between people to avoid spam blocks and let UI catch up
        time.sleep(2)

    print("--- ALL MESSAGES SENT ---")
    pyautogui.alert("WhatsApp Automation Complete!", "Wimpy De Ville Manager")

if __name__ == '__main__':
    start_bot("output/whatsapp_payload.json")
