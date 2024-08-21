import time
import pygetwindow as gw
import pyautogui
import os
from tkinter import messagebox
from pyautogui import FailSafeException
    
def send_keys(keys, repeat = 1, interval = 0.1):
    keywords = "ctrl shift alt win".split(" ")
    actionwords = "enter esc down up left right tab space".split(" ")

    try:
        for i in range(repeat):
            for key in keys:
                time.sleep(interval)
                
                if(key == ""):
                    continue

                if isinstance(key, int) or isinstance(key, float):
                    time.sleep(key)
                else:
                    
                    splitKeys = key.split(" ")
                    
                    if splitKeys[0] == "focus":
                        window = gw.getWindowsWithTitle(" ".join(splitKeys[1:]))
                        window[0].activate()
                        
                    elif key.split(" ")[0] in keywords:
                        if splitKeys[0] == "alt":                        
                            with pyautogui.hold('alt'):
                                for c in splitKeys[1:]:
                                    time.sleep(0.1)
                                    pyautogui.press(c)
                        else:
                            pyautogui.hotkey(splitKeys[0], splitKeys[1])
                        
                    elif splitKeys[0] in actionwords:
                        count = 1
                        if len(splitKeys) > 1:
                            count = int(splitKeys[1])
                            
                        for i in range(count):
                            time.sleep(0.01)
                            pyautogui.press(splitKeys[0])
                        
                    elif key[0] == "_":
                        _x, _y, _button = key[1:].split(",")
                        
                        pyautogui.click(x=int(_x), y=int(_y), button = _button[7:])
                    else:
                        if key[0] == "#":
                            key = key[1:]
                        
                        pyautogui.write(key, interval = 0.03)
    except Exception as e:
        if isinstance(e, FailSafeException):
            messagebox.showerror("Error", "The program has quit because the mouse has moved to the corner of the screen.")
            pyautogui.press("alt")
        else:
            messagebox.showerror("Error", "The program has quit with the following error: \n" + str(e))

        os._exit(0)
def run():
    send_keys(['focus IN-A'])


    keys_to_send = ['ctrl v', 'enter', 1.2, '_632,627,Button.left', 'ctrl c', 'focus 305298', 'right', 'ctrl v', 'enter', 'left', 'ctrl c', 'focus IN-A', 'tab']
    #['ctrl v', 'enter', 1.2, '_761,377,Button.left', 'ctrl c', 'focus 305298', 'right', 'right', 'right', 'ctrl v', 'left', 'left', 'left', 'down', 'ctrl c', 'focus IN-A', 'tab']
    iterations = 12

    for i in range(iterations):
        send_keys(keys_to_send)
        time.sleep(0.1)
