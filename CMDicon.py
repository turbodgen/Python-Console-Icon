# turbodgen
# only for windows

import inspect, os, time
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")

def CMDico(chosenIco = None, windowName = "Python Window"):
    if not os.path.exists(f"{windowName}.lnk"):
        # if not an ico file
        if not chosenIco.lower().endswith(".ico"):
            if not os.path.exists(os.path.splitext(chosenIco)[0]+".ico"):
                from PIL import Image
                icoP = Image.open(chosenIco)
                try:
                    icoP.save(f'{os.path.splitext(chosenIco)[0]}.ico',format = 'ICO', bitmap_format='bmp')
                except:
                    print("CMDicon: Error with ICO file!")
                    return
        chosenIco = os.path.splitext(chosenIco)[0]+".ico"
        # creating shortcut file
        importingFile = inspect.stack()[1].filename 
        batf = open(f"{windowName}.bat","w")
        batf.writelines(["@echo off","\ncls",f"\npython {os.path.basename(importingFile)}"])
        batf.close()
        shortcut = shell.CreateShortcut(f"{windowName}.lnk")
        shortcut.TargetPath = os.path.abspath(f"{windowName}.bat")
        shortcut.IconLocation = os.path.abspath(chosenIco)
        shortcut.Save()
        os.startfile(f"{windowName}.lnk")
        time.sleep(0.3)
        os.remove(f"{windowName}.lnk")
        exit()