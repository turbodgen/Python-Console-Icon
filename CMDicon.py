from PIL import Image
import inspect, os, time
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")

# Accidentally made most of these F strings. Oh well.
def CMDico(chosenIco = None, windowName = "Python Window"):
    if not os.path.exists(f"{windowName}.lnk"):
        icoP = Image.open(chosenIco)
        if not chosenIco.lower().endswith(".ico"):
            try:
                icoP.save('icon.ico',format = 'ICO', sizes=[(32,32)])
            except:
                print("Error with ICO file!")
                return
        importingFile = inspect.stack()[1].filename 
        batf = open(f"{windowName}.bat","w")
        batf.writelines(["@echo off","\ncls",f"\npython {importingFile}"])
        batf.close()
        shortcut = shell.CreateShortcut(f"{windowName}.lnk")
        shortcut.TargetPath = os.path.abspath(f"{windowName}.bat")
        shortcut.IconLocation = os.path.abspath(chosenIco)
        shortcut.Save()
        os.startfile(f"{windowName}.lnk")
        time.sleep(0.3)
        os.remove(f"{windowName}.lnk")
        exit()

