import psutil  

# if "Buddy.exe" in (p.name() for p in psutil.process_iter()):
#         print("Already running closing it")
#         os.system("TASKKILL /F /IM Buddy.exe")

for p in psutil.process_iter():
    if p.name() == "Outlook.exe":
        print(p.children)

# import win32gui


# def enumWindowsProc(hwnd, lParam):
#     print(win32gui.GetWindowText(hwnd))

# win32gui.EnumWindows(enumWindowsProc, 0)