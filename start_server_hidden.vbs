Set WshShell = CreateObject("WScript.Shell")
strDir = "C:\Users\karma\Desktop\GENRATOR 2026"
WshShell.CurrentDirectory = strDir
strCmd = """C:\Users\karma\AppData\Local\Python\pythoncore-3.14-64\pythonw.exe"" """ & strDir & "\server_watchdog.py"""
WshShell.Run strCmd, 0, False
