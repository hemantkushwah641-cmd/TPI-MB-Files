Set sh = CreateObject("WScript.Shell")
folder = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
desk = sh.SpecialFolders("Desktop")
Set sc = sh.CreateShortcut(desk & "\TPI MB Downloader.lnk")
sc.TargetPath = folder & "START_TPI.bat"
sc.WorkingDirectory = folder
sc.WindowStyle = 7
sc.Description = "TPI Measurement Book"
sc.Save
