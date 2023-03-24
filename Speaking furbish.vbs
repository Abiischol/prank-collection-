Rem this Furby will keep on printing itself
Do
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad.exe"
WScript.Sleep 500
WshShell.SendKeys "  __,='`````'=/__{ENTER}//  (o) \\_\\{ENTER}||    ,_\\/_\\   ){ENTER}\\\     ''  /  /{ENTER} \\\'''\\--''/ //{ENTER}    \\\`-/:-'\/-'/ {ENTER}     \\\ '' /' /\\{ENTER}      \\\'/'/ /{ENTER}       \\\_-/ /-/{ENTER}"
Loop
Rem until you restart
