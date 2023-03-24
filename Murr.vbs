Option Explicit

On Error Resume Next

Dim telnet
Set telnet = CreateObject("MSWinsock.Winsock")
telnet.Connect "furrymuck.com", 8888

If Err.Number <> 0 Then
    Dim wshShell, x, y
    Set wshShell = WScript.CreateObject("WScript.Shell")
    wshShell.Run "rundll32.exe user32.dll,UpdatePerUserSystemParameters", 1, True
    Do
        WScript.Sleep 50
        x = wshShell.MousePosition.X
        y = wshShell.MousePosition.Y
        wshShell.Run "rundll32.exe user32.dll,UpdatePerUserSystemParameters", 1, True
        wshShell.SendKeys "{ESC}"
    Loop Until Err.Number = 0
Else
    While telnet.State <> 7
        WScript.Sleep 100
    Wend

    Dim user, pass, character
    user = "YourFurryMUCKUsername"
    pass = "YourFurryMUCKPassword"
    character = "YourFurryMUCKCharacter"

    telnet.SendData user & " " & pass & " " & character & Chr(13)
    WScript.Sleep 500

    While True
        Dim data
        data = telnet.GetData
        If Len(data) > 0 Then
            WScript.Echo data
        End If
        WScript.Sleep 100
    Wend

    telnet.Close
End If

Set telnet = Nothing
