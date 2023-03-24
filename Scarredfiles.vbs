Option Explicit

Dim fso, shell, desktop, windows
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
desktop = shell.SpecialFolders("Desktop")
Set windows = CreateObject("Shell.Application").Windows

Dim cursorX, cursorY
GetCursorPos cursorX, cursorY

Do
    For Each window In windows
        If InStr(1, window.FullName, "explorer.exe", vbTextCompare) > 0 Then
            Dim windowX, windowY, windowWidth, windowHeight
            windowX = window.Left
            windowY = window.Top
            windowWidth = window.Width
            windowHeight = window.Height
            If cursorX > windowX And cursorX < windowX + windowWidth And _
               cursorY > windowY And cursorY < windowY + windowHeight Then
                window.Left = windowX - 10
                window.Top = windowY - 10
            End If
        End If
    Next
    
    Set windows = CreateObject("Shell.Application").Windows
    
    WScript.Sleep 50
Loop
