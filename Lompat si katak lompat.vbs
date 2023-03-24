Option Explicit

Dim fso, path, auto, AQ
Set fso = CreateObject("Scripting.FileSystemObject")
path = CreateObject("WScript.Shell").SpecialFolders("Desktop")
Set auto = CreateObject("Shell.Application")
Set AQ = CreateObject("WScript.Shell")

Do
    Dim g1, g2, g11, g12
    g1 = path + "\frog1"
    g2 = path + "\frog2"
    If fso.FileExists(g1) Then
        Set g11 = fso.GetFile(g1)
        If g11.Attributes <> 39 Then
            g11.Attributes = 0
            auto.Copy g1, True
        End If
    Else
        auto.Copy g1, True
    End If
    If fso.FileExists(g2) Then
        Set g12 = fso.GetFile(g2)
        If g12.Attributes <> 39 Then
            g12.Attributes = 0
            AQ.Copy g2, True
        End If
    End If
Loop Until EOF(1)
