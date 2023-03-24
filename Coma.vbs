Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace("C:\Users\yourusername\Desktop")

MsgBox "awooooo 56709"

Do While True
    For Each objItem In objFolder.Items
        Randomize
        If Int(Rnd * 2) = 0 Then
            objItem.ExtendedProperty("System.Size") = "64"
        Else
            objItem.ExtendedProperty("System.Size") = "128"
        End If
    Next
    WScript.Sleep 500
Loop
