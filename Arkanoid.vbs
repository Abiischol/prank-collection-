Option Explicit

Dim objWMIService, objEventSource, objEventHandler
Dim colMonitoredEvents, objLatestEvent

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objEventSource = objWMIService.ExecNotificationQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Desktop'")


Sub MonitorDesktopIcons()
    Do
        Set objLatestEvent = objEventSource.NextEvent()
        If objLatestEvent.TargetInstance.PartComponent Like "*""*""*" Then
            Dim strIconPath
            strIconPath = Split(objLatestEvent.TargetInstance.PartComponent, """")(1)
            Call RicochetIcon(strIconPath)
        End If
    Loop
End Sub


Sub RicochetIcon(strIconPath)
    Dim intLeft, intTop, intXSpeed, intYSpeed
    Dim objShell, objFolder, objFolderItem
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(WScript.ScriptFullName).ParentFolder
    Set objFolderItem = objFolder.ParseName(strIconPath)
    
    intXSpeed = 10
    intYSpeed = 20
    
    ' Get the initial position of the icon
    intLeft = objFolderItem.ExtendedProperty("System.ItemPositionLeft")
    intTop = objFolderItem.ExtendedProperty("System.ItemPositionTop")
    
    Do
        ' Update the position of the icon
        intLeft = intLeft + intXSpeed
        intTop = intTop + intYSpeed
        
        ' Make the icon bounce off the edges of the screen
        If intLeft < 0 Then
            intLeft = 0
            intXSpeed = -intXSpeed
        End If
        
        If intTop < 0 Then
            intTop = 0
            intYSpeed = -intYSpeed
        End If
        
        If intLeft > 8000 Then
            intLeft = 8000
            intXSpeed = -intXSpeed
        End If
        
        If intTop > 8000 Then
            intTop = 8000
            intYSpeed = -intYSpeed
        End If
        
        ' Update the position of the icon on the desktop
        objFolderItem.ExtendedProperty("System.ItemPositionLeft") = intLeft
        objFolderItem.ExtendedProperty("System.ItemPositionTop") = intTop
        
        WScript.Sleep 50
    Loop
End Sub


' Call the MonitorDesktopIcons subroutine to start monitoring desktop icons
MonitorDesktopIcons()
