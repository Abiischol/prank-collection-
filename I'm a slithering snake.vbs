' Set up the snake's starting position and speed
Dim snakeX, snakeY, deltaX, deltaY
snakeX = 0
snakeY = 0
deltaX = 10
deltaY = 10

' Start the animation loop
Do While True
    ' Move the snake
    snakeX = snakeX + deltaX
    snakeY = snakeY + deltaY
    
    ' Check if the snake has hit a screen edge
    If snakeX < 0 Then
        deltaX = -deltaX
        snakeX = 0
    ElseIf snakeX + 50 > Screen.Width Then
        deltaX = -deltaX
        snakeX = Screen.Width - 50
    End If
    
    If snakeY < 0 Then
        deltaY = -deltaY
        snakeY = 0
    ElseIf snakeY + 50 > Screen.Height Then
        deltaY = -deltaY
        snakeY = Screen.Height - 50
    End If
    
    ' Draw the snake
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "mspaint.exe"
    WScript.Sleep 500
    objShell.AppActivate "Paint"
    WScript.Sleep 500
    objShell.SendKeys "%fn"
    WScript.Sleep 500
    objShell.SendKeys "n"
    WScript.Sleep 500
    objShell.SendKeys "{DOWN " & snakeY & "}"
    objShell.SendKeys "{RIGHT " & snakeX & "}"
    objShell.SendKeys "{DOWN 40}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{LEFT}"
    objShell.SendKeys "{DOWN}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{DOWN}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{DOWN}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{ENTER}"
    objShell.SendKeys "{UP}"
    objShell.SendKeys "{UP}"
    objShell.SendKeys "{UP}"
Loop
