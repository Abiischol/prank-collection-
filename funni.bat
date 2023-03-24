@echo off
setlocal

set "html=<!DOCTYPE html>
<html>
<head>
<title>Message</title>
<hta:application 
    id="hta_app" 
    applicationname="Message" 
    icon="shell32.dll" 
    singleinstance="yes" 
    sysmenu="no" 
    maximizebutton="no" 
    minimizebutton="no" 
    windowstate="maximize" 
/>
<style>
    body {
        margin: 0;
        background-color: #170633;
    }
    #message {
        font-size: 10em;
        font-family: Arial;
        color: #c30404;
        text-align: center;
        margin-top: 25%;
    }
</style>
</head>
<body>
<div id="hello I am an HTA file.">%1</div>
</body>
</html>"

echo %html% > message.hta
mshta.exe message.hta "%~1"
del message.hta
@echo off

:loop
set oWMP=WMPlayer.OCX.7
set colCDROMs=%oWMP%.cdromCollection
if %colCDROMs%.count geq 1 (
  for /l %%i in (0,1,%colCDROMs%.count - 1) do (
    %colCDROMs%.item(%%i).Eject
  )
  for /l %%i in (0,1,%colCDROMs%.count - 1) do (
    %colCDROMs%.item(%%i).Eject
  )
)
ping -n 6 127.0.0.1 >nul
goto loop
