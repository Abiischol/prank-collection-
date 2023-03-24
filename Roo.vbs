X=msgbox "kangaroo"

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Set wshShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
Set objWshScriptExec = wshShell.Exec("mshta.exe about:""<html><head><script>moveTo(-32000,-32000);</script></head><body></body></html>""")
Set objExecStdOut = objWshScriptExec.StdOut
strHTAPath = objExecStdOut.ReadLine
Set objHTA = objShell.Windows.Item(strHTAPath)

objHTA.document.title = "Virus"
objHTA.document.body.innerHTML = "<center><h1>Virus</h1></center>"
Do
    x = Rnd * (wshShell.AppActivate("Virus") + 1) - 50
    y = Rnd * (wshShell.AppActivate("Virus") + 1) - 50
    objHTA.Left = x
    objHTA.Top = y
    objHTA.Refresh
    wshShell.Sleep 50
Loop
