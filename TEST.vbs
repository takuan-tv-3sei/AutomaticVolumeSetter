Set wShell=CreateObject("WScript.Shell")
wShell.Run "sndvol.exe"
WScript.Sleep 2000
wShell.SendKeys "%+{TAB}"
WScript.Sleep 1000
wShell.SendKeys "{home}{esc}"
Set wShell = Nothing