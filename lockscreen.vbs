
Set WshShell = WScript.CreateObject("WScript.Shell")
 
WScript.Sleep 5000
wshShell.SendKeys "{NUMLOCK}"
WScript.Sleep 500
wshShell.SendKeys "{NUMLOCK}"

for i=1 to 180
	WScript.Sleep 180000
	wshShell.SendKeys "{NUMLOCK}"
	WScript.Sleep 500
	wshShell.SendKeys "{NUMLOCK}"
next
