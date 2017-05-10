'#Toggles Caps Lock with delay'

'Create new object using WScript.Shell'
Set wshShell = wscript.CreateObject("WScript.Shell")
do
'Delay between every loop'
wscript.sleep 10000

'Toggle CAPSLOCK'
wshshell.sendkeys "{CAPSLOCK}"
loop