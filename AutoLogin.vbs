Set WshShell = WScript.CreateObject("WScript.Shell")

'Path of Cisco AnyConnect Client'
WshShell.Run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe""" 

WScript.Sleep 800

WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"

WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
'VPN URL'
WshShell.SendKeys "URL"
WshShell.SendKeys "{ENTER}"

WScript.Sleep 1500

'PASSWORD'
WshShell.SendKeys "PASSWORD"
WScript.Sleep 200
WshShell.SendKeys "+{TAB}"
'USERNAME'
WshShell.SendKeys "USERNAME"
WScript.Sleep 200
WshShell.SendKeys "{ENTER}"