' Edit email and password
' Modify timings as necessary
Set shell = CreateObject("WScript.Shell")
shell.Run("""C:\Program Files (x86)\World of Warcraft\_retail_\Wow.exe""")
WScript.Sleep 5000
shell.SendKeys "email"
WScript.Sleep 500
shell.SendKeys "{TAB}"
WScript.Sleep 500
shell.SendKeys "password"
WScript.Sleep 500
shell.SendKeys "{ENTER}"
