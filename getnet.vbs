strInput = InputBox("Wifi Name (You need to have connected to this network before!):","GetNet")
Dim oShell
Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.Run "cmd /K netsh wlan show profiles " & strInput & " key=clear"
Set oShell = Nothing
