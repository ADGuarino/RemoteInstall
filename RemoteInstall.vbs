'Option Explicit
Dim objWMI, objItem, colItems
Dim strComputer, VerOS, VerBig, Ver9x, Version9x, OS, OSystem
Dim WshShell

Set WshShell = CreateObject("WScript.Shell")

Set WshNetwork = Wscript.CreateObject("WScript.Network")
strProfile = WshNetwork.UserName

'strComputer = "."
strComputer = InputBox("Computer Name.")
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

		Dim fso
		Set fso = CreateObject("scripting.filesystemobject")
	
			If Not fso.FileExists("C:\Windows\System32\PSexec.exe") Then
				' WshShell.Run " 'Enter Path to SysInternals Here' C:\Windows\System32\ " 
				MsgBox("You do not have PStools installed. Now installing... Please accept the user agreement on first use.")
			End If
	
		Set WshShell = WScript.CreateObject("WScript.Shell")
		' WshShell.Run "PSexec.exe \\" & strComputer & " -s -u TEC\" & strProfile & " cmd /K 'Enter Path to silent install script here' "
		'WScript.Echo "Rollback procedure Is complete."
