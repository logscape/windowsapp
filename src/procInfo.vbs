On Error Resume Next

strComputer = "."
query = "Select * from Win32_Process"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery(query,,48)
logMessage = ""
sep = "," 
For Each objItem in colItems
	logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),3)  & sep 
	logMessage = logMessage &  objItem.ProcessId & sep 
	logMessage = logMessage &  objItem.CommandLine  
	WScript.Echo logMessage
Next
