On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Objects",,48)
sep = "," 
logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),3) & sep  
REM events,mutexes,processes,semephores,threads
For Each objItem in colItems
	logMessage = logMessage & objItem.Events  & sep 
	logMessage = logMessage & objItem.Mutexes & sep 
	logMessage =  logMessage & objItem.Processes & sep 
	logMessage =  logMessage & objItem.Semaphores & sep 
	logMessage =  logMessage & objItem.Threads
Next
WScript.Echo logMessage
