On Error Resume Next

Function pad(num)
	if num < 10 then
		num="0"&num
	end if 
	pad=num
End Function

Function formatDate(dt)
	ss=pad(Second(dt))
	mm=pad(Minute(dt))
	hh=pad(hour(dt))
	dd=pad(Day(dt))
	MM=pad(Month(dt))
	YY=Year(dt)
	formatDate=dd&"/"&MM&"/"&YY&" "&hh&":"&mm&":"&ss
End Function


strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Objects",,48)
sep = "," 
REM logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep  
logMessage = formatDate(Now()) & sep
REM events,mutexes,processes,semephores,threads
For Each objItem in colItems
	logMessage = logMessage & objItem.Events  & sep 
	logMessage = logMessage & objItem.Mutexes & sep 
	logMessage =  logMessage & objItem.Processes & sep 
	logMessage =  logMessage & objItem.Semaphores & sep 
	logMessage =  logMessage & objItem.Threads
Next
WScript.Echo logMessage
