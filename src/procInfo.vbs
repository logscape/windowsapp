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
query = "Select * from Win32_Process"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery(query,,48)
logMessage = ""
sep = "," 
For Each objItem in colItems
	logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4)  & sep 
	logMessage = formatDate(Now()) & sep
	logMessage = logMessage &  objItem.ProcessId & sep 
	logMessage = logMessage &  objItem.CommandLine  
	WScript.Echo logMessage
Next
