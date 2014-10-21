
On Error Resume Next
strComputer = "."

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

Set WshNetwork = WScript.CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PageFileUsage",,48)
logMessage = ""
sep = ","
For Each objItem in colItems
	REM logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep 
	logMessage = formatDate(Now()) & sep 
	logMessage = logMessage  & WshNetwork.ComputerName   & sep 
    logMessage = logMessage  & objItem.AllocatedBaseSize  & sep 
    logMessage = logMessage  & objItem.CurrentUsage  & sep 
    logMessage = logMessage  & objItem.Description  & sep 
    logMessage = logMessage  & objItem.Name  & sep 
    logMessage = logMessage  & objItem.PeakUsage  & sep 
    logMessage = logMessage  & objItem.Status  & sep 
Next
WScript.echo logMessage
