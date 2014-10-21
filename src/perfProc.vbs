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

Function filterCondition(obj)
	ret = True 
	If obj.Name = "_Total"  Then
		ret = False
	End If 

	If obj.IOReadBytesPerSec = 0 Then
		If obj.IOWriteBytesPerSec = 0 Then
			ret = False 
		End If 
	End If 
	filterCondition = ret 
End Function 

strComputer = "."

Set refresher = CreateObject("WbemScripting.SWbemRefresher")
Set services = GetObject("winmgmts:\\" _
    & strComputer & "\root\cimv2")
Set objRefreshableItem = _
    refresher.AddEnum(services , _
    "Win32_PerfFormattedData_PerfProc_Process")
Set colItems = objRefreshableItem.ObjectSet
refresher.Refresh

sep = "," 

While True
	For Each objItem in colItems

		If filterCondition(objItem) = True then
			REM logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep  
			logMessage = formatDate(Now()) & sep 
			logMessage =  logMessage &  objItem.Name  & sep 
			logMessage =  logMessage &  objItem.IDProcess  & sep 
			logMessage =  logMessage &  objItem.ThreadCount & sep 
			logMessage =  logMessage &  objItem.HandleCount & sep 
			logMessage =  logMessage &  objItem.IOReadBytesPerSec & sep 
			logMessage =  logMessage &  objItem.IOWriteBytesPerSec  
			WScript.Echo logMessage
		End If 
	Next
	WScript.Sleep 60000
Wend 
