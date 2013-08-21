On Error Resume Next

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
			logMessage = Now() & sep  
			logMessage =  logMessage &  objItem.Name  & sep 
			logMessage =  logMessage &  objItem.IDProcess  & sep 
			logMessage =  logMessage &  objItem.ThreadCount & sep 
			logMessage =  logMessage &  objItem.HandleCount & sep 
			logMessage =  logMessage &  objItem.IOReadBytesPerSec & sep 
			logMessage =  logMessage &  objItem.IOWriteBytesPerSec  
			WScript.Echo logMessage
		End If 
	Next
	WScript.Sleep 5000
Wend 
