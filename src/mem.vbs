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
	mth=pad(Month(dt))
	YY=Year(dt)
	formatDate=dd&"/"&mth&"/"&YY&" "&hh&":"&mm&":"&ss
End Function


Function filterCondition(obj)
	ret = True

	if obj.Name = "_Total" or obj.Name = "Idle"  Then
		ret = False 
	End If 


	filterCondition = ret
End Function

Function pidExists(dict,key)
        ret = false
        For each k in dict.Keys
                val1 = k + 0
                val2 = key + 0
                if val1 = val2  Then
                        ret = true
                End If
        Next
        pidExists =ret
End Function

Function counters(service)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

	'service =  "Win32_PerfFormattedData_PerfProc_Process"

	Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
	Set rs = objRefresher.AddEnum (objWMIService,service).objectSet 
'    (objWMIService, service).objectSet
	Set ret = CreateObject("Scripting.Dictionary")
	ret.Add "service",service
	ret.Add "resultSet", rs
	ret.Add "refresher",objRefresher
	Set counters = ret 
End Function 

Sub log(data,pids)
	Set colItems  = data.Item("resultSet") 
	Set objRefresher = data.Item("refresher")

	sep = ";"
	'For i = 1 to 2 
	    objRefresher.Refresh
	    For Each objItem in colItems
		line = "" 
		If filterCondition(objItem) = True Then
	'		line = line  &   objItem.Name & sep
			if pidExists(pids,objItem.IDProcess) Then
				REM line = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4)  & sep 
				line = formatDate(Now())& sep 
				line = line & objItem.Name & sep 
				line = line & objItem.IDProcess & sep 
				line = line & objItem.PrivateBytes & sep 
				line = line & objItem.WorkingSet & sep 
				line = line & objItem.PageFaultsPersec   & sep 
				line = line & objItem.PercentProcessorTime  
				line = line & objItem.ThreadCount  
				WScript.echo line 
			End If 
		End If 
	'	if objItem.PercentProcessorTime > 0 Then
	'	        Wscript.Echo Now()  & " " & objItem.Name & " -- " & objItem.PercentProcessorTime
	'	End If 
	    Next
	'    Wscript.Echo
	'    Wscript.Sleep 5000
	'Next
End Sub

Function dotnetpids(data)
	Set pids = CreateObject("Scripting.Dictionary")
	Set objRefresher = data("refresher")
	Set colItems = data("resultSet")
	objRefresher.Refresh

	For Each objItem in colItems
		pids.Add objItem.ProcessID, ""
	Next
 
	Set dotnetpids = pids 
End Function 

Set clrData = counters("Win32_PerfFormattedData_NETFramework_NETCLRMemory") 
Set pids = dotnetpids(clrData) 
Set data = counters("Win32_PerfFormattedData_PerfProc_Process")

log data, pids
