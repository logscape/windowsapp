Function filterCondition(obj)
	ret = True

	if obj.Name = "_Total"  Then
		ret = False 
	End If 



	if Instr(obj.Name,"HarddiskVolume") <> 0  Then
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


Sub log(data)
	Set colItems  = data.Item("resultSet") 
	Set objRefresher = data.Item("refresher")
	'Set qList = objWMIService.ExecQuery (" SELECT Name,CurrentDiskQueueLength,DiskBytesPerSec,PercentDiskReadTime,PercentDiskWriteTime,PercentDiskTime,
	'PercentIdleTime FROM Win32_PerfFormattedData_PerfDisk_PhysicalDisk,PercentIdleTime Where Name <> '_Total'")
	
	sep = ";"
	For i = 1 to 10 
	    objRefresher.Refresh
	    For Each objItem in colItems
		line = "" 
		If filterCondition(objItem) = True Then
			line = line  &   objItem.Name & sep
			'if pidExists(pids,objItem.IDProcess) Then
				line = Now()  & sep 
				deviceId= objItem.Name
				line = line & replace(deviceId," ","_") & sep 
				
				' Read  / Write Operations
				line = line & objItem.DiskReadsPerSec & sep 
				line = line & objItem.DiskWritesPerSec & sep 

				' Read / Write Bytes / per sec 
				line = line & objItem.DiskReadBytesPerSec & sep 
				line = line & objItem.DiskWriteBytesPerSec & sep 				


				' Percentage Performance Counters
				line = line & objItem.CurrentDiskQueueLength & sep 
				line = line & objItem.PercentDiskReadTime & sep 
				line = line & objItem.PercentDiskWriteTime   & sep 
				line = line & objItem.PercentIdleTime 
				WScript.echo line 
			'End If 
		End If 
	'	if objItem.PercentProcessorTime > 0 Then
	'	        Wscript.Echo Now()  & " " & objItem.Name & " -- " & objItem.PercentProcessorTime
	'	End If 
	    Next
	    Wscript.Echo
	    Wscript.Sleep 1000
	Next
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

'Set clrData = counters("Win32_PerfFormattedData_NETFramework_NETCLRMemory") 
'Set pids = dotnetpids(clrData) 
'Set data = counters("Win32_PerfFormattedData_PerfProc_Process")
Set data = counters("Win32_PerfFormattedData_PerfDisk_LogicalDisk")

log data

