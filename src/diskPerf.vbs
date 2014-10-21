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
	Dim sep, AvgN, SleepSec
	Dim DiskReadsPerSecAvg, DiskWritesPerSecAvg, DiskReadBytesPerSecAvg, DiskWriteBytesPerSecAvg, CurrentDiskQueueLengthAvg, PercentDiskReadTimeAvg, PercentDiskWriteTimeAvg, PercentIdleTimeAvg
	sep = ","
	avgN = 10: SleepSec = 0.5
	Set colItems  = data.Item("resultSet") 
	Set objRefresher = data.Item("refresher")
	'Set qList = objWMIService.ExecQuery (" SELECT Name,CurrentDiskQueueLength,DiskBytesPerSec,PercentDiskReadTime,PercentDiskWriteTime,PercentDiskTime,
	'PercentIdleTime FROM Win32_PerfFormattedData_PerfDisk_PhysicalDisk,PercentIdleTime Where Name <> '_Total'")
	objRefresher.Refresh
	For Each objItem in colItems
		DiskReadsPerSecAvg=0: DiskWritesPerSecAvg=0: DiskReadBytesPerSecAvg=0: DiskWriteBytesPerSecAvg=0
		CurrentDiskQueueLengthAvg=0.0: PercentDiskReadTimeAvg=0.0: PercentDiskWriteTimeAvg=0.0: PercentIdleTimeAvg=0.0
		line = "" 
		If filterCondition(objItem) = True Then
			line = line & objItem.Name & sep
			'if pidExists(pids,objItem.IDProcess) Then
			REM line = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep 
			line = formatDate(Now())
			deviceId= objItem.Name
			line = line & replace(deviceId," ","_") & sep 
			'ak-- average values over AvgN*SleepSec sec interval, taking AvgN readings
			For i = 1 to AvgN
				objRefresher.Refresh
				DiskReadsPerSecAvg = DiskReadsPerSecAvg + objItem.DiskReadsPerSec
				DiskWritesPerSecAvg = DiskWritesPerSecAvg + objItem.DiskWritesPerSec
				DiskReadBytesPerSecAvg = DiskReadBytesPerSecAvg + objItem.DiskReadBytesPerSec
				DiskWriteBytesPerSecAvg = DiskWriteBytesPerSecAvg + objItem.DiskWriteBytesPerSec
				CurrentDiskQueueLengthAvg = CurrentDiskQueueLengthAvg + objItem.CurrentDiskQueueLength
				PercentDiskReadTimeAvg = PercentDiskReadTimeAvg + objItem.PercentDiskReadTime
				PercentDiskWriteTimeAvg = PercentDiskWriteTimeAvg + objItem.PercentDiskWriteTime
				PercentIdleTimeAvg = PercentIdleTimeAvg + objItem.PercentIdleTime
				Wscript.Sleep 1000*SleepSec
			Next
			DiskReadsPerSecAvg = DiskReadsPerSecAvg / AvgN
			DiskWritesPerSecAvg = DiskWritesPerSecAvg / AvgN
			DiskReadBytesPerSecAvg = DiskReadBytesPerSecAvg / AvgN
			DiskWriteBytesPerSecAvg = DiskWriteBytesPerSecAvg / AvgN
			CurrentDiskQueueLengthAvg = CurrentDiskQueueLengthAvg / AvgN
			PercentDiskReadTimeAvg = PercentDiskReadTimeAvg / AvgN
			PercentDiskWriteTimeAvg = PercentDiskWriteTimeAvg / AvgN
			PercentIdleTimeAvg = PercentIdleTimeAvg / AvgN
			' Read  / Write Operations
			line = line & DiskReadsPerSecAvg & sep 
			line = line & DiskWritesPerSecAvg & sep 
			' Read / Write Bytes / per sec 
			line = line & DiskReadBytesPerSecAvg & sep 
			line = line & DiskWriteBytesPerSecAvg & sep 				
			' Percentage Performance Counters
			line = line & CurrentDiskQueueLengthAvg & sep 
			line = line & PercentDiskReadTimeAvg & sep 
			line = line & PercentDiskWriteTimeAvg   & sep 
			line = line & PercentIdleTimeAvg 
			WScript.echo line 
			'End If 
		End If 
	'	if objItem.PercentProcessorTime > 0 Then
	'	        Wscript.Echo Now()  & " " & objItem.Name & " -- " & objItem.PercentProcessorTime
	'	End If 
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

