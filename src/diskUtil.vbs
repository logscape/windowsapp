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
	if obj.Name = "_Total"  Then
		ret = False 
	End If 
	if Instr(obj.Name,"HarddiskVolume") <> 0  Then
		ret = False 
	End If 
	filterCondition = ret
End Function

Function counters(service)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
	Set rs = objRefresher.AddEnum (objWMIService,service).objectSet 
	Set ret = CreateObject("Scripting.Dictionary")
	ret.Add "service",service
	ret.Add "resultSet", rs
	ret.Add "refresher",objRefresher
	Set counters = ret 
End Function 

REM Set qList = objWMIService.ExecQuery ("SELECT Name,FreeMegaBytes,PercentFreeSpace FROM Win32_PerfFormattedData_PerfDisk_LogicalDisk")
REM For Each qItem in qList
	REM WSCript.Echo FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep & WshNetwork.ComputerName & sep & qItem.Name & sep & qItem.FreeMegaBytes & sep & qItem.PercentFreeSpace
REM Next

REM WScript.Quit

Sub log(data)
	Dim sep, AvgN, SleepSec
	Dim FreeMegaBytesAvg, PercentFreeSpaceAvg
	sep = ","
	avgN = 3: SleepSec = 0.3
	Set colItems  = data.Item("resultSet") 
	Set objRefresher = data.Item("refresher")
	'ak-- average values over AvgN*SleepSec sec interval, taking AvgN readings
	objRefresher.Refresh
	For Each objItem in colItems
		FreeMegaBytesAvg = 0.0: PercentFreeSpaceAvg = 0.0
		line = "" 
		If filterCondition(objItem) = True Then
			line = line & objItem.Name & sep
			REM line = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep 
			line = formatDate(Now()) & sep 
			deviceId= objItem.Name
			line = line & replace(deviceId," ","_") & sep 
			'ak-- average values over AvgN*SleepSec sec interval, taking AvgN readings
			For i = 1 to AvgN
				objRefresher.Refresh
				FreeMegaBytesAvg = FreeMegaBytesAvg + objItem.FreeMegaBytes
				PercentFreeSpaceAvg = PercentFreeSpaceAvg + objItem.PercentFreeSpace
				Wscript.Sleep 1000*SleepSec
			Next
			FreeMegaBytesAvg = FreeMegaBytesAvg / AvgN
			PercentFreeSpaceAvg = PercentFreeSpaceAvg / AvgN
			line = line & FreeMegaBytesAvg   & sep 
			line = line & PercentFreeSpaceAvg
			WScript.echo line 
		End If
	Next
End Sub	


REM ' Main
Set data = counters("Win32_PerfFormattedData_PerfDisk_LogicalDisk")

log data
