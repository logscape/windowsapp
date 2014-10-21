Option Explicit
Dim objWMIService, objProcess, colProcess, qList
Dim strComputer, strList,qItem,WshNetWork, sep, cName

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
Set WshNetwork = WScript.CreateObject("WScript.Network")
cName = WshNetwork.ComputerName
sep = ","

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

Set qList = objWMIService.ExecQuery ("SELECT PagesPerSec,AvailableMBytes,CommittedBytes,PercentCommittedBytesInUse FROM Win32_PerfFormattedData_PerfOS_Memory")
For Each qItem in qList
		REM WSCript.Echo FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep & cName & sep & qItem.PagesPerSec & sep & qItem.AvailableMBytes & sep & qItem.CommittedBytes & sep & qItem.PercentCommittedBytesInUse

		WSCript.Echo formatDate(Now()) & sep & cName & sep & qItem.PagesPerSec & sep & qItem.AvailableMBytes & sep & qItem.CommittedBytes & sep & qItem.PercentCommittedBytesInUse
Next

WScript.Quit
