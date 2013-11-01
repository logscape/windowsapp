Option Explicit
Dim objWMIService, objProcess, colProcess, qList
Dim strComputer, strList,qItem,WshNetWork, sep

strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")

sep = ","

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

Set qList = objWMIService.ExecQuery ("SELECT PagesPerSec,AvailableMBytes,CommittedBytes,PercentCommittedBytesInUse FROM Win32_PerfFormattedData_PerfOS_Memory")
For Each qItem in qList
		WSCript.Echo FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep & WshNetwork.ComputerName & sep & qItem.PagesPerSec & sep & qItem.AvailableMBytes & sep & qItem.CommittedBytes & sep & qItem.PercentCommittedBytesInUse
Next

WScript.Quit
