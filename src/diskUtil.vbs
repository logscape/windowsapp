Option Explicit
Dim objWMIService, objProcess, colProcess, qList
Dim strComputer, strList,qItem,WshNetWork
Dim sep

strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")

sep = ","

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

Set qList = objWMIService.ExecQuery ("SELECT Name,FreeMegaBytes,PercentFreeSpace FROM Win32_PerfFormattedData_PerfDisk_LogicalDisk")
For Each qItem in qList
	WSCript.Echo FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep & WshNetwork.ComputerName & sep & qItem.Name & sep & qItem.FreeMegaBytes & sep & qItem.PercentFreeSpace
Next

WScript.Quit
