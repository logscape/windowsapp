strComputer = "."
sep = ","

Set WshNetwork = WScript.CreateObject("WScript.Network")
host = WshNetwork.ComputerName

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

If IsEmpty(waitNetwork) Then
	waitNetwork = 500
End If
	
'If IsEmpty(maxSpeedMbps) Then
'	maxSpeedMbps = 100
'End If

Set CurrentBW = CreateObject("Scripting.Dictionary")

Set TimeStamp1 = CreateObject("Scripting.Dictionary")
Set ErrorsIn1 = CreateObject("Scripting.Dictionary")
Set ErrorsOut1 = CreateObject("Scripting.Dictionary")
Set PacketsIn1 = CreateObject("Scripting.Dictionary")
Set PacketsOut1 = CreateObject("Scripting.Dictionary")
Set BytesIn1 = CreateObject("Scripting.Dictionary")
Set BytesOut1 = CreateObject("Scripting.Dictionary")

Set TimeStamp2 = CreateObject("Scripting.Dictionary")
Set ErrorsIn2 = CreateObject("Scripting.Dictionary")
Set ErrorsOut2 = CreateObject("Scripting.Dictionary")
Set PacketsIn2 = CreateObject("Scripting.Dictionary")
Set PacketsOut2 = CreateObject("Scripting.Dictionary")
Set BytesIn2 = CreateObject("Scripting.Dictionary")
Set BytesOut2 = CreateObject("Scripting.Dictionary")

'Get values 1
Set colProcess = objWMIService.ExecQuery("SELECT Name,TimeStamp_Sys100NS,PacketsReceivedErrors,PacketsOutboundErrors,PacketsReceivedPerSec,PacketsSentPerSec,BytesReceivedPerSec,BytesSentPerSec,CurrentBandwidth FROM Win32_PerfRawData_Tcpip_NetworkInterface")
For Each objProcess in colProcess
	sNetInterf = objProcess.Name
	TimeStamp1.Add objProcess.Name, objProcess.Timestamp_Sys100NS
	ErrorsIn1.Add sNetInterf, objProcess.PacketsReceivedErrors
	ErrorsOut1.Add sNetInterf, objProcess.PacketsOutboundErrors
	PacketsIn1.Add sNetInterf, objProcess.PacketsReceivedPerSec
	PacketsOut1.Add sNetInterf, objProcess.PacketsSentPerSec
	BytesIn1.Add sNetInterf, objProcess.BytesReceivedPerSec
	BytesOut1.Add sNetInterf, objProcess.BytesSentPerSec
	'WScript.Sleep(1)
Next

WScript.Sleep(waitNetwork)

'Get values 2
Set colProcess = objWMIService.ExecQuery("SELECT Name,TimeStamp_Sys100NS,PacketsReceivedErrors,PacketsOutboundErrors,PacketsReceivedPerSec,PacketsSentPerSec,BytesReceivedPerSec,BytesSentPerSec,CurrentBandwidth FROM Win32_PerfRawData_Tcpip_NetworkInterface")
For Each objProcess in colProcess
	sNetInterf = objProcess.Name
	TimeStamp2.Add sNetInterf, objProcess.Timestamp_Sys100NS
	ErrorsIn2.Add sNetInterf, objProcess.PacketsReceivedErrors
	ErrorsOut2.Add sNetInterf, objProcess.PacketsOutboundErrors
	PacketsIn2.Add sNetInterf, objProcess.PacketsReceivedPerSec
	PacketsOut2.Add sNetInterf, objProcess.PacketsSentPerSec
	BytesIn2.Add sNetInterf, objProcess.BytesReceivedPerSec
	BytesOut2.Add sNetInterf, objProcess.BytesSentPerSec
	CurrentBW.Add sNetInterf, objProcess.CurrentBandwidth
	'WScript.Sleep(1)
Next

timestamp = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4)
On Error Resume Next
For Each key In TimeStamp1.Keys
	CurrentBandWidth = 1000*round(0.001*CurrentBW.Item(key)/(1024*1024),2)
	If CurrentBandWidth > 0 then
		DeltaTimeStamp = (TimeStamp2.Item(key) - TimeStamp1.Item(key)) * 100 / 1000 / 1000 / 1000
		ErrorsIn = ErrorsIn2.Item(key) - ErrorsIn1.Item(key)
		ErrorsOut = ErrorsOut2.Item(key) - ErrorsOut1.Item(key)
		PacketsIn = PacketsIn2.Item(key) - PacketsIn1.Item(key)
		PacketsOut = PacketsOut2.Item(key) - PacketsOut1.Item(key)
		BytesIn = BytesIn2.Item(key) - BytesIn1.Item(key)
		BytesOut = BytesOut2.Item(key) - BytesOut1.Item(key)
		'ErrorsInPerSec = Round(ErrorsIn / DeltaTimeStamp, 2)
		'ErrorsOutPerSec = Round(ErrorsOut / DeltaTimeStamp, 2)
		'PacketsPerSecIn = Round(PacketsIn / DeltaTimeStamp, 2)
		'PacketsPerSecOut = Round(PacketsOut / DeltaTimeStamp, 2)
		BytesPerSecIn = Round(BytesIn / DeltaTimeStamp, 2)
		BytesPerSecOut = Round(BytesOut / DeltaTimeStamp, 2)
		If PacketsIn = 0 Then
			Rx = 1
		Else
			Rx = (PacketsIn - ErrorsIn) / PacketsIn
		End If
		If PacketsOut = 0 Then
			Tx = 1
		Else
			Tx = (PacketsOut - ErrorsOut) / PacketsOut
		End If
		WScript.Echo timestamp & sep & host & sep & key & sep & BytesPerSecIn & sep& BytesPerSecOut & sep & ErrorsIn & sep & ErrorsOut & sep & Rx & sep & Tx & sep & CurrentBandWidth
	End If
Next



'WScript.Echo " "
'WScript.Echo "DONE"

WScript.Quit
