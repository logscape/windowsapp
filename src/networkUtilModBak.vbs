strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")

sep = ","

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

If IsEmpty(waitNetwork) Then
	'waitNetwork = 1.5 * 60 * 1000
	waitNetwork = 3000
End If
	
If IsEmpty(maxSpeedMbps) Then
	maxSpeedMbps = 100
End If

'Set Formatted = CreateObject("Scripting.Dictionary")
Set Raw = CreateObject("Scripting.Dictionary")

Set colNI = objWMIService.ExecQuery("SELECT Name FROM Win32_PerfFormattedData_Tcpip_NetworkInterface")
For Each objNI in colNI
	'WScript.Echo "Adding: " & objNI.Name
	'Formatted.Add objNI.Name, objWMIService.Get("Win32_PerfFormattedData_Tcpip_NetworkInterface.Name='" & objNI.Name & "'")
	Raw.Add objNI.Name, objWMIService.Get("Win32_PerfRawData_Tcpip_NetworkInterface.Name='" & objNI.Name & "'")
Next
'WScript.Echo " "

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
For Each sNetInterf In Raw.Keys
	Raw.Item(sNetInterf).Refresh_
	TimeStamp1.Add sNetInterf, Raw.Item(sNetInterf).Timestamp_Sys100NS
	ErrorsIn1.Add sNetInterf, Raw.Item(sNetInterf).PacketsReceivedErrors
	ErrorsOut1.Add sNetInterf, Raw.Item(sNetInterf).PacketsOutboundErrors
	PacketsIn1.Add sNetInterf, Raw.Item(sNetInterf).PacketsReceivedPerSec
	PacketsOut1.Add sNetInterf, Raw.Item(sNetInterf).PacketsSentPerSec
	BytesIn1.Add sNetInterf, Raw.Item(sNetInterf).BytesReceivedPerSec
	BytesOut1.Add sNetInterf, Raw.Item(sNetInterf).BytesSentPerSec
Next

WScript.Sleep(waitNetwork)

'Get values 2
For Each sNetInterf In Raw.Keys
	Raw.Item(sNetInterf).Refresh_
	TimeStamp2.Add sNetInterf, Raw.Item(sNetInterf).Timestamp_Sys100NS
	ErrorsIn2.Add sNetInterf, Raw.Item(sNetInterf).PacketsReceivedErrors
	ErrorsOut2.Add sNetInterf, Raw.Item(sNetInterf).PacketsOutboundErrors
	PacketsIn2.Add sNetInterf, Raw.Item(sNetInterf).PacketsReceivedPerSec
	PacketsOut2.Add sNetInterf, Raw.Item(sNetInterf).PacketsSentPerSec
	BytesIn2.Add sNetInterf, Raw.Item(sNetInterf).BytesReceivedPerSec
	BytesOut2.Add sNetInterf, Raw.Item(sNetInterf).BytesSentPerSec
Next

host = WshNetwork.ComputerName
timestamp = Now()

For Each key In Raw.Keys

	DeltaTimeStamp = (TimeStamp2.Item(key) - TimeStamp1.Item(key)) * 100 / 1000 / 1000 / 1000
	ErrorsIn = ErrorsIn2.Item(key) - ErrorsIn1.Item(key)
	ErrorsOut = ErrorsOut2.Item(key) - ErrorsOut1.Item(key)
	PacketsIn = PacketsIn2.Item(key) - PacketsIn1.Item(key)
	PacketsOut = PacketsOut2.Item(key) - PacketsOut1.Item(key)
	'BytesIn = BytesIn2.Item(key) - BytesIn1.Item(key)
	'BytesOut = BytesOut2.Item(key) - BytesOut1.Item(key)
	
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
	
	WScript.Echo timestamp & sep & host & sep & key & sep & BytesPerSecIn & sep& BytesPerSecOut & sep & ErrorsIn & sep & ErrorsOut & sep & Rx & sep & Tx & sep & maxSpeedMbps
	
Next



'WScript.Echo " "
'WScript.Echo "DONE"

WScript.Quit
