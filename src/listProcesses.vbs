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

Function getProperty(properties,key,defaultValue)
	Dim value
	If properties.Exists(key) = true Then
		value=properties(key)
	else
		value=defaultValue
	End If
	getProperty=value
End Function

'----------------------------------------------------------------------------------------------------------------------------
'Functions Processing Section
'----------------------------------------------------------------------------------------------------------------------------
'Name       : getPropertiesFromArguments -> Creates a Dictionary from logscape properties
'Parameters : None          ->
'Return     : Dictionary    ->
'----------------------------------------------------------------------------------------------------------------------------
Function getPropertiesFromArguments()
	Dim i
	Dim properties
	Set properties=CreateObject("Scripting.Dictionary")
	For i=0 to WScript.Arguments.Count() - 1 

		If InStr(WScript.Arguments(i),"=") Then
			Dim elems
			elems=Split(WScript.Arguments(i),"=")
			properties.Add elems(0),elems(1)
		End If 
	Next 
	Set getPropertiesFromArguments=properties 
End Function



Function filterProcessName(excludes,code)
	Dim res
	res=0
	excludes=LCase(excludes)
	code=LCase(code)
	If InStr(excludes,","&code) > 0 Then
		res=1
	End If
	
	If InStr(excludes,code&",") > 0 Then
		res=1
	End If
	
	If excludes=CStr(code) Then 
		
		res=1
		
	End If 
	
	filterProcessName=res
End Function

Dim properties
Set properties = getPropertiesFromArguments()
doNotFilter=getProperty(properties,"doNotFilter","")
	

strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")

host = WshNetwork.ComputerName

sep = ","

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

If IsEmpty(waitProcesses) Then
	waitProcesses = 100
End If

Set TimeStamp1 = CreateObject("Scripting.Dictionary")
Set PercentProcessorTime1 = CreateObject("Scripting.Dictionary")


Set objCPU = objWMIService.Get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
TS1 = objCPU.TimeStamp_Sys100NS
PPT1 = objCPU.PercentProcessorTime

Set colProcess = objWMIService.ExecQuery("SELECT Name,IDProcess,TimeStamp_Sys100NS,PercentProcessorTime FROM Win32_PerfRawData_PerfProc_Process WHERE Name <> 'Idle'")
For Each objProcess in colProcess
	TimeStamp1.Add objProcess.IDProcess, objProcess.TimeStamp_Sys100NS
	PercentProcessorTime1.Add objProcess.IDProcess, objProcess.PercentProcessorTime
'	WScript.Sleep(1)
Next

WScript.Sleep(waitProcesses)

Set TimeStamp2 = CreateObject("Scripting.Dictionary")
Set PercentProcessorTime2 = CreateObject("Scripting.Dictionary")
'Set ElapsedTime = CreateObject("Scripting.Dictionary")
Set WorkingSetPrivate = CreateObject("Scripting.Dictionary")

Set objCPU = objWMIService.Get("Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
objCPU.Refresh_
TS2 = objCPU.TimeStamp_Sys100NS
PPT2 = objCPU.PercentProcessorTime

Set colProcess = objWMIService.ExecQuery("SELECT IDProcess,TimeStamp_Sys100NS,PrivateBytes,PercentProcessorTime,ElapsedTime,WorkingSetPrivate FROM Win32_PerfRawData_PerfProc_Process WHERE Name <> 'Idle'")
For Each objProcess in colProcess
	TimeStamp2.Add objProcess.IDProcess, objProcess.TimeStamp_Sys100NS
	PercentProcessorTime2.Add objProcess.IDProcess, objProcess.PercentProcessorTime
	'ElapsedTime.Add objProcess.IDProcess, objProcess.ElapsedTime
	WorkingSetPrivate.Add objProcess.IDProcess, objProcess.WorkingSetPrivate
	
'	WScript.Sleep(1)
Next

Set PerfOS = objWMIService.Get("Win32_PerfFormattedData_PerfOS_Memory=@")
PerfOS.Refresh_
UsedMem = CDbl(PerfOS.PercentCommittedBytesInUse)

DT = Abs(CDbl(TS2 - TS1))
DP = Abs(CDbl(PPT2 - PPT1))

UsedProc = -1
If DT > 0 Then
	UsedProc = Round((1 - DP / DT) * 100, 2)
Else
	UsedProc = 0	
End If

If UsedProc < 0 Then
	UsedProc = 0
End If
If UsedProc > 100 Then
	UsedProc = 100
End If

DeltaTimeTotal = Abs(CDbl(TimeStamp2.Item(0) - TimeStamp1.Item(0)))
DeltaProcTotal = Abs(CDbl(PercentProcessorTime2.Item(0) - PercentProcessorTime1.Item(0)))
TotalMemory = CDbl(WorkingSetPrivate.Item(0))

Set ProcessPriority = CreateObject("Scripting.Dictionary")
Set ProcessName = CreateObject("Scripting.Dictionary")
Set ProcessOwner = CreateObject("Scripting.Dictionary")
'Set ProcessCmdLine = CreateObject("Scripting.Dictionary")

'Set colProcess = objWMIService.ExecQuery("SELECT Name,Priority,CommandLine FROM Win32_Process")
Set colProcess = objWMIService.ExecQuery("SELECT ProcessId,Name,Priority,CommandLine FROM Win32_Process")
For Each objProcess In colProcess
	PID = -1
	On Error Resume Next
	PID = objProcess.ProcessId
	ProcessName.Add PID, objProcess.Name
	ProcessPriority.Add PID, objProcess.Priority
'	ProcessCmdLine.Add PID, objProcess.CommandLine
	Ret = 1
	If PID <> -1 Then
		Ret = objProcess.GetOwner(User,Domain)
	End If
	If Err.number<>0 Then
		Err.Clear
	End If
	On Error Goto 0
	If Ret <> 0 Then
		User = "?"
		Domain = "?"
	End If
	ProcessOwner.Add PID, Domain & "\" & User
	
	WScript.Sleep(1)
Next

REM timestamp = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) 
timestamp = formatDate(Now())
For Each processId In TimeStamp1.Keys

	If processId <> 0 AND ProcessPriority.Item(processId) <> "" Then

		ProcPct = -1
		If (DeltaProcTotal > 0) Then
			ProcPct = Abs(PercentProcessorTime2.Item(processId) - PercentProcessorTime1.Item(processId)) / DeltaProcTotal
		Else
			ProcPct = 0
		End If
		If ProcPct < 0 Then
			ProcPct = 0
		End If
		If ProcPct > 1 Then
			ProcPct = 1
		End If
		ProcPct = Round(UsedProc * ProcPct,2)
		
		If TotalMemory = 0 Then
			MemPct = 0
		Else
			MemPct = Round(WorkingSetPrivate.Item(processId) / TotalMemory * UsedMem, 2)
		End If

		procName=ProcessName.Item(processId)

		if ProcPct > 0.05 Then 
				WSCript.Echo timestamp & sep & host _
				& sep & ProcessName.Item(processId) _
				& sep & processId & sep & ProcPct _
				& sep & MemPct & sep & ProcessPriority.Item(processId) _
				& sep & ProcessOwner.Item(processId)
	'			& sep & ProcessCmdLine.Item(processId)
		elseif filterProcessName(doNotFilter,procName)	> 0 Then
				WSCript.Echo timestamp & sep & host _
				& sep & ProcessName.Item(processId) _
				& sep & processId & sep & ProcPct _
				& sep & MemPct & sep & ProcessPriority.Item(processId) _
				& sep & ProcessOwner.Item(processId)
	'			& sep & ProcessCmdLine.Item(processId)
		else
		End If
	End If
	
Next

WScript.Quit
