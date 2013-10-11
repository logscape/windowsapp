strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")

host = WshNetwork.ComputerName

sep = ","

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

If IsEmpty(waitProcesses) Then
	waitProcesses = 3000
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
	WScript.Sleep(10)
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
	
	WScript.Sleep(10)
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

'Set colProcess = objWMIService.ExecQuery("SELECT Name,Priority,CommandLine FROM Win32_Process")
Set colProcess = objWMIService.ExecQuery("SELECT ProcessId,Name,Priority FROM Win32_Process")
For Each objProcess In colProcess
	PID = -1
	On Error Resume Next
	PID = objProcess.ProcessId
	ProcessName.Add PID, objProcess.Name
	ProcessPriority.Add PID, objProcess.Priority
	'CommandLine = objProcess.CommandLine
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
	
	WScript.Sleep(10)
Next

timestamp = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),3) 

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
			
			if ProcPct > 0.05 Then 
				WSCript.Echo timestamp & sep & host & sep & ProcessName.Item(processId) & sep & processId & sep & ProcPct & sep & MemPct & sep & ProcessPriority.Item(processId) & sep & ProcessOwner.Item(processId) '& sep & ElapsedTime.Item(processId) & sep & objP.CommandLine
			End If
				
			WScript.Sleep(10)
	
	End If
	
Next

WScript.Quit
