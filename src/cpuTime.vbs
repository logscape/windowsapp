


sub getSamples(count)
	Set objService = GetObject( _
		"Winmgmts:{impersonationlevel=impersonate}!\Root\Cimv2")

		
		
	For i = 1 to count
		Set objInstance1 = objService.Get( _
			"Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
		N1   = objInstance1.PercentProcessorTime
		D1   = objInstance1.TimeStamp_Sys100NS
		PUT1 = objInstance1.PercentUserTime
		PPT1 = objInstance1.PercentPrivilegedTime
		PIT1 = objInstance1.PercentInterruptTime

	'Sleep for two seconds = 2000 ms
		WScript.Sleep(2000)

		Set objInstance2 = objService.Get( _
			"Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
		N2 = objInstance2.PercentProcessorTime
		D2 = objInstance2.TimeStamp_Sys100NS
		PUT2 = objInstance2.PercentUserTime
		PPT2 = objInstance2.PercentPrivilegedTime
		PIT2 = objInstance2.PercentInterruptTime
		
		
		DeltaTime = Abs(CDbl(D2 - D1))

		PercentProcessorTime = -1
		PercentUserTime= -1
		PercentPrivilegedTime = -1
		PercentInterruptTime = -1
 
		
		If DeltaTime > 0 Then
			PercentProcessorTime = Round((1 - ( N2 - N1) / (D2-D1)) * 100, 2)
			PercentUserTime = Round((Abs(PUT2 - PUT1) / (D2-D1)) * 100, 2)
			PercentPrivilegedTime = Round((Abs(PPT2 - PPT1) / (D2-D1)) * 100, 2)
			PercentInterruptTime = Round((Abs(PIT2 - PIT1) / (D2-D1)) * 100, 2)
		End If	
		
	' Look up the CounterType qualifier for the PercentProcessorTime 
	' and obtain the formula to calculate the meaningful data. 
	' CounterType - PERF_100NSEC_TIMER_INV
	' Formula - (1- ((N2 - N1) / (D2 - D1))) x 100

		If PercentProcessorTime < 0 Then
			PercentProcessorTime = 0
		End If
		If PercentUserTime < 0 Then
			PercentUserTime = 0
		End If
		If PercentPrivilegedTime < 0 Then
			PercentPrivilegedTime = 0
		End If
		If PercentInterruptTime < 0 Then
			PercentInterruptTime = 0
		End If

		If PercentProcessorTime > 100 Then
			PercentProcessorTime = 100
		End If
		If PercentUserTime > 100 Then
			PercentUserTime = 100
		End If
		If PercentPrivilegedTime > 100 Then
			PercentPrivilegedTime = 100
		End If
		If PercentInterruptTime > 100 Then
			PercentInterruptTime = 100
		End If

		WSCript.Echo FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),3) & sep & WshNetwork.ComputerName & sep & PercentProcessorTime & sep & PercentUserTime & sep & PercentPrivilegedTime & sep & PercentInterruptTime


		REM PercentProcessorTime = (1 - ((N2 - N1)/(D2-D1)))*100
		REM WScript.Echo "% Processor Time=" , Round(PercentProcessorTime,2)
	Next
	
End Sub	


strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")

 sep = ","
 numOfSamples=4
 intervalSecs=15
 numberOfIntervals=3

For i = 1 to 3

	getSamples(numOfSamples)
	WScript.Sleep(15000)
	getSamples(numOfSamples)
Next 
	
