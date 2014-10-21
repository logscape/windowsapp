' cputime.vbs
'ak-- 26/11/13 -- modified to improve averaging and reduce frequency/amount of log data

Dim objInst
Dim pppAvg, pptAvg, putAvg, pitAvg
Dim cName, sep, strComputer
Dim numOfSamples, intervalSecs, numberOfIntervals

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

sub getSamples(count)
	Set objService = GetObject( _
		"Winmgmts:{impersonationlevel=impersonate}!\Root\Cimv2")
		
	pppAvg=0.0: pptAvg=0.0: putAvg=0.0: pitAvg=0.0
	For i = 1 to count
		Set objInst = objService.Get( _
			"Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
		N1   = objInst.PercentProcessorTime
		D1   = objInst.TimeStamp_Sys100NS
		PUT1 = objInst.PercentUserTime
		PPT1 = objInst.PercentPrivilegedTime
		PIT1 = objInst.PercentInterruptTime

	'Sleep for 200 ms
		WScript.Sleep(200)

		Set objInst = objService.Get( _
			"Win32_PerfRawData_PerfOS_Processor.Name='_Total'")
		N2 = objInst.PercentProcessorTime
		D2 = objInst.TimeStamp_Sys100NS
		PUT2 = objInst.PercentUserTime
		PPT2 = objInst.PercentPrivilegedTime
		PIT2 = objInst.PercentInterruptTime
		
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

		pppAvg = pptAvg + PercentProcessorTime
		putAvg = putAvg + PercentUserTime
		pptAvg = pptAvg + PercentPrivilegedTime
		pitAvg = pitAvg + PercentInterruptTime

		'WSCript.Echo FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep & cName & sep & PercentProcessorTime & sep & PercentUserTime & sep & PercentPrivilegedTime & sep & PercentInterruptTime

		REM  PercentProcessorTime = (1 - ((N2 - N1)/(D2-D1)))*100
		REM WScript.Echo "% Processor Time=" , Round(PercentProcessorTime,2)
	Next

	pppAvg=pppAvg/count: pptAvg=pptAvg/count: putAvg=putAvg/count: pitAvg=pitAvg/count
	REM WSCript.Echo FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),4) & sep & cName & sep & pppAvg & sep & putAvg & sep & pptAvg & sep & pitAvg
	WScript.Echo formatDate(Now()) & " " & FormatDateTime(Now(),4) & sep & cName & sep & pppAvg & sep & putAvg & sep & pptAvg & sep & pitAvg
	
End Sub	

strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")

cName = WshNetwork.ComputerName
sep = ","
numOfSamples=5
intervalSecs=15
numberOfIntervals=1

For i = 1 to numberOfIntervals
	getSamples(numOfSamples)
	'WScript.Sleep(15000)
	'getSamples(numOfSamples)
Next 
