
On Error Resume Next
strComputer = "."
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PageFileUsage",,48)
logMessage = ""
sep = ","
For Each objItem in colItems
	logMessage = FormatDateTime(Now(),2) & " " & FormatDateTime(Now(),3) & sep 
	logMessage = logMessage  & WshNetwork.ComputerName   & sep 
    logMessage = logMessage  & objItem.AllocatedBaseSize  & sep 
    logMessage = logMessage  & objItem.CurrentUsage  & sep 
    logMessage = logMessage  & objItem.Description  & sep 
    logMessage = logMessage  & objItem.Name  & sep 
    logMessage = logMessage  & objItem.PeakUsage  & sep 
    logMessage = logMessage  & objItem.Status  & sep 
Next
WScript.echo logMessage
