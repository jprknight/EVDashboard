'********************************************************************************
'* EV DASHBOARD EV DISK INFO START
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing EV Disk Info."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing EV Disk Info."
End if

'Get today's date.
Styr = Year(Now)
Stmo = Month(Now)
if Stmo < 10 Then
	Stmo = "0" & Stmo
end if
Stdt = Day(Now)
if Stdt < 10 Then
	Stdt = "0" & Stdt
end if

'* Required Const declaration in EVDashboard_Data_Collector.vbs
'Const adCmdText = 1

On Error Resume Next

For i = 0 to Ubound(EVServersList,2)
	
	Dim objWMIService, objItem, colItems
	Dim strDriveType, strDiskSize

	If LoggingLevel >= 1 then
		wscript.echo DateTime() & ": (EV Disk Info) - " & EVServersList(1,i)
		OutputLogFile.writeline DateTime() & ": (EV Disk Info) - " & EVServersList(1,i)
	End if
	
	Set objWMIService = GetObject("winmgmts:\\" & EVServersList(1,i) & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume WHERE DriveType=3 AND Capacity > 1073741824")
	
	For Each objItem in colItems
			
		DIM pctFreeSpace,strFreeSpace,strusedSpace
		
		pctFreeSpace = INT(((objItem.FreeSpace / objItem.Capacity) * 1000)/10)' & "%"
		strDiskSize = Int(objItem.Capacity /1073741824)' & "GiB"
		strFreeSpace = Int(objItem.FreeSpace /1073741824)' & "GiB"
		strUsedSpace = Int((objItem.Capacity-objItem.FreeSpace)/1073741824)' & "GiB"
		
		If EVDConnectionString = 1 Then
			Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
		Elseif EVDConnectionString = 2 Then
			Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
		End if

		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (EV Disk Info) - Connection1" & vbcrlf & Connection1
			OutputLogFile.writeline DateTime() & ": (EV Disk Info) - Connection1" & vbcrlf & Connection1
		End if
		
		'* Open database connection
		Set myconn1 = CreateObject("adodb.connection")
		myconn1.ConnectionTimeout = strSQLConnect
		myconn1.open(Connection1)

		strSQLQuery1 = "UPDATE dbo.DiskUsageInfo " & _
			"SET EVServer='" & EVServersList(1,i) & "'," & _
			"EVRole='" & EVServersList(5,i) & "'," & _
			"Drive='" & objItem.Name & "'," & _
			"VolumeName='" & objItem.Label & "'," & _
			"SizeGB='" & strDiskSize & "'," & _
			"UsedGB='" & strUsedSpace & "'," & _
			"FreeGB='" & strFreeSpace & "'," & _
			"PercentageFree='" & pctFreeSpace & "' " & _
			"WHERE EVServer='" & EVServersList(1,i) & "' AND " & _
			"Drive='" & objItem.Name & "' AND " & _
			"RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
			"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' " & _
			"IF @@ROWCOUNT=0 " & _
			"INSERT INTO dbo.DiskUsageInfo " & _
			"(EVServer,EVRole,Drive,VolumeName,SizeGB,UsedGB,FreeGB,PercentageFree) " & _
			"VALUES ('" & EVServersList(1,i) & "','" & EVServersList(5,i) & "','" & objItem.Name & _
			"','" & objItem.Label & "','" & strDiskSize & "','" & strUsedSpace & "','" & strFreeSpace & _
			"','" & pctFreeSpace & "');"

		'* VERBOSE Logging.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (EV Disk Info - strSQLQuery1) " & vbcrlf & strSQLQuery1
			OutputLogFile.writeline DateTime() & ": (EV Disk Info - strSQLQuery1) " & strSQLQuery1
		End if			
			
		Set objCmd1 = CreateObject("adodb.command")
		objCmd1.activeconnection = myconn1
		objCmd1.commandtimeout = strSQLExecute
		objCmd1.commandtype = adCmdText
		objCmd1.commandtext = strSQLQuery1

		Set result1 = CreateObject("adodb.Recordset")
		result1.open objCmd1
	Next
Next
'********************************************************************************
'* EV DASHBOARD EV DISK INFO END
'********************************************************************************