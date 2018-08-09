'********************************************************************************
'* EV DASHBOARD EXCHANGE DISK INFO START
'********************************************************************************
On Error Resume Next

'* MINIMAL LOGGING.
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Exchange Disk Info."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Exchange Disk Info."
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

If EVDConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (Exchange Disk Info) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (Exchange Disk Info) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "select ExchangeServer " & _
	"from ExchangeMailboxes " & _
	"Group by ExchangeServer " & _
	"Order By ExchangeServer"

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": (Exchange Disk Info) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Exchange Disk Info) - strSQLQuery1" & vbcrlf & strSQLQuery1
End if	
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

'************************************
'* Write data.
While not result1.EOF
	Dim objWMIService, objItem, colItems
	Dim strDriveType, strDiskSize
	
	CurrentEXServer = (result1("ExchangeServer"))
	
	If LoggingLevel >= 1 then
		wscript.echo DateTime() & ": (Exchange Disk Info) - " & CurrentEXServer
		OutputLogFile.writeline DateTime() & ": (Exchange Disk Info) - " & CurrentEXServer
	End if
	
	Set objWMIService = GetObject("winmgmts:\\" & CurrentEXServer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume WHERE DriveType=3 AND Capacity > 1073741824")
	
	For Each objItem in colItems
			
		DIM pctFreeSpace,strFreeSpace,strusedSpace
		
		pctFreeSpace = INT(((objItem.FreeSpace / objItem.Capacity) * 1000)/10)' & "%"
		strDiskSize = Int(objItem.Capacity /1073741824)' & "GiB"
		strFreeSpace = Int(objItem.FreeSpace /1073741824)' & "GiB"
		strUsedSpace = Int((objItem.Capacity-objItem.FreeSpace)/1073741824)' & "GiB"
		
		If EVDConnectionString = 1 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
		Elseif EVDConnectionString = 2 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
		End if

		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Exchange Disk Info) - Connection2" & vbcrlf & Connection2
			OutputLogFile.writeline DateTime() & ": (Exchange Disk Info) - Connection2" & vbcrlf & Connection2
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)

		strSQLQuery2 = "UPDATE dbo.DiskUsageInfo " & _
			"SET EVServer='" & CurrentEXServer & "'," & _
			"EVRole='Email'," & _
			"Drive='" & objItem.Name & "'," & _
			"VolumeName='" & objItem.Label & "'," & _
			"SizeGB='" & strDiskSize & "'," & _
			"UsedGB='" & strUsedSpace & "'," & _
			"FreeGB='" & strFreeSpace & "'," & _
			"PercentageFree='" & pctFreeSpace & "' " & _
			"WHERE EVServer='" & CurrentEXServer & "' AND " & _
			"Drive='" & objItem.Name & "' AND " & _
			"RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
			"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' " & _
			"IF @@ROWCOUNT=0 " & _
			"INSERT INTO dbo.DiskUsageInfo " & _
			"(EVServer,EVRole,Drive,VolumeName,SizeGB,UsedGB,FreeGB,PercentageFree) " & _
			"VALUES ('" & CurrentEXServer & "','Email','" & objItem.Name & _
			"','" & objItem.Label & "','" & strDiskSize & "','" & strUsedSpace & "','" & strFreeSpace & _
			"','" & pctFreeSpace & "');"
		
		'* VERBOSE Logging.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Exchange Disk Info - strSQLQuery2) " & vbcrlf & strSQLQuery2
			OutputLogFile.writeline DateTime() & ": (Exchange Disk Info - strSQLQuery2) " & strSQLQuery2
		End if
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2

		Set result2 = CreateObject("adodb.Recordset")
		result2.open objCmd2
	Next
	result1.movenext()
Wend
'********************************************************************************
'* EV DASHBOARD EXCHANGE DISK INFO FINISH
'********************************************************************************