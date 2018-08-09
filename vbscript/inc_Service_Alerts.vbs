'********************************************************************************
'* SERVICE ALERTS START
'********************************************************************************
'* MINIMAL LOGGING.
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Service Alerts."
	If LoggingLevel >= 2 then
		OutputLogFile.writeline ""
		OutputLogFile.writeline "*************************************************************************"
		OutputLogFile.writeline ""
	End if
	OutputLogFile.writeline DateTime() & ": Processing Service Alerts."
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

Dim strWriteDataCount: strWriteDataCount = 0

'**********************************************
'* Remove any previously records for the current day.
If EVDConnectionString = 1 Then
	EVDconnection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	EVDconnection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (Service Alerts) - EVDConnection1" & vbcrlf & EVDconnection1
	OutputLogFile.writeline DateTime() & ": (Service Alerts) - EVDConnection1" & vbcrlf & EVDconnection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(EVDconnection1)

strSQLQuery1 = "Delete FROM dbo.ServiceAlerts " & _
	"WHERE RecordCreateTimestamp > '" & styr & "/" & stmo & "/" & stdt & " 00:00:00';"

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": (Service Alerts) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Service Alerts) - strSQLQuery1" & vbcrlf & strSQLQuery1
End if

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

On Error Resume Next

For i = 0 to Ubound(EVServersList,2)
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & EVServersList(0,i) & "\root\cimv2")
	Set colRunningServices = objWMIService.ExecQuery _
		("Select * from Win32_Service")
	For Each objService in colRunningServices
		Select case objService.DisplayName 
			Case "Enterprise Vault Admin Service"
				If objService.State <> "Running" then			
					Call WriteData
					strWriteDataCount = strWriteDataCount + 1
				End if
			Case "Enterprise Vault Directory Service"
				If objService.State <> "Running" then			
					Call WriteData
					strWriteDataCount = strWriteDataCount + 1
				End if
			Case "Enterprise Vault Indexing Service"
				If objService.State <> "Running" then			
					Call WriteData
					strWriteDataCount = strWriteDataCount + 1
				End if
			Case "Enterprise Vault Shopping Service"
				If objService.State <> "Running" then			
					Call WriteData
					strWriteDataCount = strWriteDataCount + 1
				End if
			Case "Enterprise Vault Storage Service"
				If objService.State <> "Running" then			
					Call WriteData
					strWriteDataCount = strWriteDataCount + 1
				End if
			Case "Enterprise Vault Task Controller Service"
				If objService.State <> "Running" then			
					Call WriteData
					strWriteDataCount = strWriteDataCount + 1
				End if
			Case "Message Queuing"
				If objService.State <> "Running" then			
					Call WriteData
					strWriteDataCount = strWriteDataCount + 1
				End if
		End Select
	Next
Next

'* MINIMAL LOGGING.
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": (Service Alerts) - Inserted " & strWriteDataCount & " new row(s) into database."
	OutputLogFile.writeline DateTime() & ": (Service Alerts) - Inserted " & strWriteDataCount & " new row(s) into database."
End if

'********************************************************************************
Function WriteData
	If EVDConnectionString = 1 Then
		EVDconnection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
			EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
	Elseif EVDConnectionString = 2 Then
		EVDconnection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
			EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
	End if
	
	'* NORMAL LOGGING.
	If LoggingLevel >= 2 then
		'wscript.echo DateTime() & ": (Service Alerts) - EVDconnection2" & vbcrlf & EVDconnection2
		OutputLogFile.writeline DateTime() & ": (Service Alerts) - EVDconnection2" & vbcrlf & EVDconnection2
	End if
	
	'* Open database connection
	Set myconn2 = CreateObject("adodb.connection")
	myconn2.ConnectionTimeout = strSQLConnect
	myconn2.open(EVDconnection2)

	strSQLQuery2 = "INSERT INTO dbo.ServiceAlerts " & _
		"(DNSAlias,Server,Role,Service,Status) " & _
		"VALUES ('" & EVServersList(0,i) & "','" & EVServersList(1,i) & "','" & EVServersList(5,i) & _
		"','" & objService.DisplayName & "','" & objService.State & "');"

	'* VERBOSE LOGGING.
	If LoggingLevel >= 3 then
		'wscript.echo DateTime() & ": (Service Alerts) - strSQLQuery2" & vbcrlf & strSQLQuery2
		OutputLogFile.writeline DateTime() & ": (Service Alerts) - strSQLQuery2" & vbcrlf & strSQLQuery2
	End if
	
	Set objCmd2 = CreateObject("adodb.command")
	objCmd2.activeconnection = myconn2
	objCmd2.commandtimeout = strSQLExecute
	objCmd2.commandtype = adCmdText
	objCmd2.commandtext = strSQLQuery2

	Set result2 = CreateObject("adodb.Recordset")
	result2.open objCmd2
End Function
'********************************************************************************
'* SERVICE ALERTS END
'********************************************************************************