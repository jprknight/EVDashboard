'********************************************************************************
'* MSMQ Messages Count Alerts START
'********************************************************************************
On Error Resume Next

'* MINIMAL LOGGING.
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing MSMQ Message Count Alerts."
	If LoggingLevel >= 2 then
		OutputLogFile.writeline ""
		OutputLogFile.writeline "*************************************************************************"
		OutputLogFile.writeline ""
	End if
	OutputLogFile.writeline DateTime() & ": Processing MSMQ Message Count Alerts."
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

Const adCmdText = 1
Dim MSMQApp ' As MsmqApplication
Set MSMQApp = CreateObject("MSMQ.MSMQApplication")
Dim qFormat ' As String

Dim Mgmt ' As new MSMQManagement

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
	'wscript.echo DateTime() & ": Processing Service Alerts - EVDConnection1" & vbcrlf & EVDconnection1
	OutputLogFile.writeline DateTime() & ": Processing Service Alerts - EVDConnection1" & vbcrlf & EVDconnection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(EVDconnection1)

strSQLQuery1 = "Delete FROM dbo.MSMQMessageCountAlerts " & _
	"WHERE RecordCreateTimestamp > '" & styr & "/" & stmo & "/" & stdt & " 00:00:00';"

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": Processing Service Alerts - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": Processing Service Alerts - strSQLQuery1" & vbcrlf & strSQLQuery1
End if
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("ADODB.Recordset")
result1.open objCmd1

'**********************************************
'* Record new values into table for the current day.
For i = 0 to Ubound(EVServersList,2)

	MSMQApp.Machine = EVServersList(1,i)
	
	For each qFormat in MSMQApp.PrivateQueues
		
		Set Mgmt = CreateObject("MSMQ.MSMQManagement")
		Mgmt.Init MSMQApp.Machine,,"DIRECT=OS:" & qFormat
		
		Dim strMessageCount: strMessageCount = int(Mgmt.MessageCount)
		
		If strMessageCount > 0 then
			If EVDConnectionString = 1 Then
				EVDconnection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
					EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
			Elseif EVDConnectionString = 2 Then
				EVDconnection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
					EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
			End if
			
			'* VERBOSE LOGGING.
			If LoggingLevel >= 3 then
				'wscript.echo DateTime() & ": Processing Service Alerts - EVDconnection2" & vbcrlf & EVDconnection2
				OutputLogFile.writeline DateTime() & ": Processing Service Alerts - EVDconnection2" & vbcrlf & EVDconnection2
			End if
			
			'* Open database connection
			Set myconn2 = CreateObject("adodb.connection")
			myconn2.ConnectionTimeout = strSQLConnect
			myconn2.open(EVDconnection2)

			strSQLQuery2 = "INSERT INTO dbo.MSMQMessageCountAlerts " & _
				"(EVServer,MSMQPath,RecordCount,FileSize) " & _
				"VALUES ('" & UCase(EVServersList(1,i)) & "','" & qFormat & "','" & int(CLng(Mgmt.MessageCount)) & _
				"','" & int(CLng(Mgmt.BytesInQueue) / 2024) & "');"

			'* VERBOSE LOGGING.
			If LoggingLevel >= 3 then
				'wscript.echo DateTime() & ": Processing Service Alerts - strSQLQuery2" & vbcrlf & strSQLQuery2
				OutputLogFile.writeline DateTime() & ": Processing Service Alerts - strSQLQuery2" & vbcrlf & strSQLQuery2
			End if	
				
			Set objCmd2 = CreateObject("adodb.command")
			objCmd2.activeconnection = myconn2
			objCmd2.commandtimeout = strSQLExecute
			objCmd2.commandtype = adCmdText
			objCmd2.commandtext = strSQLQuery2

			Set result2 = CreateObject("adodb.Recordset")
			result2.open objCmd2
		End if
	Next
Next
Err.Clear
'********************************************************************************
'* MSMQ Messages Count Alerts FINISH
'********************************************************************************