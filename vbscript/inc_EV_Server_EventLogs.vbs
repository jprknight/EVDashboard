'********************************************************************************
'* EventLogs START
'********************************************************************************
ON ERROR RESUME NEXT

'* MINIMAL LOGGING.
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing EV Server Event Logs."
	If LoggingLevel >= 2 then
		OutputLogFile.writeline ""
		OutputLogFile.writeline "*************************************************************************"
		OutputLogFile.writeline ""
	End if
	OutputLogFile.writeline DateTime() & ": Processing EV Server EventLogs."
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

Dim objWMI, objEvent, strComputer, colLoggedEvents
Dim strTimeWritten, strTWYear, strTWMonth, strTWDay, strTWHour, strTWMin, strTWSec

'On Error Resume Next

For i = 0 to Ubound(EVServersList,2)
	
	'* MINIMAL LOGGING.
	If LoggingLevel >= 1 then
		wscript.echo DateTime() & ": (EV Server Event Logs) " & EVServersList(1,i)
		OutputLogFile.writeline DateTime() & ": (EV Server Event Logs) " & EVServersList(1,i)
	End if
	
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & EVServersList(1,i) & "\root\cimv2")
	Set colLoggedEvents = objWMI.ExecQuery ("Select * from Win32_NTLogEvent Where Logfile = 'Application' " & _
		"and Type <> 'Success'" )

	For Each objEvent in colLoggedEvents
		If objEvent.SourceName = "Enterprise Vault" Then
			
			strTimeWritten = objEvent.TimeWritten
			strTimeWritten = Left(strTimeWritten,14)
			strTWYear = Mid(strTimeWritten,1,4)
			strTWMonth = Mid(strTimeWritten,5,2)
			strTWDay = Mid(strTimeWritten,7,2)
			strTWHour = Mid(strTimeWritten,9,2)
			strTWMin = Mid(strTimeWritten,11,2)
			strTWSec = Mid(strTimeWritten,13,2)
			strTimeWritten = strTWYear & "/" & strTWMonth & "/" & strTWDay & " " & strTWHour & ":" & strTWMin & ":" & strTWSec
			
			'* Replace ' characters with " to prevent unexpected database field delimiters.
			strMessage = objEvent.Message
			strMessage = replace(strMessage,chr(39),chr(34))
			
			If DateDiff("d",strTimeWritten,now) < 1 then
				'wscript.echo ("Category: " & objEvent.Category)
				'wscript.echo ("Computer Name: " & objEvent.ComputerName)
				'wscript.echo ("Event Code: " & objEvent.EventCode)
				'wscript.echo ("Time Written: " & strTimeWritten)
				'wscript.echo ("Event Type: " & objEvent.Type)
				'wscript.echo ("Message: " & objEvent.Message)
				'wscript.echo ("Record Number: " & objEvent.RecordNumber)
				'wscript.echo ("Source Name: " & objEvent.SourceName)
				'wscript.echo ("User: " & objEvent.User)
				'wscript.echo (" ")
				
				'* NORMAL LOGGING.
				If LoggingLevel >= 2 then
					'wscript.echo DateTime() & ": (EV Server Event Logs) : Event Code - " & objEvent.EventCode  & " - Event Type: " & objEvent.Type
					OutputLogFile.writeline DateTime() & ": (EV Server Event Logs) : Event Code - " & objEvent.EventCode  & " - Event Type: " & objEvent.Type
				End if
				
				If EVDConnectionString = 1 Then
					Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
						EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
				Elseif EVDConnectionString = 2 Then
					Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
						EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
				End if
				
				'* Open database connection
				Set myconn1 = CreateObject("adodb.connection")
				myconn1.ConnectionTimeout = strSQLConnect
				myconn1.open(Connection1)

				strSQLQuery1 = "UPDATE dbo.EVServerEventLogs " & _
					"SET EVServer='" & objEvent.ComputerName & "'," & _
					"EventCode='" & objEvent.EventCode & "'," & _
					"TimeWritten='" & strTimeWritten & "'," & _
					"EventType='" & objEvent.Type & "'," & _
					"Message='" & strMessage & "' " & _
					"WHERE EVServer='" & objEvent.ComputerName & "' AND " & _
					"EventCode='" & objEvent.EventCode & "' AND " & _
					"TimeWritten='" & strTimeWritten & "' AND " & _
					"RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
					"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' " & _
					"IF @@ROWCOUNT=0 " & _
					"INSERT INTO dbo.EVServerEventLogs " & _
					"(EVServer,EventCode,TimeWritten,EventType,Message) " & _
					"VALUES ('" & objEvent.ComputerName & "','" & objEvent.EventCode & "','" & strTimeWritten & _
					"','" & objEvent.Type & "','" & strMessage & "');"

				'* VERBOSE LOGGING.
				If LoggingLevel >= 3 then
					'wscript.echo DateTime() & ": (EV Server Event Logs) - strSQLQuery1" & vbcrlf & strSQLQuery1
					OutputLogFile.writeline DateTime() & ": (EV Server Event Logs) - strSQLQuery1" & vbcrlf & strSQLQuery1
				End if
					
				Set objCmd1 = CreateObject("adodb.command")
				objCmd1.activeconnection = myconn1
				objCmd1.commandtimeout = strSQLExecute
				objCmd1.commandtype = adCmdText
				objCmd1.commandtext = strSQLQuery1

				Set result1 = CreateObject("adodb.Recordset")
				result1.open objCmd1
			End if
		End if
	Next
Next
'********************************************************************************
'* EventLogs END
'********************************************************************************
