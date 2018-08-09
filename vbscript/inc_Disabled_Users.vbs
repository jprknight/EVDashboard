'********************************************************************************
'* DISABLED USERS
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Disabled Users."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Disabled Users."
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

If ConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
Elseif ConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (Disabled Users) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (Disabled Users) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "USE EnterpriseVaultDirectory " & _
	"SELECT MbxDisplayName as 'Mailbox', " & _
	"MbxWarningLimit/1000 as 'Warning (MB)', " & _
	"MbxSendLimit/1000 as 'Send (MB)', " & _
	"MbxReceiveLimit/1000 as 'Receive (MB)', " & _
	"MbxSize/1024 as 'Mbx Size (MB)', " & _
	"MbxItemCount as '#Items (Mailbox)' " & _
	"FROM " & _
	"EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry as EME " & _
	"WHERE " & _
	"EME.mbxarchivingstate = 2 " & _
	"ORDER by Mailbox"

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": (Disabled Users) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Disabled Users) - strSQLQuery1" & vbcrlf & strSQLQuery1
End if
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1
	
While not result1.EOF

	'****************************************************
	'* Manipulate data to be inserted into EVDashboard database.
		' apostrophe fix start.
		strMailbox = (result1("Mailbox"))
		If InStr(strMailbox,chr(39)) then
			strMailbox = replace(strMailbox,chr(39)," ")
			'* VERBOSE LOGGING.
			If LoggingLevel >= 3 then
				'wscript.echo DateTime() & ": (Disabled Users) - Apostrophe Fix" & vbcrlf & strMailbox
				OutputLogFile.writeline DateTime() & ": (Disabled Users) - Apostrophe Fix" & vbcrlf & strMailbox
			End if
		end if
		' apostrophe fix end.

	'****************************************************
	'* Connect to EVDashboard database and insert data.
		If EVDConnectionString = 1 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
		Elseif EVDConnectionString = 2 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
		End if
		
		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Disabled Users) - Connection2" & vbcrlf & Connection2
			OutputLogFile.writeline DateTime() & ": (Disabled Users) - Connection2" & vbcrlf & Connection2
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)

		strSQLQuery2 = "UPDATE dbo.DisabledUsers " & _
		"SET mailbox='" & strMailbox & "'," & _
		"WarningMB='" & (result1("Warning (MB)")) & "'," & _
		"SendMB='" & (result1("Send (MB)")) & "'," & _
		"ReceiveMB='" & (result1("Receive (MB)")) & "'," & _
		"MailboxSizeMB='" & (result1("Mbx Size (MB)")) & "'," & _
		"NumItemsMailbox='" & (result1("#Items (Mailbox)")) & "' " & _
		"WHERE mailbox='" & strMailbox & "' AND " & _
		"RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'" & _
		"IF @@ROWCOUNT=0 " & _
		"INSERT INTO dbo.DisabledUsers " & _
		"(Mailbox,WarningMB,SendMB,ReceiveMB,MailboxSizeMB,NumItemsMailbox) " & _
		"VALUES ('" & strMailbox & "'," & _
		(result1("Warning (MB)")) & "," & _
		(result1("Send (MB)")) & "," & _
		(result1("Receive (MB)")) & "," & _
		(result1("Mbx Size (MB)")) & "," & _
		(result1("#Items (Mailbox)")) & ");"
		
		'* VERBOSE LOGGING.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Disabled Users) - strSQLQuery2" & vbcrlf & strSQLQuery2
			OutputLogFile.writeline DateTime() & ": (Disabled Users) - strSQLQuery2" & vbcrlf & strSQLQuery2
		End if
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2

		Set result2 = CreateObject("ADODB.Recordset")
		result2.open objCmd2
	'* Move Result1 recordset to next record.
	result1.movenext()
Wend
'********************************************************************************
'* DISABLED USERS END
'********************************************************************************