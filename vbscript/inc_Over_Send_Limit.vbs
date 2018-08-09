'********************************************************************************
'* OVER SEND LIMIT
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Over Send Limit."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Over Send Limit."
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

If ConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
Elseif ConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (Over Send Limit) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (Over Send Limit) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

'* Bug #3180069 EV 9.0 Schema update fix.
If EVVersion = 1 then
	'* Option for 'Pre EV 9.0'.
	strSQLQuery1 = "Use EnterpriseVaultDirectory " & _
		"SELECT MbxDisplayName as 'Mailbox', " & _
		"MBxSize/1000 - MbxSendLimit/1000 as 'Over Send (MB)', " & _
		"MbxWarningLimit/1000 as 'Warning (MB)', " & _
		"MbxSendLimit/1000 as 'Send (MB)', " & _
		"MbxReceiveLimit/1000 as 'Receive (MB)', " & _
		"MbxSize/1000 as 'Mbx Size (MB)', " & _
		"MbxItemCount as '#Items (Mailbox)', " & _
		"ExchangeComputer as 'Exchange Server', " & _
		"YoungestArchivedDateUTC as 'Last Archived', " & _
		"MbxExchangeState as 'Exchange State', " & _
		"PE.poName as 'Policy Used' " & _
		"FROM " & _
		"EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry as EME, " & _
		"EnterpriseVaultDirectory.dbo.ExchangeServerEntry as ESE, " & _
		"EnterpriseVaultDirectory.dbo.Root as RT, " & _
		"EnterpriseVaultDirectory.dbo.IndexVolume as IV, " & _
		"EnterpriseVaultDirectory.dbo.PolicyEntry as PE " & _
		"WHERE " & _
		"EME.MBxSendLimit > 0 AND " & _
		"EME.MbxArchivingState = 1 AND " & _
		"EME.MBxSize > EME.MbxSendLimit and " & _
		"EME.ExchangeServerIdentity = ESE.ExchangeServerIdentity AND " & _
		"EME.PolicyEntryId = poPolicyEntryID AND " & _
		"EME.DefaultVaultID = RT.VaultEntryID AND " & _
		"RT.rootidentity = IV.RootIdentity " & _
		"ORDER by Mailbox asc"
Elseif EVVersion = 2 then
	'* Option for 'EV 9.0 and newer'.
	strSQLQuery1 = "Use EnterpriseVaultDirectory " & _
		"SELECT MbxDisplayName as 'Mailbox', " & _
		"MBxSize/1000 - MbxSendLimit/1000 as 'Over Send (MB)', " & _
		"MbxWarningLimit/1000 as 'Warning (MB)', " & _
		"MbxSendLimit/1000 as 'Send (MB)', " & _
		"MbxReceiveLimit/1000 as 'Receive (MB)', " & _
		"MbxSize/1000 as 'Mbx Size (MB)', " & _
		"MbxItemCount as '#Items (Mailbox)', " & _
		"ExchangeComputer as 'Exchange Server', " & _
		"YoungestArchivedDateUTC as 'Last Archived', " & _
		"MbxExchangeState as 'Exchange State', " & _
		"PE.poName as 'Policy Used' " & _
		"FROM " & _
		"EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry as EME, " & _
		"EnterpriseVaultDirectory.dbo.ExchangeMailboxStore as EMS, " & _
		"EnterpriseVaultDirectory.dbo.ExchangeServerEntry as ESE, " & _
		"EnterpriseVaultDirectory.dbo.Root as RT, " & _
		"EnterpriseVaultDirectory.dbo.IndexVolume as IV, " & _
		"EnterpriseVaultDirectory.dbo.PolicyEntry as PE " & _
		"WHERE " & _
		"EME.MBxSendLimit > 0 AND " & _
		"EME.MbxArchivingState = 1 AND " & _
		"EME.MBxSize > EME.MbxSendLimit and " & _
		"EME.MbxStoreIdentity = EMS.MbxStoreIdentity and " & _
		"EMS.ExchangeServerIdentity = ESE.ExchangeServerIdentity AND " & _
		"EME.PolicyEntryId = poPolicyEntryID AND " & _
		"EME.DefaultVaultID = RT.VaultEntryID AND " & _
		"RT.rootidentity = IV.RootIdentity " & _
		"ORDER by Mailbox asc"
End if

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": (Over Send Limit) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Over Send Limit) - strSQLQuery1" & vbcrlf & strSQLQuery1
End if	
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

Dim RecordCount: RecordCount = 0

While not result1.EOF

	'****************************************************
	'* Manipulate data to be inserted into EVDashboard database.
		' apostrophe fix start.
		strMailbox = (result1("Mailbox"))
		If InStr(strMailbox,chr(39)) then
			strMailbox = replace(strMailbox,chr(39)," ")
			'* VERBOSE LOGGING.
			If LoggingLevel >= 3 then
				'wscript.echo DateTime() & ": (Over Receive Limit) - Apostrophe Fix" & vbcrlf & strMailbox
				OutputLogFile.writeline DateTime() & ": (Over Receive Limit) - Apostrophe Fix" & vbcrlf & strMailbox
			End if
		end if
		' apostrophe fix end.

		Select Case (result1("Exchange State"))
			Case 0
				ExchangeState = "Normal"
			Case 1
				ExchangeState = "Hidden"
			Case 2
				ExchangeState = "Deleted"
		End Select

		'* VERBOSE LOGGING.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Over Receive Limit) - ExchangeState - " & ExchangeState
			OutputLogFile.writeline DateTime() & ": (Over Receive Limit) - ExchangeState - " & ExchangeState
		End if
		
	RecordCount = RecordCount + 1
	'* MINIMAL LOGGING.
	If LoggingLevel >= 1 then
		wscript.echo DateTime() & ": (Over Receive Limit) - " & RecordCount & " - " & strMailbox
		OutputLogFile.writeline DateTime() & ": (Over Receive Limit) - " & RecordCount & " - " & strMailbox
	End if
	
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
			'wscript.echo DateTime() & ": (Over Receive Limit) - Connection2" & vbcrlf & Connection2
			OutputLogFile.writeline DateTime() & ": (Over Receive Limit) - Connection2" & vbcrlf & Connection2
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)

		strSQLQuery2 = "UPDATE dbo.OverSendLimit " & _
			"SET Mailbox='" & strMailbox & "'," & _
			"OverSendMB='" & (result1("Over Send (MB)")) & "'," & _
			"WarningMB='" & (result1("Warning (MB)")) & "'," & _
			"SendMB='" & (result1("Send (MB)")) & "'," & _
			"ReceiveMB='" & (result1("Receive (MB)")) & "'," & _
			"MailboxSizeMB='" & (result1("Mbx Size (MB)")) & "'," & _
			"NumItemsMailbox='" & (result1("#Items (Mailbox)")) & "'," & _
			"ExchangeServer='" & (result1("Exchange Server")) & "'," & _
			"LastArchived='" & dbDate((result1("Last Archived"))) & "'," & _
			"ExchangeState='" & ExchangeState & "'," & _
			"PolicyUsed='" & (result1("Policy Used")) & "' " & _
			"WHERE Mailbox='" & strMailbox & "' AND " & _
			"RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
			"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'" & _
			"IF @@ROWCOUNT=0 " & _
			"INSERT INTO dbo.OverSendLimit " & _
			"(Mailbox,OverSendMB,WarningMB,SendMB,ReceiveMB,MailboxSizeMB,NumItemsMailbox," & _
			"ExchangeServer,LastArchived,ExchangeState,PolicyUsed) " & _
			"VALUES ('" & strMailbox & "'," & (result1("Over Send (MB)")) & "," & (result1("Warning (MB)")) & _
			"," & (result1("Send (MB)")) & "," & (result1("Receive (MB)")) & "," & _
			(result1("Mbx Size (MB)")) & "," & (result1("#Items (Mailbox)")) & _
			",'" & (result1("Exchange Server")) & "','" & dbDate((result1("Last Archived"))) & _
			"','" & ExchangeState & "','" & (result1("Policy Used")) & "');"
		
		'* VERBOSE LOGGING.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Over Receive Limit) - strSQLQuery2" & vbcrlf & strSQLQuery2
			OutputLogFile.writeline DateTime() & ": (Over Receive Limit) - strSQLQuery2" & vbcrlf & strSQLQuery2
		End if
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2

		Set result2 = CreateObject("ADODB.Recordset")
		result2.open objCmd2
	
	result1.movenext()
Wend

'********************************************************************************
'* OVER SEND LIMIT END
'********************************************************************************
