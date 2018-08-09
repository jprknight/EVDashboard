'********************************************************************************
'* ENABLED MAILBOXES
'********************************************************************************
'*
'* This script loads in the data into the first columns of the enabledmailboxes table.

'* Added ON ERROR RESUME NEXT to allow for network disconnections from volume data retrievals.
'ON ERROR RESUME NEXT

'* MINIMAL LOGGING.
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Enabled Mailboxes."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Enabled Mailboxes."
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
	'wscript.echo DateTime() & ": (Enabled Mailboxes) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes) - ConnectionString" & vbcrlf & ConnectionString
	OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "SELECT MbxDisplayName AS 'Mailbox'," & _
	"ExchangeComputer AS 'ExchangeServer'," & _
	"MbxItemCount AS 'NumItemsMailbox'," & _
	"MbxSize/1024 AS 'MailboxSizeMB'," & _
	"MbxExchangeState AS 'ExchangeState'," & _
	"DefaultVaultID AS 'DefaultVaultID'," & _
	"ComputerNameAlternate AS 'EVServer'," & _
	"SQLServer AS 'DBServer'," & _
	"DatabaseDSN AS 'EVDatabase' " & _
	"FROM " & _
	"EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry AS EME," & _
	"EnterpriseVaultDirectory.dbo.ExchangeServerEntry AS ESE, " & _
	"EnterpriseVaultDirectory.dbo.Archive AS EA, " & _
	"EnterpriseVaultDirectory.dbo.VaultStoreEntry AS EVSE," & _
	"EnterpriseVaultDirectory.dbo.ComputerEntry AS ECE " & _
	"WHERE " & _
	"EME.MbxArchivingState = 1 AND " & _
	"EME.ExchangeServerIdentity = ESE.ExchangeServerIdentity AND " & _
	"EME.MbxDisplayname = EA.ArchiveName AND " & _
	"EA.VaultStoreEntryId = EVSE.VaultStoreEntryId AND " & _
	"EVSE.VaultStoreEntryId = ECE.DefVaultStoreEntryID " & _
	"ORDER BY Mailbox"

'* VERBOSE Logging.
If LoggingLevel >= 3 then
	wscript.echo DateTime() & ": (Enabled Mailboxes - strSQLQuery1) " & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes - strSQLQuery1) " & vbcrlf & strSQLQuery1
End if
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

Dim Result1Count: Result1Count = 0

While not result1.EOF

	'****************************************************
	'* Manipulate data to be inserted into EVDashboard database.
		' apostrophe fix start.
		strMailbox = (result1("Mailbox"))
		If InStr(strMailbox,chr(39)) then
			strMailbox = replace(strMailbox,chr(39)," ")
			'* VERBOSE LOGGING.
			If LoggingLevel >= 3 then
				'wscript.echo DateTime() & ": (Enabled Users) - Apostrophe Fix" & vbcrlf & strMailbox
				OutputLogFile.writeline DateTime() & ": (Enabled Users) - Apostrophe Fix : " & strMailbox
			End if
		end if
		' apostrophe fix end.

	Result1Count = Result1Count + 1
	
	If LoggingLevel >= 2 then
		wscript.echo DateTime() & ": (Enabled Mailboxes) " & Result1Count & " " & strMailbox
		OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes) " & Result1Count & " " & strMailbox
	End if
		
	'****************************************************
	'* Connect to EVDashboard database and update or insert data.
		If EVDConnectionString = 1 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
		Elseif EVDConnectionString = 2 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
		End if
		
		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Enabled Mailboxes) - Connection2" & vbcrlf & Connection2
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes) - EVDConnectionString : " & EVDConnectionString
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes) - Connection2" & vbcrlf & Connection2
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)

		strSQLQuery2 = "UPDATE dbo.EnabledMailboxes " & _
			"SET Mailbox='" & strMailbox & "'," & _
			"ExchangeServer='" & UCase((result1("ExchangeServer"))) & "'," & _
			"NumItemsMailbox='" & (result1("NumItemsMailbox")) & "'," & _
			"MailboxSizeMB='" & (result1("MailboxSizeMB")) & "'," & _
			"ExchangeState='" & (result1("ExchangeState")) & "'," & _
			"DefaultVaultID='" & (result1("DefaultVaultID")) & "'," & _
			"EVServer='" & UCase((result1("EVServer"))) & "'," & _
			"DBServer='" & UCase((result1("DBServer"))) & "'," & _
			"EVDatabase='" & (result1("EVDatabase")) & "' " & _
			"WHERE Mailbox='" & strMailbox & "' AND " & _
			"RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
			"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' " & _
			"IF @@ROWCOUNT=0 " & _					
			"INSERT INTO dbo.EnabledMailboxes " & _
			"(Mailbox,ExchangeServer,NumItemsMailbox,MailboxSizeMB,ExchangeState,DefaultVaultID,EVServer,DBServer,EVDatabase) " & _
			"VALUES ('" & strMailbox & "','" & _
			UCase((result1("ExchangeServer"))) & "'," & _
			(result1("NumItemsMailbox")) & "," & _
			(result1("MailboxSizeMB")) & "," & _
			(result1("ExchangeState")) & ",'" & _
			(result1("DefaultVaultID")) & "','" & _
			UCase((result1("EVServer"))) & "','" & _
			UCase((result1("DBServer"))) & "','" & _
			(result1("EVDatabase")) & "');"
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2
		
		'* VERBOSE Logging.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Enabled Mailboxes - strSQLQuery2) " & vbcrlf & strSQLQuery2
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes - strSQLQuery2) " & strSQLQuery2
		End if
		
		Set result2 = CreateObject("ADODB.Recordset")
		result2.open objCmd2
	'* Move Result1 recordset to next record.
	result1.movenext()
Wend
'********************************************************************************
'* ENABLED MAILBOXES END
'********************************************************************************