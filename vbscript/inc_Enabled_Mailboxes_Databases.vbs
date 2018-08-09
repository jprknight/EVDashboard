'********************************************************************************
'* ENABLED MAILBOXES DATABASES
'********************************************************************************
'*
'* This script does three things:
'*	1. Connects to the EnabledMailboxes database and retrives data already populated for the current day.
'*	2. Connects to the ArchivePoint table in the correct database for the current archive.
'*	3. Writes the data retrieved from the archivepoint table back into the EnabledMailboxes table.
'*
'********************************************************************************

'* Added ON ERROR RESUME NEXT to allow for network disconnections from volume data retrievals.
'ON ERROR RESUME NEXT

If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Enabled Mailboxes Databases."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Enabled Mailboxes Databases."
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

'* DATA RETRIEVE SQL QUERY 1.

If EVDConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (Enabled Mailboxes Databases) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) - EVDConnectionString : " & EVDConnectionString
	OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "Select Mailbox AS 'Mailbox'," & _
	"ExchangeServer AS 'ExchangeServer'," & _
	"NumItemsMailbox AS 'NumItemsMailbox'," & _
	"MailboxSizeMB AS 'MailboxSizeMB'," & _
	"ExchangeState AS 'ExchangeState'," & _
	"DefaultVaultID AS 'DefaultVaultID'," & _
	"DBServer AS 'DBServer'," & _
	"EVDatabase AS 'EVDatabase' " & _
	"FROM dbo.EnabledMailboxes " & _
	"WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": (Enabled Mailboxes Databases) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) - strSQLQuery1" & vbcrlf & strSQLQuery1
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

	'* The following need to be in numerical order, so all variables are set correctly.
	strMailbox = result1.Fields.Item(0)
	strMailboxSizeMB = result1.Fields.Item(3)
	strDefaultVaultID = result1.Fields.Item(5)
	strDBServer = result1.Fields.Item(6)
	strEVDatabase = result1.Fields.Item(7)
	
	'* If the row has a blank DefaultVaultID skip it.
	If Len(strDefaultVaultID) > 1 then
		
		Result1Count = Result1Count + 1
		
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Enabled Mailboxes Databases) " & Result1Count & " " & strMailbox
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) " & Result1Count & " " & strMailbox
		End if
		
		'* DATA RETRIEVE SQL QUERY 2.
		
		If ConnectionString = 1 Then
			Connection2 = "Driver={SQL Server};Server=" & strDBServer & ";Database=" & _
				strEVDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLAltPortNumber & ";"
		Elseif ConnectionString = 2 Then
			Connection2 = "Driver={SQL Server};Server=" & strDBServer & ";Database=" & _
				strEVDatabase & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
		End if
		
		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Enabled Mailboxes Databases) - Connection2" & vbcrlf & Connection2
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) - ConnectionString : " & ConnectionString
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) - Connection2" & vbcrlf & Connection2
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)

		strSQLQuery2 = "Select VS1.ArchivedItems AS 'NumItemsArchive'," & _ 
			"VS1.ArchivedItemsSize/1024 AS 'ArchiveSizeMB'," & _
			"VS1.CreatedDate AS 'ArchiveCreated'," & _
			"VS1.ModifiedDate AS 'ArchiveUpdated' " & _
			"FROM " & strEVDatabase & "..ArchivePoint AS VS1 " & _
			'"WHERE VS1.ArchivePointID like '%" & strDefaultVaultID & "%'"
			"WHERE VS1.ArchivePointID = '" & strDefaultVaultID & "'"
		
		'* VERBOSE Logging.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Enabled Mailboxes Databases - strSQLQuery2) " & vbcrlf & strSQLQuery2
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases - strSQLQuery2) " & strSQLQuery2
		End if
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2

		Set result2 = CreateObject("ADODB.Recordset")
		result2.open objCmd2
		
		While not result2.EOF
				
				'************************************************
				'* UPDATE SQL QUERY START.
				Dim strArchiveSizeMB: strArchiveSizeMB = (result2("ArchiveSizeMB"))
				
				If EVDConnectionString = 1 Then
					Connection3 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
						EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
				Elseif EVDConnectionString = 2 Then
					Connection3 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
						EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
				End if

				'* NORMAL LOGGING.
				If LoggingLevel >= 2 then
					'wscript.echo DateTime() & ": (Enabled Mailboxes Databases) - Connection3" & vbcrlf & Connection3
					OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) - EVDConnectionString : " & EVDConnectionString
					OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) - Connection3" & vbcrlf & Connection3
				End if
				
				'* Open database connection
				Set myconn3 = CreateObject("adodb.connection")
				myconn3.ConnectionTimeout = strSQLConnect
				myconn3.open(Connection3)

				Dim strTotalSizeMB: strTotalSizeMB = 0
				strTotalSizeMB = strTotalSizeMB + Cint(strMailboxSizeMB)
				strTotalSizeMB = strTotalSizeMB + Cint(strArchiveSizeMB)
				
				strSQLQuery3 = "UPDATE dbo.EnabledMailboxes " & _
					"SET NumItemsArchive='" & (result2("NumItemsArchive")) & "'," & _
					"ArchiveSizeMB='" & (result2("ArchiveSizeMB")) & "'," & _
					"TotalSizeMB='" & strTotalSizeMB & "'," & _
					"ArchiveCreated='" & dbDate((result2("ArchiveCreated"))) & "'," & _
					"ArchiveUpdated='" & dbDate((result2("ArchiveUpdated"))) & "' " & _
					"WHERE DefaultVaultID = '" & strDefaultVaultID & "' AND " & _
					"RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
					"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"

				'"WHERE DefaultVaultID like '%" & strDefaultVaultID & "%' AND " & _
				
				'* VERBOSE Logging.
				If LoggingLevel >= 3 then
					'wscript.echo DateTime() & ": (Enabled Mailboxes Databases - strSQLQuery3) " & vbcrlf & strSQLQuery3
					OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases - strSQLQuery3) " & strSQLQuery3
				End if	
					
				Set objCmd3 = CreateObject("adodb.command")
				objCmd3.activeconnection = myconn3
				objCmd3.commandtimeout = strSQLExecute
				objCmd3.commandtype = adCmdText
				objCmd3.commandtext = strSQLQuery3
				
				Set result3 = CreateObject("adodb.Recordset")
				result3.open objCmd3
				'* UPDATE SQL QUERY FINISH.
				'************************************************
		result2.movenext()
		Wend
	Else
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Enabled Mailboxes Databases) " & Result1Count & " " & strMailbox
			OutputLogFile.writeline DateTime() & ": (Enabled Mailboxes Databases) " & Result1Count & " " & _
				strMailbox & " NULL DefaultVaultID found, skipping record."
		End if
	End if
	
result1.movenext()

Wend

'********************************************************************************
'* ENABLED MAILBOXES DATABASES END
'********************************************************************************