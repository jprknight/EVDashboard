'********************************************************************************
'* USAGE REPORT BY ARCHIVE
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Usage Report By Archive."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Usage Report By Archive."
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

'***************************************************
'* 1 - Connect to the Directory database server and get all mailboxes which are enabled for archiving.

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
	'wscript.echo DateTime() & ": (Usage Report By Archive) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (Usage Report By Archive) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "SELECT mbxDisplayName AS Mailbox," & _
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

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": (Usage Report By Archive) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Usage Report By Archive) - strSQLQuery1" & vbcrlf & strSQLQuery1
End if		
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

Dim RecordCount : RecordCount = 0

While not result1.EOF
	
	'***************************************************
	'* 2 - Connect to each archive's mailbox archive database and retrieve information from the saveset table.
	
	'***************************************************
	'* Manipulate data to be inserted into EVDashboard database.
		' apostrophe fix start.
		strMailbox = (result1("Mailbox"))
		If InStr(strMailbox,chr(39)) then
			strMailbox = replace(strMailbox,chr(39)," ")
			'* VERBOSE LOGGING.
			If LoggingLevel >= 3 then
				'wscript.echo DateTime() & ": (Usage Report By Archive) - Apostrophe Fix" & vbcrlf & strMailbox
				OutputLogFile.writeline DateTime() & ": (Usage Report By Archive) - Apostrophe Fix" & vbcrlf & strMailbox
			End if
		end if
		' apostrophe fix end.

	RecordCount = RecordCount + 1
	'* NORMAL LOGGING.
	If LoggingLevel >= 2 then
		wscript.echo DateTime() & ": Processing: " & RecordCount & " : " & strMailbox
		OutputLogFile.writeline DateTime() & ": Processing: " & RecordCount & " : " & strMailbox
	End if
	
	'***************************************************
	'* Connect to EVDashboard database and insert data.
		If ConnectionString = 1 Then
			Connection2 = "Driver={SQL Server};Server=" & (result1("DBServer")) & _
				";Database=" & (result1("EVDatabase")) & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
		Elseif ConnectionString = 2 Then
			Connection2 = "Driver={SQL Server};Server=" & (result1("DBServer")) & _
				";Database=" & (result1("EVDatabase")) & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
		End if
		
		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Usage Report By Archive) - Connection2" & vbcrlf & Connection2
			OutputLogFile.writeline DateTime() & ": (Usage Report By Archive) - Connection2" & vbcrlf & Connection2
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)
		
		strSQLQuery2 = "SELECT LEFT(convert(varchar, S.ArchivedDate,120),4) AS 'ArchivedDate', " & _
			"COUNT(S.ArchivedDate) AS 'ArchivedAmount' " & _
			"FROM SaveSet S, ArchivePoint AP " & _
			"WHERE S.ArchivePointIdentity = AP.ArchivePointIdentity " & _
			"AND AP.ArchivePointId = '" & (result1("DefaultVaultID")) & "' " & _
			"GROUP BY LEFT(CONVERT(varchar, S.ArchivedDate,120), 4) " '& _
			'"ORDER BY ArchivedDate ASC"
		
		'* VERBOSE LOGGING.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Usage Report By Archive) - strSQLQuery2" & vbcrlf & strSQLQuery2
			OutputLogFile.writeline DateTime() & ": (Usage Report By Archive) - strSQLQuery2" & vbcrlf & strSQLQuery2
		End if
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2

		Set result2 = CreateObject("ADODB.Recordset")
		result2.open objCmd2
		
		While not result2.EOF
		
			'***************************************************
			'* 3 - Write the information into the UsageReportByArchive table in the EV Dashboard database.
		
				If EVDConnectionString = 1 Then
					Connection3 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
						EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
				Elseif EVDConnectionString = 2 Then
					Connection3 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
						EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
				End if
				
				'* NORMAL LOGGING.
				If LoggingLevel >= 2 then
					'wscript.echo DateTime() & ": (Usage Report By Archive) - Connection3" & vbcrlf & Connection3
					OutputLogFile.writeline DateTime() & ": (Usage Report By Archive) - Connection3" & vbcrlf & Connection3
				End if
				
				'* Open database connection
				Set myconn3 = CreateObject("adodb.connection")
				myconn3.ConnectionTimeout = strSQLConnect
				myconn3.open(Connection3)

				strSQLQuery3 = 	"UPDATE dbo.UsageReportByArchive " & _
					"SET mailbox='" & strMailbox & "'," & _
					"EMEDefaultVaultID='" & (result1("DefaultVaultID")) & "', " & _
					"ArchivePeriod='" & (result2("ArchivedDate")) & "', " & _ 
					"ItemCount='" & (result2("ArchivedAmount")) & "' " & _
					"WHERE mailbox='" & strMailbox & "' AND " & _
					"ArchivePeriod='" & (result2("ArchivedDate")) & "' " & _
					"IF @@ROWCOUNT=0 " & _
					"INSERT INTO dbo.UsageReportByArchive " & _
					"(Mailbox,EMEDefaultVaultID,ArchivePeriod,ItemCount) " & _
					"VALUES ('" & strMailbox & "','" & _
					(result1("DefaultVaultID")) & "','" & _
					(result2("ArchivedDate")) & "','" & _
					(result2("ArchivedAmount")) & "');"
				
				'* VERBOSE LOGGING.
				If LoggingLevel >= 3 then
					'wscript.echo DateTime() & ": (Usage Report By Archive) - strSQLQuery3" & vbcrlf & strSQLQuery3
					OutputLogFile.writeline DateTime() & ": (Usage Report By Archive) - strSQLQuery3" & vbcrlf & strSQLQuery3
				End if
				
				Set objCmd3 = CreateObject("adodb.command")
				objCmd3.activeconnection = myconn3
				objCmd3.commandtimeout = strSQLExecute
				objCmd3.commandtype = adCmdText
				objCmd3.commandtext = strSQLQuery3

				Set result3 = CreateObject("ADODB.Recordset")
				result3.open objCmd3
			
			result2.movenext()
			
		Wend
	
	result1.movenext()
Wend
'********************************************************************************
'* DISABLED USERS END
'********************************************************************************