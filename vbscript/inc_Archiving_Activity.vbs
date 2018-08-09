'********************************************************************************
'* ARCHIVING ACTIVITY START
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Archiving Activity."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Archiving Activity."
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

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

If LoggingLevel >= 2 then
	wscript.echo DateTime() & ": Processing Archiving Activity - strSQLConnect: " & strSQLConnect
	OutputLogFile.writeline DateTime() & ": Processing Archiving Activity - strSQLExecute: " & strSQLExecute
	
	wscript.echo DateTime() & ": Processing Archiving Activity - Connection String 1:" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": Processing Archiving Activity - Connection String 1:" & vbcrlf & Connection1
End if

strSQLQuery1 = "SELECT ArchiveName AS 'Mailbox', " & _
		"ComputerEntry.ComputerNameAlternate AS 'EVServer'," & _
		"VaultStoreEntry.SQLServer AS 'DBServer'," & _
		"VaultStoreEntry.DatabaseDSN AS 'EVDatabase', " & _
		"ExchangeServerEntry.ExchangeComputer AS 'ExchangeComputer', " & _
		"ExchangeServerEntry.ExchangeServerIdentity AS 'ExchangeServerIdentity', " & _
		"root.rootidentity AS 'root_rootidentity' " & _
		"FROM enterprisevaultdirectory.dbo.root root " & _
		"INNER JOIN enterprisevaultdirectory.dbo.archive archive " & _
		"ON root.rootidentity = archive.rootidentity " & _
		"INNER JOIN enterprisevaultdirectory.dbo.ExchangeMailboxEntry ExchangeMailboxEntry " & _
		"ON ExchangeMailboxEntry.DefaultVaultId = root.VaultEntryId " & _
		"INNER JOIN enterprisevaultdirectory.dbo.ExchangeServerEntry ExchangeServerEntry " & _
		"ON ExchangeServerEntry.ExchangeServerIdentity = ExchangeMailboxEntry.ExchangeServerIdentity " & _
		"INNER JOIN EnterpriseVaultDirectory.dbo.VaultStoreEntry VaultStoreEntry " & _
		"ON archive.VaultStoreEntryId = VaultStoreEntry.VaultStoreEntryId " & _
		"INNER JOIN EnterpriseVaultDirectory.dbo.ComputerEntry ComputerEntry " & _
		"ON VaultStoreEntry.VaultStoreEntryId = ComputerEntry.DefVaultStoreEntryID " & _
		"GROUP BY archivename, ExchangeServerEntry.ExchangeComputer," & _
		"ExchangeServerEntry.ExchangeServerIdentity,ComputerEntry.ComputerNameAlternate," & _
		"VaultStoreEntry.SQLServer, VaultStoreEntry.DatabaseDSN,root.rootidentity " & _
		"ORDER BY archivename, ExchangeServerEntry.ExchangeComputer,root_rootidentity"

If LoggingLevel >= 3 then
	wscript.echo DateTime() & ": (Processing Archiving Activity - strSQLQuery1)" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (Processing Archiving Activity - strSQLQuery1)" & vbcrlf & strSQLQuery1
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
		end if
		' apostrophe fix end.

	'****************************************************
	'* Connect to mailbox archiving database and retrieve NumItemsArchived
	
		If EVDConnectionString = 1 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
		Elseif EVDConnectionString = 2 Then
			Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
		End if

		If LoggingLevel >= 2 then			
			wscript.echo DateTime() & ": Processing Archiving Activity - Connection String 2:" & vbcrlf & Connection2
			OutputLogFile.writeline DateTime() & ": Processing Archiving Activity - Connection String 2:" & vbcrlf & Connection2
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)
		
		strSQLQuery2 = "INSERT INTO dbo.ArchivingActivity " & _
				"(Mailbox,ExchangeServer,EVServer,DBServer,EVDatabase,root_rootidentity,ExchangeServerIdentity) " & _
				"VALUES ('" & strMailbox & "','" & (result1("ExchangeComputer")) & "','" & (result1("EVServer")) & "','" & _
				(result1("DBServer")) & "','" & (result1("EVDatabase")) & "','" & (result1("root_rootidentity")) & "','" & _
				(result1("ExchangeServerIdentity")) & "');"
				
		If LoggingLevel >= 3 then
			wscript.echo DateTime() & ": (Processing Archiving Activity - strSQLQuery2)" & vbcrlf & strSQLQuery2
			OutputLogFile.writeline DateTime() & ": (Processing Archiving Activity- strSQLQuery2)" & vbcrlf & strSQLQuery2
		End if
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2
			
		Set result2 = CreateObject("adodb.Recordset")
		result2.open objCmd2
		
	result1.movenext()
	
Wend

'********************************************************************************
'* ARCHIVING ACTIVITY END
'********************************************************************************

'********************************************************************************
Function DateTime
	Styr = Year(Now)
	Stmo = Month(Now)
	if Stmo < 10 Then
		Stmo = "0" & Stmo
	end if
	Stdt = Day(Now)
	if Stdt < 10 Then
		Stdt = "0" & Stdt
	end if
	Sthr = Hour(Now)
	if Sthr < 10 Then
		Sthr = "0" & Sthr
	end if
	Stmin = Minute(Now)
	if Stmin < 10 Then
		Stmin = "0" & Stmin
	end if
	Stsec = Second(Now)
	if Stsec < 10 Then
		Stsec = "0" & Stsec
	end if
	DateTime = styr & "." & stmo & "." & stdt & " " & sthr & ":" & stmin & ":" & stsec
End Function
'********************************************************************************
