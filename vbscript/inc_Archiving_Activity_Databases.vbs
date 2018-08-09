'********************************************************************************
'* ARCHIVING ACTIVITY START
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Archiving Activity Databases."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Archiving Activity Databases."
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

strSQLQuery1 = "SELECT Mailbox AS 'Mailbox'," & _
	"ExchangeServer AS 'ExchangeServer'," & _
	"EVServer AS 'EVServer'," & _
	"DBServer AS 'DBServer'," & _
	"EVDatabase AS 'EVDatabase'," & _
	"root_rootidentity AS 'root_rootidentity'," & _
	"ExchangeServerIdentity AS 'ExchangeServerIdentity' " & _
	"FROM ArchivingActivity " & _
	"WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF

	'****************************************************
	'* Connect to mailbox archiving database and retrieve NumItemsArchived
	
		'* The following need to be in numerical order, so all variables are set correctly.
		strMailbox = result1.Fields.Item(0)
		strExchangeComputer = result1.Fields.Item(1)
		strDBServer = result1.Fields.Item(3)
		strEVDatabase = result1.Fields.Item(4)
		strroot_rootidentity = result1.Fields.Item(5)
		strExchangeServerIdentity = result1.Fields.Item(6)
		
		If LoggingLevel >= 1 then
			wscript.echo "Processing mailbox archive " & chr(34) & strMailbox & chr(34) & "."
			OutputLogFile.writeline "Processing mailbox archive " & chr(34) & strMailbox & chr(34) & "."
		End if
		
		'* DATA RETRIEVE SQL QUERY 2.
		
		If ConnectionString = 1 Then
			Connection2 = "Driver={SQL Server};Server=" & strDBServer & ";Database=" & _
				strEVDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLAltPortNumber & ";"
		Elseif ConnectionString = 2 Then
			Connection2 = "Driver={SQL Server};Server=" & strDBServer & ";Database=" & _
				strEVDatabase & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
		End if
		
		wscript.echo Connection2
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection2)

		REM strSQLQuery2 = "Declare @Days int " & _
			REM "set @Days = 14 " & _
			REM "SELECT Count(*) AS 'NumItemsArchived', " & _
			REM "archivepointid " & _
			REM "FROM saveset,archivepoint " & _
			REM "WHERE archivepoint.archivepointid=" & strroot_rootidentity & " " & _
			REM "AND saveset.archivepointidentity = archivepoint.archivepointidentity " & _
			REM "AND saveset.archiveddate > dateadd(day, -@Days ,getdate()) " & _
			REM "GROUP BY archivepointid"
			REM "NumItemsArchived " & _
			REM "ORDER BY " & chr(34) & strMailbox & chr(34) & "," & strExchangeComputer
		
		strSQLQuery2 = "Declare @Days int " & _
			"set @Days = 14 " & _
			"SELECT Count(*) AS 'NumItemsArchived' " & _
			"FROM saveset " & _
			"INNER JOIN archivepoint " & _
			"ON " & strroot_rootidentity & " = archivepoint.archivepointid " & _
			"WHERE saveset.archiveddate > dateadd(day, -@Days ,getdate()) " '& _
			'"GROUP BY archivepointid " '& _
			'"ORDER BY " & chr(34) & strMailbox & chr(34) & "," & strExchangeComputer
		
		
		wscript.echo strSQLQuery2
		
		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2

		Set result2 = CreateObject("adodb.Recordset")
		result2.open objCmd2
		
		While not result2.EOF
			
			wscript.echo "TEST:" & (result2("NumItemsArchived"))
			
			REM '************************************************
			REM '* UPDATE SQL QUERY START.
			REM Dim strNumItemsArchived: strNumItemsArchived = (result2("strNumItemsArchived"))
			
			REM If EVDConnectionString = 1 Then
				REM Connection3 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
					REM EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";"
			REM Elseif EVDConnectionString = 2 Then
				REM Connection3 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
					REM EVDSQLDatabase & ";Trusted_Connection=yes;"
			REM End if

			REM '* Open database connection
			REM Set myconn3 = CreateObject("adodb.connection")
			REM myconn3.ConnectionTimeout = strSQLConnect
			REM myconn3.open(Connection3)
			
			REM strSQLQuery3 = "UPDATE dbo.ArchivingActivity " & _
				REM "SET NumItemsArchived='" & (result2("NumItemsArchived")) & "'" & _
				REM "WHERE root_rootidentity = '" & strroot_rootidentity & "' AND " & _
				REM "RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
				REM "AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"

			REM Set objCmd3 = CreateObject("adodb.command")
			REM objCmd3.activeconnection = myconn3
			REM objCmd3.commandtimeout = strSQLExecute
			REM objCmd3.commandtype = adCmdText
			REM objCmd3.commandtext = strSQLQuery3

			REM Set result3 = CreateObject("adodb.Recordset")
			REM result3.open objCmd3
			REM '* UPDATE SQL QUERY FINISH.
			REM '************************************************
			
			result2.movenext()
			
		Wend
		
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
