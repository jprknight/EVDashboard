'********************************************************************************
'* SIS REPORT START
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing SIS Report."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing SIS Report."
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

'* Remove entries for the current day. As this script is rerun within the same day more 
'* rows are written with slightly different ArchivePeriod times.
'* I am still not sure why the Archive Period times are not on the hour, or if it possible to set this.
If EVDConnectionString = 1 Then
	Connection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	Connection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (SIS Report) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (SIS Report) - ConnectionString" & vbcrlf & EVDConnectionString
	OutputLogFile.writeline DateTime() & ": (SIS Report) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn = CreateObject("adodb.connection")
myconn.ConnectionTimeout = strSQLConnect
myconn.open(Connection)

strSQLQuery = "Delete FROM dbo.SISReport " & _
	"WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"

'* VERBOSE Logging.
If LoggingLevel >= 3 then
	wscript.echo DateTime() & ": (SIS Report - strSQLQuery) " & vbcrlf & strSQLQuery
	OutputLogFile.writeline DateTime() & ": (SIS Report - strSQLQuery) " & vbcrlf & strSQLQuery
End if
	
Set objCmd = CreateObject("adodb.command")
objCmd.activeconnection = myconn
objCmd.commandtimeout = strSQLExecute
objCmd.commandtype = adCmdText
objCmd.commandtext = strSQLQuery

Set result = CreateObject("adodb.Recordset")
result.open objCmd

For i = 0 to Ubound(EVServersList,2)
	If Left(EVServersList(5,i),5) = "Email" then
				
		If ConnectionString = 1 Then
			Connection1 = "Driver={SQL Server};Server=" & EVServersList(3,i) & _
				";Database=" & EVServersList(4,i) & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
		Elseif ConnectionString = 2 Then
			Connection1 = "Driver={SQL Server};Server=" & EVServersList(3,i) & _
				";Database=" & EVServersList(4,i) & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
		End if

		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (SIS Report) - Connection1" & vbcrlf & Connection1
			OutputLogFile.writeline DateTime() & ": (SIS Report) - ConnectionString" & vbcrlf & ConnectionString
			OutputLogFile.writeline DateTime() & ": (SIS Report) - Connection1" & vbcrlf & Connection1
		End if
		
		'* Open database connection
		Set myconn1 = CreateObject("adodb.connection")
		myconn1.ConnectionTimeout = strSQLConnect
		myconn1.open(Connection1)

		strSQLQuery1 = "SELECT A.ArchiveName AS AArchiveName, " & _
				"COUNT(S.IdTransaction) AS SItemsShared, " & _
				"AP.ArchivedItems AS APArchivedItems, " & _
				"AP.ArchivedItemsSize/1024 AS APArchiveSizeMB " & _			
				"FROM EnterpriseVaultDirectory.dbo.Archive A," & _
				"EnterpriseVaultDirectory.dbo.Root R," & _
				EVServersList(4,i) & ".dbo.ArchivePoint AP," & _
				EVServersList(4,i) & ".dbo.Saveset S," & _
				EVServersList(4,i) & ".dbo.SISPart SP " & _
				"WHERE SP.ParentTransactionId = S.idTransaction " & _
				"AND S.ArchivePointIdentity = AP.ArchivePointIdentity " & _
				"AND AP.ArchivePointId = R.VaultEntryID " & _
				"AND R.RootIdentity = A.RootIdentity " & _
				"GROUP BY A.ArchiveName, AP.ArchivedItems, AP.ArchivedItemsSize/1024 " & _
				"Order by ArchiveName"

		'* VERBOSE Logging.
		If LoggingLevel >= 3 then
			wscript.echo DateTime() & ": (SIS Report - strSQLQuery1) " & vbcrlf & strSQLQuery1
			OutputLogFile.writeline DateTime() & ": (SIS Report - strSQLQuery1) " & vbcrlf & strSQLQuery1
		End if		
				
		Set objCmd1 = CreateObject("adodb.command")
		objCmd1.activeconnection = myconn1
		objCmd1.commandtimeout = strSQLExecute
		objCmd1.commandtype = adCmdText
		objCmd1.commandtext = strSQLQuery1

		Set result1 = CreateObject("adodb.Recordset")
		result1.open objCmd1
		
		While not result1.EOF
			
			If EVDConnectionString = 1 Then
				Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
					EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
			Elseif EVDConnectionString = 2 Then
				Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
					EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
			End if
			
				
			'* NORMAL LOGGING.
			If LoggingLevel >= 2 then
				'wscript.echo DateTime() & ": (SIS Report) - Connection1" & vbcrlf & Connection1
				OutputLogFile.writeline DateTime() & ": (SIS Report) - EVDConnectionString" & vbcrlf & EVDConnectionString
				OutputLogFile.writeline DateTime() & ": (SIS Report) - Connection2" & vbcrlf & Connection2
			End if
			
			'* Open database connection
			Set myconn2 = CreateObject("adodb.connection")
			myconn2.ConnectionTimeout = strSQLConnect
			myconn2.open(Connection2)

			strSQLQuery2 = "INSERT INTO dbo.SISReport " & _
				"(ArchiveName,ItemsShared,ItemsArchived,ArchiveSizeMB,EVDatabase) " & _
				"VALUES ('" & (result1("AArchiveName")) & "','" & _
				(result1("SItemsShared")) & "','" & _
				(result1("APArchivedItems")) & "','" & _
				(result1("APArchiveSizeMB")) & "','" & _
				EVServersList(4,i) & "');"
				
			'* VERBOSE Logging.
			If LoggingLevel >= 3 then
				wscript.echo DateTime() & ": (SIS Report - strSQLQuery2) " & vbcrlf & strSQLQuery2
				OutputLogFile.writeline DateTime() & ": (SIS Report - strSQLQuery2) " & vbcrlf & strSQLQuery2
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
	End if
Next

'********************************************************************************
'* SIS REPORT END
'********************************************************************************