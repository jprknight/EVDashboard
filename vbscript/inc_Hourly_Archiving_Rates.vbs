'********************************************************************************
'* HOURLY ARCHIVING RATES START
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Hourly Archiving Rates."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Hourly Archiving Rates."
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

'* Open database connection
Set myconn = CreateObject("adodb.connection")
myconn.ConnectionTimeout = strSQLConnect
myconn.open(Connection)

strSQLQuery = "Delete FROM dbo.HourlyArchivingRates " & _
	"WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"

Set objCmd = CreateObject("adodb.command")
objCmd.activeconnection = myconn
objCmd.commandtimeout = strSQLExecute
objCmd.commandtype = adCmdText
objCmd.commandtext = strSQLQuery

Set result = CreateObject("adodb.Recordset")
result.open objCmd

For i = 0 to Ubound(EVServersList,2)
	If Left(EVServersList(5,i),5) = "Email" or Left(EVServersList(5,i),4) = "File" then
		If LoggingLevel >= 1 then
			wscript.echo DateTime() & ": Processing Hourly Archiving Rates for server " & EVServersList(0,i)
			OutputLogFile.writeline ""
			OutputLogFile.writeline "*************************************************************************"
			OutputLogFile.writeline ""
			OutputLogFile.writeline DateTime() & ": Processing Hourly Archiving Rates for server " & EVServersList(0,i)
			OutputLogFile.writeline DateTime() & ": Database Server" & EVServersList(3,i)
			OutputLogFile.writeline DateTime() & ": Database Name" & EVServersList(4,i)
		End if
		If ConnectionString = 1 Then
			Connection1 = "Driver={SQL Server};Server=" & EVServersList(3,i) & _
				";Database=" & EVServersList(4,i) & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
		Elseif ConnectionString = 2 Then
			Connection1 = "Driver={SQL Server};Server=" & EVServersList(3,i) & _
				";Database=" & EVServersList(4,i) & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
		End if

		'* Open database connection
		Set myconn1 = CreateObject("adodb.connection")
		myconn1.ConnectionTimeout = strSQLConnect
		myconn1.open(Connection1)

		strSQLQuery1 = "USE " & EVServersList(4,i) & " " & _
			"SELECT Convert(nvarchar(30),Min(archiveddate), 120) AS ArchivedDate, " & _
			chr(34) & "Hourly Rate" & chr(34) & " = count (*), " & _
			chr(34) & "Av Size" & chr(34) & " = sum (itemsize)/count (*) " & _
			"FROM Saveset s " & _
			"WHERE archiveddate > dateadd(hh, -24, getUTCdate ()) " & _
			"GROUP BY left (convert (varchar, s.archiveddate,20),14) " & _
			"ORDER BY left (convert (varchar, s.archiveddate,20),14) desc"

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
					EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";"
			Elseif EVDConnectionString = 2 Then
				Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
					EVDSQLDatabase & ";Trusted_Connection=yes;"
			End if
			
			'* Open database connection
			Set myconn2 = CreateObject("adodb.connection")
			myconn2.ConnectionTimeout = strSQLConnect
			myconn2.open(Connection2)

			strSQLQuery2 = "INSERT INTO dbo.HourlyArchivingRates " & _
				"(EVServer,EVRole,DBServer,EVDatabase,ArchivePeriod,MsgCount,AverageSize) " & _
				"VALUES ('" & EVServersList(0,i) & "','" & _
				EVServersList(5,i) & "','" & _
				EVServersList(3,i) & "','" & _
				EVServersList(4,i) & "','" & _
				dbDate((result1("ArchivedDate"))) & "','" & _
				(result1("Hourly Rate")) & "','" & _
				(result1("Av Size")) & "');"
			
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
'* HOURLY ARCHIVING RATES END
'********************************************************************************