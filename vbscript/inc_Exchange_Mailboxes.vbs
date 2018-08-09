'********************************************************************************
'* ECHANGE MAILBOXES
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Exchange Mailboxes."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Exchange Mailboxes."
End if

EVDConnectionString = 1

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

Redim EXServersMailboxes(2,0)
EXServersMailboxes(0,0) = strservername
EXServersMailboxes(1,0) = database
EXServersMailboxes(2,0) = mbnum

set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPrivateMDB);name,distinguishedName;subtree"
'pfQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPublicMDB);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute

i = 0

While Not Rs.EOF
	objmailstorename = "LDAP://" & Rs.Fields("distinguishedName") 
	mbnum = 0
	rangeStep = 999
	lowRange = 0
	highRange = lowRange + rangeStep
	quit = false
	set objmailstore = getObject(objmailstorename)
	Do until quit = true
		on error resume next
		strCommandText = "homeMDBBL;range=" & lowRange & "-" & highRange
		objmailstore.GetInfoEx Array(strCommandText), 0
		if err.number <> 0 then quit = true
		varReports = objmailstore.GetEx("homeMDBBL")
		if quit <> true then mbnum = mbnum + ubound(varReports)+1
			lowRange = highRange + 1
			highRange = lowRange + rangeStep
	loop
	err.clear
	
	strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
	
	Redim preserve EXServersMailboxes(ubound (EXServersMailboxes, 1), ubound (EXServersMailboxes, 2) + 1)
	EXServersMailboxes(0,i) = strservername
	EXServersMailboxes(1,i) = Rs.Fields("name")
	EXServersMailboxes(2,i) = EXServersMailboxes(2,i) + mbnum
	
	i = i + 1
	
	Rs.MoveNext
Wend

Const adCmdText = 1

For i = 0 to Ubound(EXServersMailboxes,2)
	'* For some strange reason I had a null row written the ExchangeMailboxes table. The below if then else line stops this from happening.
	If EXServersMailboxes(0,i) <> "" AND EXServersMailboxes(1,i) <> "" AND EXServersMailboxes(2,i) <> "" then
		If EVDConnectionString = 1 Then
			EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
		Elseif EVDConnectionString = 2 Then
			EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
				EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
		End if

		'* NORMAL LOGGING.
		If LoggingLevel >= 2 then
			'wscript.echo DateTime() & ": (Exchange Mailboxes) - EVDconnection" & vbcrlf & EVDconnection
			OutputLogFile.writeline DateTime() & ": (Exchange Mailboxes) - EVDconnection" & vbcrlf & EVDconnection
		End if
		
		'* Open database connection
		Set myconn = CreateObject("adodb.connection")
		myconn.ConnectionTimeout = strSQLConnect
		myconn.open(EVDconnection)

		strSQLQuery = "UPDATE dbo.ExchangeMailboxes " & _
			"SET ExchangeServer='" & EXServersMailboxes(0,i) & "'," & _
			"Mailstore='" & EXServersMailboxes(1,i) & "'," & _
			"MbxCount='" & EXServersMailboxes(2,i) & "' " & _
			"WHERE ExchangeServer='" & EXServersMailboxes(0,i) & "' " & _
			"AND Mailstore='" & EXServersMailboxes(1,i) & "' " & _
			"AND RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
			"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' " & _
			"IF @@ROWCOUNT=0 " & _
			"INSERT INTO dbo.ExchangeMailboxes " & _
			"(ExchangeServer,Mailstore,MbxCount) " & _
			"VALUES ('" & EXServersMailboxes(0,i) & "','" & _
			EXServersMailboxes(1,i) & "','" & _
			EXServersMailboxes(2,i) & "');"

		'* VERBOSE LOGGING.
		If LoggingLevel >= 3 then
			'wscript.echo DateTime() & ": (Exchange Mailboxes) - strSQLQuery" & vbcrlf & strSQLQuery
			OutputLogFile.writeline DateTime() & ": (Exchange Mailboxes) - strSQLQuery" & vbcrlf & strSQLQuery
		End if	
			
		Set objCmd = CreateObject("adodb.command")
		objCmd.activeconnection = myconn
		objCmd.commandtimeout = strSQLExecute
		objCmd.commandtype = adCmdText
		objCmd.commandtext = strSQLQuery

		Set result = CreateObject("ADODB.Recordset")
		result.open objCmd
	end if
Next

Rs.Close
Conn.Close
Set Rs = Nothing
Set Com = Nothing
Set Conn = Nothing

'********************************************************************************
'* ECHANGE MAILBOXES END
'********************************************************************************