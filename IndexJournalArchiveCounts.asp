<span>
<br><br>
<h1>Journal Archive Counts (Live Data)</h1>
<br><br>
<table  id="table10" class="insidetable" style="width:300px;">
<thead>
<tr>
<th nowrap width="70px" align="left">EV Server</th>
<th nowrap width="140px" align="left">DB Server</th>
<th>EV Database</th>
<th nowrap width="60px" align="left">Count</th>
</tr>
</thead>
<tbody>
<%

For i = 0 to Ubound(EVServersList,2)
	If Left(EVServersList(5,i),5) = "Email" then

		'Retrieve data from EVDashboard database.
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

		strSQLQuery1 = "select count(*) AS Count from JournalArchive"

		Set objCmd1 = CreateObject("adodb.command")
		objCmd1.activeconnection = myconn1
		objCmd1.commandtimeout = strSQLExecute
		objCmd1.commandtype = adCmdText
		objCmd1.commandtext = strSQLQuery1

		Set result1 = CreateObject("adodb.Recordset")
		result1.open objCmd1
			
		While not result1.EOF
			Response.write ("<tr><td>" & EVServersList(1,i) & "</td><td>" & ucase(EVServersList(3,i)) & "</td><td>" & _
				EVServersList(4,i) & "</td><td>" & (result1("Count")) & "</td></tr>")
			result1.movenext()
		Wend
	End if
Next
%>
</tbody>
</table><br>
</span>