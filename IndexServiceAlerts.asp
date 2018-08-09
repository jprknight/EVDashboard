<span>
<h1>Service Alerts</h1><br><br>
<table id="table3" class="insidetable" style="width:300px;">
<thead>
<tr>
	<th nowrap width="90px" align="left">Recorded</th>
	<th nowrap width="90px" align="left">DNS Alias</th>
	<th nowrap width="90px" align="left">EV Server</th>
	<th nowrap width="100px" align="left">Role</th>
	<th nowrap width="150px" align="left">Service</th>
	<th nowrap width="70px" align="left">Status</th>
</tr>
</thead>
<tbody>
<%
If EVDConnectionString = 1 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if
'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "select RecordCreateTimestamp, " & _
	"DNSAlias, " & _
	"Server, " & _
	"Role, " & _
	"Service, " &_
	"Status " & _
	"from ServiceAlerts " & _
	"Where RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' "

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

Dim DataCount: DataCount = 0

While not result1.EOF
	response.write "<tr><td>" & dbDate((result1("RecordCreateTimestamp"))) & _
		"</td><td>" & (result1("DNSAlias")) & _
		"</td><td>" & (result1("Server")) & _
		"</td><td>" & (result1("Role")) & _
		"</td><td>" & (result1("Service")) & _
		"</td><td><h5>" & (result1("Status")) & _
		"</h5></td></tr>"
	result1.movenext()
	DataCount = DataCount + 1
Wend

If DataCount < 1 then
	response.write "<tr><td>No</td><td>Service</td><td>Alerts</td><td>Recorded</td><td>Today</td></tr>"
End if

%>
</tbody>
</table>
</span>