<span>
<br><br>
<h1>MSMQ Message Count Alerts</h1><br><br>
<table id="table5" class="insidetable" style="width:300px;">
<thead>
<tr>
	<th nowrap width="80px" align="left">Recorded</th>
	<th nowrap width="100px" align="left">EV Server</th>
	<th nowrap width="270px" align="left">MSMQ Path</th>
	<th nowrap width="70px" align="left">Count</th>
	<th nowrap width="70px" align="left">Size (KB)</th>
</tr>
</thead>
<tbody>
<%
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

strSQLQuery1 = "select RecordCreateTimestamp, " & _
	"EVServer, " & _
	"MSMQPath, " & _
	"RecordCount, " &_
	"FileSize " & _
	"from MSMQMessageCountAlerts " & _
	"Where RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' "

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

Dim MSMQDataCount: MSMQDataCount = 0

While not result1.EOF	
	response.write "<tr><td>" & dbDate((result1("RecordCreateTimestamp"))) & _
		"</td><td>" & (result1("EVServer")) & _
		"</td><td>" & (result1("MSMQPath")) & _
		"</td><td>" & (result1("RecordCount")) & _
		"</td><td>" & (result1("FileSize")) & _
		"</td></tr>"
	result1.movenext()
	MSMQDataCount = MSMQDataCount + 1
Wend

If MSMQDataCount < 1 then
	response.write "<tr><td>No</td><td>MSMQ</td><td>Alerts</td><td>Recorded</td><td>Today</td></tr>"
End if

%>
</tbody>
</table>
</span>