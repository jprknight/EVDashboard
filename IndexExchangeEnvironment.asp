<span>
<h1>Exchange Environment (<%=DomainName%>)</h1><br><br>
<table id="table2" class="insidetable" style="width:300px;">
<thead>
<tr>
	<th nowrap align="left">Exchange Server</th>
	<th nowrap align="left">Mailbox Count</th>
</tr>
</thead>
<tbody>
<%

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

strSQLQuery1 = "select ExchangeServer, " & _
	"sum(MbxCount) AS MbxCount " & _
	"from ExchangeMailboxes " & _
	"Where RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' " & _
	"Group By ExchangeServer " & _
	"Order by ExchangeServer"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	LineCount = LineCount + 1
	response.write "<tr><td>" & UCase((result1("ExchangeServer"))) & "</td><td>" & (result1("MbxCount")) & "</td></tr>"
	result1.movenext()
Wend
If LineCount = 0 then
	Response.Write ("<tr><td>No data</td><td>retrieved today.</td></tr>")
End if
%>
</tbody>
</table>
</span>
