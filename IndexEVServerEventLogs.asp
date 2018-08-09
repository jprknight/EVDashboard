<span>
<%
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
%>
<!-- Table for page heading -->
<table width="100%" class="Main">
	<tr>
		<td>
			<h1>EV Server Event Logs</h1><br>
			<h3>Warnings and errors recorded in the 24 hours preceeding the time when the data collector was last run.</h3>
		</td>
	</tr>
</table>
<br><br>
<!-- Table for EV Dashboard database data -->
<table id="table8" class="insidetable">
<br>
<thead>
<tr>
	<th nowrap align="left"></th>
	<th width="70" nowrap align="left">EV Server</th>
	<th nowrap align="left">Event Type</th>
	<th nowrap align="left">Event Code</th>
	<th width="100" nowrap align="left">Time Written</th>
	<th nowrap align="left">Message</th>
</tr>
</thead>
<tbody>
<%
'Retrieve data from EVDashboard database.
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

strSQLQuery1 = "Select EventType,EVServer,EventCode,TimeWritten,Message " & _
	"From EVServerEventLogs " & "WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' " & _
		"AND EventType <> 'Information'"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF

	LineCount = LineCount + 1
	
	strEventType = (result1("EventType"))
	strEVServer = (result1("EVServer"))
	strEventCode = (result1("EventCode"))
	strTimeWritten = (result1("TimeWritten"))
	strMessage = (result1("Message"))
	
	If strEventType = "Warning" Then
		strEventTypeIcon = "<img src=./images/ico_warning.gif>"
	Elseif strEventType = "Error" then
		strEventTypeIcon = "<img src=./images/ico_error.gif>"
	End if
		
	Response.Write ("<tr><td>" & strEventTypeIcon & _
		"</td><td>" & strEVServer & _
		"</td><td>" & strEventType & _
		"</td><td>" & strEventCode & _
		"</td><td>" & strTimeWritten & _
		"</td><td>" & strMessage & "</td></tr>")
	result1.movenext()
Wend

If LineCount = 0 then
	Response.Write ("<tr><td>No</td><td>warnings</td><td>or</td><td>errors</td>" & _
		"<td>recorded</td><td>today</td></tr></tbody></table>")
Else
	Response.Write ("</tbody></table>")
End if

%>
<br><br>
</span>