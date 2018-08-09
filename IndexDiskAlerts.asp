<span>
<br><br>
<h1>Disk Space Alerts</h1><br><br>
<table id="table7" class="insidetable">
<thead>
	<tr>
		<th nowrap align="left">Recorded</th>
		<th nowrap align="left">Server</th>
		<th nowrap align="left">Drive</th>
		<th nowrap align="left">Size (GB)</th>
		<th nowrap align="left">Free (GB)</th>
		<th nowrap align="left">% Free</th>
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

strSQLQuery1 = "SELECT RecordCreateTimestamp AS 'RecordCreateTimestamp'," & _
	"EVServer AS 'Server'," & _
	"Drive AS 'Drive'," & _
	"SizeGB AS 'SizeGB'," & _
	"FreeGB AS 'FreeGB'," & _
	"PercentageFree AS 'PercentageFree' " & _
	"FROM DiskUsageInfo " & _
	"WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59' "
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF

	strRecordCreateTimestamp = dbDate((result1("RecordCreateTimestamp")))
	strServer = (result1("Server"))
	strDrive = (result1("Drive"))
	strSizeGB = (result1("SizeGB"))
	strFreeGB = (result1("FreeGB"))
	strPercentageFree = (result1("PercentageFree"))
	
	If Len(strPercentageFree) >= 1 then
		If strPercentageFree <= 25 then 'AND strFreeGB <= 150 then
			response.write "<tr><td>" & strRecordCreateTimestamp & _
				"</td><td>" & ucase(strServer) & _
				"</td><td>" & strDrive & _
				"</td><td>" & strSizeGB & _
				"</td><td>" & strFreeGB & _
				"</td><td>" & strPercentageFree & _
				"</td></tr>"
		End if
	End if
	result1.movenext()
Wend
%>
</tbody>
</table><br><br>
</span>