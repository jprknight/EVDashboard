<!--#include file="./config.asp"-->
<!--#include file="./header.asp"-->
<!--#include file="./includes/styles.asp"-->
<head>
	<script type="text/javascript">
		window.addEvent('domready', function() { myCal = new Calendar({ date: 'Y/m/d' }); });
	</script>
</head>
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

'React differently depending on whether no buttons have been pressed, refresh data or today buttons have been pressed.
If Request.Form("btnRefreshData") <> "" Then
	SQLWhereClause = "WHERE RecordCreateTimestamp between '" & request.form("date") & " 00:00:00' AND '" & request.form("date") & " 23:59:59'"
	ProcessDate = request.form("date")
Elseif Request.Form("btnToday") = "" then
	SQLWhereClause = "WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	ProcessDate = styr & "/" & stmo & "/" & stdt
Elseif Request.Form("btnToday") <> "" Then
	SQLWhereClause = "WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	ProcessDate = styr & "/" & stmo & "/" & stdt
End if
%>
<!-- Main Form -->
<form id="frmconfig" method="post" action="HourlyArchivingRates.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1>Hourly Archiving Rates</h1>
</td>
<td align="right">
<input id="date" name="date" type="text" value="<%=ProcessDate%>" />
<span><input type="submit" id="btnRefreshData" name="btnRefreshData" value="Refresh Data" /></button></span>
<span><input type="submit" id="btnToday" name="btnToday" value="Today" /></button></span><br>
</td>
</tr>
</table>
<br><br>
<!-- Table for EV Dashboard database data -->
<table id="table1" class="mytable" align="center">
<thead>
<tr>
	<th nowrap align="left" width="100px">EV Server</th>
	<th nowrap align="left">Role</th>
	<th nowrap align="left">DB Server</th>
	<th nowrap align="left" width="100px">Database</th>
	<th nowrap align="left">Archive Period</th>
	<th nowrap align="right">Count</th>
	<th nowrap align="right">Average Size (KB)</th>
	<th nowrap align="right">Total Size (KB)</th>
</tr>
</thead>
<tbody>
<%
'Retrieve data from EVDashboard database.
If EVDConnectionString = 1 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* Open database connection
Set myconn = CreateObject("adodb.connection")
myconn.ConnectionTimeout = strSQLConnect
myconn.open(EVDconnection)

strSQLQuery = "SELECT EVServer,EVRole,DBServer,EVDatabase,ArchivePeriod," & _
	"MsgCount,AverageSize,MsgCount*AverageSize as TotalSize FROM HourlyArchivingRates " & _
	SQLWhereClause

Set objCmd = CreateObject("adodb.command")
objCmd.activeconnection = myconn
objCmd.commandtimeout = strSQLExecute
objCmd.commandtype = adCmdText
objCmd.commandtext = strSQLQuery

Set result = CreateObject("ADODB.Recordset")
result.open objCmd

'************************************
'* Write data.
While not result.EOF
	Response.Write ("<tr><td>" & (result("EVServer")) & _
		"</td><td>" & (result("EVRole")) & _
		"</td><td>" & (result("DBServer")) & _
		"</td><td>" & (result("EVDatabase")) & _
		"</td><td>" & (result("ArchivePeriod")) & _
		"</td><td align=""right"">" & FormatNumber(result("MsgCount"), 0, 0, 0, -1) & _
		"</td><td align=""right"">" & FormatNumber(result("AverageSize"), 0, 0, 0, -1) & _
		"</td><td align=""right"">" & FormatNumber(result("TotalSize"), 0, 0, 0, -1) & _
		"</td></tr>")
	result.movenext()
Wend
%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		col_0: "select",    
		col_3: "select",    
		display_all_text: " [ Show all ] ",  
		sort: true,
		sort_config: {
			sort_types:['String','String','String','String','String','Number','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->
