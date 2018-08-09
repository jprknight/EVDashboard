<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
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

Dim SQLWhereClause

'React differently depending on whether no buttons have been pressed, refresh data or today buttons have been pressed.
If Request.Form("btnRefreshData") <> "" Then
	SQLWhereClause = "AND RecordCreateTimestamp between '" & request.form("date") & " 00:00:00' AND '" & request.form("date") & " 23:59:59'"
	ProcessDate = request.form("date")
Elseif Request.Form("btnToday") = "" then
	SQLWhereClause = "AND RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	ProcessDate = styr & "/" & stmo & "/" & stdt
Elseif Request.Form("btnToday") <> "" Then
	SQLWhereClause = "AND RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	ProcessDate = styr & "/" & stmo & "/" & stdt
End if
%>
<!-- Main Form -->
<form id="frmconfig" method="post" action="EnabledMailboxesCounts.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1><img src="images/users.png"> Enabled Mailboxes - Counts</h1>
</td>
<td align="right">
<input id="date" name="date" type="text" value="<%=ProcessDate%>" class="dateselector" />
<span style="text-align:left;width:100%;"><input type="submit" id="btnRefreshData" name="btnRefreshData" value="Refresh Data" class="pgInp" /></button></span>
<span style="text-align:left;width:100%;"><input type="submit" id="btnToday" name="btnToday" value="Today" class="pgInp" /></button></span><br>
</td>
</tr>
</table>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th nowrap align="left">EV Server</th>
<th nowrap align="left">DB Server</th>
<th nowrap align="left">Database</th>
<th nowrap align="left">Count</th>
</tr>
</thead>
<tbody>
<%

Dim TotalUsers
TotalUsers = 0

'*********************************************************************************************************
For i = 0 to Ubound(EVServersList,2)
	If Left(EVServersList(5,i),5) = "Email" then
	
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

		strSQLQuery1 = "SELECT Count(*) AS Count FROM EnabledMailboxes " & _
			"WHERE EVDatabase = '" & EVServersList(4,i) & "' " & SQLWhereClause

		Set objCmd1 = CreateObject("adodb.command")
		objCmd1.activeconnection = myconn1
		objCmd1.commandtimeout = strSQLExecute
		objCmd1.commandtype = adCmdText
		objCmd1.commandtext = strSQLQuery1

		Set result1 = CreateObject("adodb.Recordset")
		result1.open objCmd1
		
		While not result1.EOF
			response.write ("<tr><td>" & EVServersList(1,i) & "</td><td>" & EVServersList(3,i) & "</td>" & _
				"<td>" & EVServersList(4,i) & "</td><td>" & (result1("Count")) & "</td></tr>")
			TotalUsers = TotalUsers + (result1("Count"))
			result1.movenext()
		Wend
	End if
Next


response.write ("<tr><td>...</td><td></td><td></td><td>...</td></tr>" & _
	"<tr><td>Total Users</td><td></td><td></td><td>" & TotalUsers & "</td></tr>")

%>
</tbody>
</table><br>
</form>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		filters_row_index: 1,
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number']
		},	
		sort_select: true,
		alternate_rows: true,
		loader: true,  
		loader_html: '<img src="./images/loader.gif" alt="" ' +
							'style="vertical-align:middle; margin:0 5px 0 5px" />' +
							'<span>Loading...</span>',
		loader_css_class: 'myLoader',
		enter_key: false,
		fill_slc_on_demand: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->