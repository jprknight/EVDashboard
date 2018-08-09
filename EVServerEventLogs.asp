<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
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

Dim SQLWhereClause

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
<form id="frmconfig" method="post" action="EVServerEventLogs.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1>EV Server Event Logs</h1>
</td>

<td align="right">
<input id="date" name="date" type="text" value="<%=ProcessDate%>" class="dateselector" />
<span style="text-align:left;width:100%;"><input type="submit" id="btnRefreshData" name="btnRefreshData" value="Refresh Data" class="pgInp" /></button></span>
<span style="text-align:left;width:100%;"><input type="submit" id="btnToday" name="btnToday" value="Today" class="pgInp" /></button></span><br>
</td>
</tr>
</table>
<br><br>
<!-- Table for EV Dashboard database data -->
<table id="table1" class="mytable" align="center">
<br>
<thead>
<tr>
	<th nowrap align="left"></th>
	<th nowrap align="left">EV Server</th>
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
	"From EVServerEventLogs " & SQLWhereClause

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	strEventType = (result1("EventType"))
	strEVServer = (result1("EVServer"))
	strEventCode = (result1("EventCode"))
	strTimeWritten = (result1("TimeWritten"))
	strMessage = (result1("Message"))
	
	If strEventType = "Warning" Then
		strEventTypeIcon = "<img src=./images/ico_warning.gif>"
	Elseif strEventType = "Error" then
		strEventTypeIcon = "<img src=./images/ico_error.gif>"
	Elseif strEventType = "Information" then
		strEventTypeIcon = "<img src=./images/ico_success.gif>"
	End if
		
	Response.Write ("<tr><td>" & strEventTypeIcon & _
		"</td><td>" & strEVServer & _
		"</td><td>" & strEventType & _
		"</td><td>" & strEventCode & _
		"</td><td>" & strTimeWritten & _
		"</td><td>" & strMessage & "</td></tr>")
	result1.movenext()
Wend

%>
</tbody>
</table><br>
</form>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		on_keyup: true,
		on_keyup_delay: 1200,
		filters_row_index: 1,
		extensions: { 
		name:['ColsVisibility','FiltersRowVisibility'], 
		src:['./js/TFExt_ColsVisibility.js','./js/TFExt_FiltersRowVisibility.js'],
		description:['Columns visibility manager','Expand/Collapse filters row'], 
		initialize:[function(o){o.SetColsVisibility();},function(o){o.SetFiltersRowVisibility();}] 
		},
		showHide_cols_text: 'Tick columns to hide: ',
		sort: true,
		sort_config: {
			sort_types:['String','String','Number','String','String']
		},	
		display_all_text: " [ Show all ] ",  
		sort_select: true,
		paging: true,  
		paging_length: 100,  
		rows_counter: true,  
		rows_counter_text: "Rows:",  
		results_per_page: ['# rows per page',[100,200,500,1000]],
		btn_reset: true,
		display_all_text: " [ Show all ] ",  
		sort_select: true,
		col_1: "select",
		col_2: "select",
		col_3: "select",
		alternate_rows: true,
		loader: true,  
		loader_html: '<img src="./images/loader.gif" alt="" ' +
							'style="vertical-align:middle; margin:0 5px 0 5px" />' +
							'<span>Loading...</span>',
		loader_css_class: 'myLoader',
		status_bar: true,
		enter_key: false,
		fill_slc_on_demand: true,
		btn_reset: true,  
		btn_reset_html: '<input type="button" value="Reset" class="pgInp" />',  
		btn_next_page_text: 'Next >',
		btn_prev_page_text: '< Prev',
		btn_last_page_text: 'Last >>',
		btn_first_page_text: '<< First'
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->