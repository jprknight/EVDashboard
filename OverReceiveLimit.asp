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
<form id="frmconfig" method="post" action="OverReceiveLimit.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1><img src="images/users.png"> Users Over Receive Limit</h1>
</td>
<td align="right">
<input id="date" name="date" type="text" value="<%=ProcessDate%>" class="dateselector"/>
<span><input type="submit" id="btnRefreshData" name="btnRefreshData" value="Refresh Data" class="pgInp"/></button></span>
<span><input type="submit" id="btnToday" name="btnToday" value="Today" class="pgInp" /></button></span><br>
</td>
</tr>
</table>
<a href="OverReceiveLimitExport.asp" title="Download data as Excel Spreadsheet"><img border="0" src="images/page_excel.png"> Download</a>
<br><br>
<!-- Table for EV Dashboard database data -->
<table id="table1" class="mytable" align="center">
<br>
<thead>
	<tr>
		<th nowrap align="left">Name</th>
		<th nowrap align="left">Exchange Server</th>
		<th nowrap align="left">Over Receive (MB)</th>
		<th nowrap align="left">Warning (MB)</th>
		<th nowrap align="left">Send (MB)</th>
		<th nowrap align="left">Receive (MB)</th>
		<th nowrap align="left">Mbx Size (MB)</th>
		<th nowrap align="left"># Items (Mailbox)</th>
		<th nowrap align="left">Last Archived</th>
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

strSQLQuery1 = "Select Mailbox,ExchangeServer,OverReceiveMB,WarningMB,SendMB,ReceiveMB," & _
	"MailboxSizeMB,NumItemsMailbox,LastArchived from OverReceiveLimit " & _
	SQLWhereClause
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF		
	Response.Write("<tr><td>" & (result1("Mailbox")) & _
		"</td><td>" & UCase((result1("ExchangeServer"))) & _
		"</td><td>" & (result1("OverReceiveMB")) & _
		"</td><td>" & (result1("WarningMB")) & _
		"</td><td>" & (result1("SendMB")) & _
		"</td><td>" & (result1("ReceiveMB")) & _
		"</td><td>" & (result1("MailboxSizeMB")) & _
		"</td><td>" & (result1("NumItemsMailbox")) & _
		"</td><td>" & (result1("LastArchived")) & "</td></tr>")
		'"</td><td>" & dbDate((result1("LastArchived"))) & "</td></tr>")
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
			sort_types:['String','Number','Number','Number','Number','Number','Number','String','String']
		},	
		col_1: "select",    
		display_all_text: " [ Show all ] ",  
		sort_select: true,
		paging: true,  
		paging_length: 100,  
		rows_counter: true,  
		rows_counter_text: "Rows:",  
		results_per_page: ['# rows per page',[100,200,500,1000]],
		btn_reset: true,
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