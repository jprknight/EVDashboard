<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!-- Main Form -->
<form id="frmconfig" method="post" action="UsageReportByArchiveByYear.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1><img src="images/users.png"> Usage Report By Archive</h1>
</td>
</tr>
</table>
<a href="UsageReportByArchive.asp" title="Download data as Excel Spreadsheet"><img border="0" src="images/page_excel.png"> Download</a>
<br><br>
<!-- Table for EV Dashboard database data -->
<table id="table1" class="mytable">
<br>
<thead>
<tr>
	<th width="200px" nowrap align="left">Name</th>
	<th width="80px" nowrap align="left">Item Count</th>
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

'* Bug Fix #3180056: UsageReportByArchive timeout when table is empty
'* QUERY 1: RETRIEVE NUMBER OF ROWS FROM UsageReportByArchive TABLE.
'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "Select Count(*) AS RowCount FROM UsageReportByArchive"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	intRowCount = (result1("RowCount"))
	'* Only one row returned, the below should not be required.
	' result1.movenext()
Wend

'* QUERY 2: SELECT THE ROWS FROM THE UsageReportByArchive TABLE.
If intRowCount > 0 then
	'* Open database connection
	Set myconn2 = CreateObject("adodb.connection")
	myconn2.ConnectionTimeout = strSQLConnect
	myconn2.open(Connection1)

	strSQLQuery2 = "Select Mailbox As 'Mailbox'," & _
		"SUM(CAST(ItemCount as bigint)) AS 'ItemCount' " & _
		"FROM UsageReportByArchive " & _
		"GROUP BY Mailbox " & _
		"ORDER BY Mailbox"

	Set objCmd2 = CreateObject("adodb.command")
	objCmd2.activeconnection = myconn2
	objCmd2.commandtimeout = strSQLExecute
	objCmd2.commandtype = adCmdText
	objCmd2.commandtext = strSQLQuery2

	Set result2 = CreateObject("adodb.Recordset")
	result2.open objCmd2

	While not result2.EOF
		Response.Write ("<tr><td>" & (result1("Mailbox")) & _
			"</td><td>" & (result1("ItemCount")) & "</td></tr>")
		result2.movenext()
	Wend
Else
	'* No rows found.
	Response.Write ("<tr><td>No rows found.</td></tr>")
End if
%>
</tbody>
</table><br>
</form>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		on_keyup: true,
		on_keyup_delay: 1200,
		sort: true,
		sort_config: {
			sort_types:['String','Number']
		},	
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