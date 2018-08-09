<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!-- Main Form -->
<form id="frmconfig" method="post" action="UsageReportByArchiveByYear.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1><img src="images/users.png"> Usage Report By Archive By Year</h1>
</td>
</tr>
</table>
<a href="UsageReportByArchiveByYear.asp" title="Download data as Excel Spreadsheet"><img border="0" src="images/page_excel.png"> Download</a>
<br><br>
<!-- Table for EV Dashboard database data -->
<table id="table1" class="mytable">
<br>
<thead>
<tr>
	<th width="200px" nowrap align="left">Name</th>
	<th width="80px" nowrap align="left">Period</th>
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

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "Select Mailbox As 'Mailbox'," & _
	"LEFT(ArchivePeriod,4) AS 'ArchivePeriod'," & _
	"SUM(CAST(ItemCount as bigint)) AS 'ItemCount' " & _
	"FROM UsageReportByArchive " & _
	"GROUP BY Mailbox,LEFT(ArchivePeriod,4) " & _
	"ORDER BY Mailbox,LEFT(ArchivePeriod,4) DESC"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	Response.Write ("<tr><td>" & (result1("Mailbox")) & _
		"</td><td>" & (result1("ArchivePeriod")) & _
		"</td><td>" & (result1("ItemCount")) & "</td></tr>")
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
		sort: true,
		sort_config: {
			sort_types:['String','String','Number']
		},	
		display_all_text: " [ Show all ] ",  
		sort_select: true,
		col_1: "select",
		paging: true,  
		paging_length: 100,  
		rows_counter: true,  
		rows_counter_text: "Rows:",  
		results_per_page: ['# rows per page',[100,200,500,1000,5000]],
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