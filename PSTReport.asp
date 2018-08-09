<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<head>
	<script type="text/javascript">
		window.addEvent('domready', function() { myCal = new Calendar({ date: 'Y/m/d' }); });
	</script>
</head>
<%
HTML_SPAN_GREEN = "<span style=""color: #006600; font-weight: bold;"">"
HTML_SPAN_GREY  = "<span style=""color: #666666; font-weight: bold; font-style: italic;"">"
HTML_SPAN_RED   = "<span style=""color: #ff3333; font-weight: bold;"">"
HTML_SPAN_END   = "</span>"

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
Today = styr & "/" & stmo & "/" & stdt

'React differently depending on whether no buttons have been pressed, refresh data or today buttons have been pressed.
If Request.Form("btnRefreshData") <> "" Then
	SQLWhereClause = "WHERE MigrationStartTime between '" & request.form("date") & " 00:00:00' AND '" & request.form("date") & " 23:59:59'"
	ProcessDate = request.form("date")
'Elseif Request.Form("btnToday") = "" then
'	SQLWhereClause = "WHERE MigrationStartTime between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
'		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
'	ProcessDate = styr & "/" & stmo & "/" & stdt
Else 'if Request.Form("btnToday") <> "" Then
	SQLWhereClause = "WHERE MigrationStartTime between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	ProcessDate = styr & "/" & stmo & "/" & stdt
End if
%>
<!-- Main Form -->
<form id="frmconfig" method="post" action="PSTReport.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1><img src="images/users.png"> PST Migration Report</h1>
</td>
<td align="right">
<input id="date" name="date" type="text" value="<%=ProcessDate%>" class="dateselector"/>
<span><input type="submit" id="btnRefreshData" name="btnRefreshData" value="Refresh Data" class="pgInp"/></button></span>
<span><input type="submit" id="btnToday" name="btnToday" value="Today" class="pgInp" /></button></span><br>
</td>
</tr>
</table>
<a href="PSTReportExport.asp" title="Download data as Excel Spreadsheet"><img border="0" src="images/page_excel.png"> Download</a>
<br><br>
<!-- Table for EV Dashboard database data -->
<table id="table1" class="mytable" align="center">
<br>
<thead>
	<tr>
		<th nowrap align="left">PST FileName</th>
		<th nowrap align="left">Mailbox Name</th>
		<th nowrap align="left">Size (KB)</th>
		<th nowrap align="left">Current Status</th>
		<th nowrap align="center">Total Items</th>
		<th nowrap align="center">Items Archived</th>
		<th nowrap align="center">Items Stored</th>
		<th nowrap align="left">Start Time</th>
		<th nowrap align="left">Completion Time</th>
	</tr>
</thead>
<tbody>
<%
'Retrieve data from EnterpriseVaultDirectory database.
strDatabaseName = "EnterpriseVaultDirectory"
		
If ConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
		strDatabase & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
Elseif ConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
		strDatabase & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

' ********************************************************************
' * Revised to include items not previously attempted (or in progress)
' ********************************************************************
Response.Flush

strSQLQuery1 = "SELECT PST.FileSpecification as PSTPath, " & _
		       "PST.FileSize as PSTSize, " & _
		       "EME.MbxDisplayName AS MailboxName, " & _
		       "PST.MigrationStatus as LastStatus, " & _
		       "PSTMH.MigratedMsgCount as ImportedItems, " & _
		       "PSTMH.MigratedMsgCount-PSTMH.NotEligibleForArchMsgCount as ArchivedItems, " & _
		       "PSTMH.NotEligibleForArchMsgCount as StoredItems, " & _
		       "PSTMH.MigrationStartTime as StartTime, " & _
		       "PSTMH.MigrationEndTime as EndTime " & _
		"FROM EnterpriseVaultDirectory.dbo.PstFile PST " & _
		     "INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId " & _
		     "INNER JOIN EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry EME ON PST.ExchangeMailboxEntryId = EME.ExchangeMailboxEntryId " & _
		     "INNER JOIN ( " & _
		     "    SELECT PST.FileSpecification as PSTF, Max(PSTMH.MigrationEndTime) as PSTT " & _
		     "    FROM EnterpriseVaultDirectory.dbo.PstFile PST " & _
		     "         INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId " & _
		     "    GROUP BY PST.FileSpecification " & _
		     "         ) AS PSTMaxTime ON PST.FileSpecification = PSTMaxTime.PSTF AND PSTMH.MigrationEndTime = PSTMaxTime.PSTT " & _
			SQLWhereClause & _
		" ORDER BY PST.FileSpecification"

if ProcessDate = Today then
	strSQLQuery1 =	"SELECT PSTPath, PSTSize, MailboxName, LastStatus, ImportedItems, ArchivedItems, StoredItems, StartTime, EndTime " & _
			"FROM " & _
			"( " & _
			"SELECT  PST.FileSpecification as PSTPath,  " & _
			"		PST.FileSize as PSTSize,  " & _
			"		EME.MbxDisplayName AS MailboxName,  " & _
			"		PST.MigrationStatus as LastStatus,  " & _
			"		PSTMH.MigratedMsgCount as ImportedItems,  " & _
			"		PSTMH.MigratedMsgCount-PSTMH.NotEligibleForArchMsgCount as ArchivedItems,  " & _
			"		PSTMH.NotEligibleForArchMsgCount as StoredItems,  " & _
			"		PSTMH.MigrationStartTime as StartTime,  " & _
			"		PSTMH.MigrationEndTime as EndTime  " & _
			"FROM EnterpriseVaultDirectory.dbo.PstFile PST  " & _
			"		INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId  " & _
			"		INNER JOIN EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry EME ON PST.ExchangeMailboxEntryId = EME.ExchangeMailboxEntryId  " & _
			"		INNER JOIN (  " & _
			"			SELECT PST.FileSpecification as PSTF, Max(PSTMH.MigrationEndTime) as PSTT  " & _
			"			FROM EnterpriseVaultDirectory.dbo.PstFile PST  " & _
			"				INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId  " & _
			"			GROUP BY PST.FileSpecification  " & _
			"			) AS PSTMaxTime ON PST.FileSpecification = PSTMaxTime.PSTF AND PSTMH.MigrationEndTime = PSTMaxTime.PSTT  " & _
				SQLWhereClause & _
			"UNION ALL " & _
			"SELECT  PST.FileSpecification as PSTPath,  " & _
			"		PST.FileSize as PSTSize,  " & _
			"		EME.MbxDisplayName AS MailboxName,  " & _
			"		PST.MigrationStatus as JobResult, " & _
			"		0 as ImportedItems, " & _
			"		0 as ArchivedItems, " & _
			"		0 as StoredItems, " & _
			"		PST.LastAccessedTime as StartTime, " & _
			"		PST.LastCopiedTime as EndTime " & _
			"FROM EnterpriseVaultDirectory.dbo.PstFile PST  " & _
			"		INNER JOIN EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry EME ON PST.ExchangeMailboxEntryId = EME.ExchangeMailboxEntryId  " & _
			"WHERE NOT PST.PstFileEntryId IN (SELECT DISTINCT PSTMH.PstFileEntryId FROM EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH) " & _
			"  AND PST.MigrationStatus <> 100 " & _
			") Data " & _
			"ORDER BY PSTPath "
End If

Response.Write "<!-- SQL Query: " & strSQLQuery1 & " -->"
Response.Flush

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	Response.Write "<!-- " & result1("MailboxName") & " -->" & vbCrLf

	Response.Write("<tr><td>" & result1("PSTPath") & _
		"</td><td>" & (result1("MailboxName")) & _
		"</td><td align=""right"">" & (FormatNumber(result1("PSTSize"),0)) & _
		"</td><td>")
	Select Case result1("LastStatus")
		Case "0"		: Response.Write HTML_SPAN_GREY & "Not Ready" & HTML_END_SPAN
		Case "100"		: Response.Write HTML_SPAN_GREY & "Do Not Migrate" & HTML_END_SPAN
		Case "200"		: Response.Write "Ready to Copy"
		Case "300"		: Response.Write "Copying"
		Case "400"		: Response.Write HTML_SPAN_RED & "Copy Failed" & HTML_END_SPAN
		Case "500"		: Response.Write "Ready to Backup"
		Case "600"		: Response.Write "Backing Up"
		Case "700"		: Response.Write HTML_SPAN_RED & "Backup Failed" & HTML_END_SPAN
		Case "800"		: Response.Write "Ready to Migrate"
		Case "900"		: Response.Write "Migrating"
		Case "1000"		: Response.Write HTML_SPAN_RED & "Failed" & HTML_END_SPAN
		Case "1100"		: Response.Write "Ready for Post Processing"
		Case "1200"		: Response.Write "Completing"
		Case "1300"		: Response.Write HTML_SPAN_RED & "Completion Failed" & HTML_END_SPAN
		Case "1400"		: Response.Write HTML_SPAN_GREEN & "Complete" & HTML_END_SPAN
		Case Else		: Response.Write result1("LastStatus")
	End Select

	Response.Write(	"</td><td align=""right"">" & (FormatNumber(result1("ImportedItems"),0)) & _
		"</td><td align=""right"">" & (FormatNumber(result1("ArchivedItems"),0)) & _
		"</td><td align=""right"">" & (FormatNumber(result1("StoredItems"),0)) & _
		"</td><td>" & (result1("StartTime")) &_
		"</td><td>" & (result1("EndTime")) & "</td></tr>" & vbCrLf)
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
			sort_types:['String','String','String','String','Number','Number','Number','String','String']
		},	
		col_1: "select",    
		col_3: "select",    
		display_all_text: " [ Show all ] ",  
		sort_select: true,
		paging: true,  
		paging_length: 100,  
		rows_counter: true,  
		rows_counter_text: "Rows:",  
		results_per_page: ['# rows per page',[100,200,500,1000,5000,10000]],
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