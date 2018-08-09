<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br>
<h1>Organisation Archiving History</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
	<th width="80px" nowrap align="left">Month</th>
	<th width="80px" nowrap align="left">DNS Alias</th>
	<th width="80px" nowrap align="left">EV Server</th>
	<th width="100px" nowrap align="left">EV Database</th>
	<th width="80px" nowrap align="left">Count</th>
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

strSQLQuery1 = "SELECT ArchivedDate,DNSAlias,EVServer,EVDatabase,RecordCount " & _
	"FROM OrganisationArchivingHistory " & _
	"Order by ArchivedDate DESC"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	Response.Write ("<tr><td>" & (result1("ArchivedDate")) & _
		"</td><td>" & UCase((result1("DNSAlias"))) & _
		"</td><td>" & (result1("EVServer")) & _
		"</td><td>" & (result1("EVDatabase")) & _
		"</td><td>" & (result1("RecordCount")) & _
		"</td></tr>")
	result1.movenext()
Wend

%>
</tbody>
</table>
<br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {      
		sort: true,
		sort_config: {
			sort_types:['Number','String','String','String','Number']
		},	
		col_0: "select",
		col_2: "select",
		display_all_text: " [ Show all ] ",  
		sort_select: true,
		alternate_rows: true,
		loader: true,  
		loader_html: '<img src="./images/loader.gif" alt="" ' +
							'style="vertical-align:middle; margin:0 5px 0 5px" />' +
							'<span>Loading...</span>',
		loader_css_class: 'myLoader'
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script>  
<!--#include file="footer.asp"-->