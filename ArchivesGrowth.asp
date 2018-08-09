<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->

<h1>EV Archives Growth</h1><br>

<table class="Chart">
	<tr> 
    <td valign="top" class="text" align="left">
	<%
	Call renderChartHTML("./includes/FusionCharts/FCF_Column3D.swf", "", strXML, "myNext", 600, 300)
	%>
	</td>
	</tr>
</table><br>

<a href="ArchivesGrowthExport.asp" title="Download data as Excel Spreadsheet"><img border="0" src="./images/page_excel.png"> Download</a>
<br><br>

<table id="table1" class="mytable">
<thead>
<tr>
	<th nowrap align="left">Date</th>
	<th nowrap align="left">Count</th>
</tr>
</thead>
<tbody>
<%

response.Buffer=true

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
	Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	Connection2 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* Open database connection
Set myconn2 = CreateObject("adodb.connection")
myconn2.ConnectionTimeout = strSQLConnect
myconn2.open(Connection2)

strSQLQuery2 = "SELECT CONVERT(nvarchar(10), RecordCreateTimestamp, 120) AS Date, count(*) AS Count " & _
	"FROM EnabledMailboxes " & _
	"GROUP BY CONVERT(nvarchar(10), RecordCreateTimestamp, 120) " & _
	"ORDER BY CONVERT(nvarchar(10), RecordCreateTimestamp, 120) DESC"

Set objCmd2 = CreateObject("adodb.command")
objCmd2.activeconnection = myconn2
objCmd2.commandtimeout = strSQLExecute
objCmd2.commandtype = adCmdText
objCmd2.commandtext = strSQLQuery2

Set result2 = CreateObject("adodb.Recordset")
result2.open objCmd2

While not result2.EOF
	Response.Write ("<tr><td>" & (result2("Date")) & _
		"</td><td>" & (result2("Count")) & _
		"</td></tr>")
	result2.movenext()
Wend

%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number','Number','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script>  
<!--#include file="footer.asp"-->