<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br>
<h1>EV Dashboard Database Statistics</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th align="left">Table</th>
<th align="left">Oldest</th>
<th align="left">Newest</th>
<th align="left">Count</th>
</tr>
</thead>
<tbody>
<%

DBTables = Array("ArchiveStates", _
	"ArchivingActivity", _
	"DisabledUsers", _
	"DiskUsageInfo", _
	"EnabledUsers", _
	"EnabledMailboxes",_
	"ExchangeMailboxes", _
	"ExchangeTransactionLogs", _
	"HourlyArchivingRates", _
	"MSMQMessageCountAlerts", _
	"OrganisationArchivingHistory", _
	"OverReceiveLimit", _
	"OverSendLimit", _
	"OverWarningLimit", _
	"ServiceAlerts",_
	"UsageReportByArchive")
	
For each Table in DBTables

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

	strSQLQuery1 = "select Convert(nvarchar(30),Min(RecordCreateTimestamp), 120) AS Oldest," & _
		"Convert(nvarchar(30),Max(RecordCreateTimestamp), 120) AS Newest, Count(*) AS Count from " & Table

	Set objCmd1 = CreateObject("adodb.command")
	objCmd1.activeconnection = myconn1
	objCmd1.commandtimeout = strSQLExecute
	objCmd1.commandtype = adCmdText
	objCmd1.commandtext = strSQLQuery1

	Set result1 = CreateObject("adodb.Recordset")
	result1.open objCmd1

	While not result1.EOF
		Response.Write("<tr><td>" & Table & _
					"</td><td>" & (result1("Oldest")) & _
					"</td><td>" & (result1("Newest")) & _
					"</td><td>" & (result1("Count")) & "</td></tr>")
		result1.movenext()
	Wend
	
	EVDmyconn.close
Next
%>
</tbody>
</table>
<br>
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