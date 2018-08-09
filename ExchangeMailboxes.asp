<!--#include file="./header.asp"-->
<!--#include file="./global.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<!--#include file="./includes/FC_Colors.asp"-->
<!--#include file="./includes/FusionCharts.asp" -->
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

OrgTotalMailboxes = 0

Styr = Year(Now)
Stmo = Month(Now)
if Stmo < 10 Then
	Stmo = "0" & Stmo
end if
Stdt = Day(Now)
if Stdt < 10 Then
	Stdt = "0" & Stdt
end if

strXML = "<graph xAxisName='Exchange Server' yAxisName='Number of Mailboxes' rotateNames='1' decimalPrecision='0' formatNumberScale='0' bgColor='EAEAEA'>"

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

strSQLQuery1 = "select ExchangeServer, " & _
	"sum(MbxCount) AS MbxCount " & _
	"from ExchangeMailboxes " & _
	SQLWhereClause & _
	"Group By ExchangeServer " & _
	"Order by ExchangeServer"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF		
	strXML = strXML & "<set name='" & (result1("ExchangeServer")) & "' value='" & (result1("MbxCount")) & "' color='" & getFCColor() & "'/>"
	OrgTotalMailboxes = OrgTotalMailboxes + (result1("MbxCount"))
	result1.movenext()
Wend

strXML = strXML & "</graph>"

%>
<!-- Main Form -->
<form id="frmconfig" method="post" action="ExchangeMailboxes.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td>
<h1><img src="images/users.png"> Exchange Mailboxes (<%=OrgTotalMailboxes%>)</h1>
</td>
<td align="right">
<input id="date" name="date" type="text" value="<%=ProcessDate%>" class="dateselector" />
<span style="text-align:left;width:100%;"><input type="submit" id="btnRefreshData" name="btnRefreshData" value="Refresh Data" class="pgInp" /></button></span>
<span style="text-align:left;width:100%;"><input type="submit" id="btnToday" name="btnToday" value="Today" class="pgInp" /></button></span><br>
</td>
</tr>
</table>
<span>
<table class="Main">
<tr> 
<td>
<%
Call renderChartHTML("./includes/FusionCharts/FCF_Column3D.swf", "", strXML, "myNext", 800, 400)
%>
</td>
</tr>
</table>
</span>
<br>
</form>
<!--#include file="footer.asp"-->