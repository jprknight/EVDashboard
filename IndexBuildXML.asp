<%
Styr = Year(Now)
Stmo = Month(Now)
if Stmo < 10 Then
	Stmo = "0" & Stmo
end if
Stdt = Day(Now)
if Stdt < 10 Then
	Stdt = "0" & Stdt
end if

Dim TotalUsers: TotalUsers = 0

strXML = "<graph xAxisName='Enterprise Vault Storage Database' yAxisName='Number of Archives' decimalPrecision='0' formatNumberScale='0' bgColor='EAEAEA' rotateNames='1'>"

If EVDConnectionString = 1 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* Required Const declaration in EVDashboard_Data_Collector.vbs
'Const adCmdText = 1

'* Open database connection
Set myconn = CreateObject("adodb.connection")
myconn.ConnectionTimeout = strSQLConnect
myconn.open(EVDconnection)

strSQLQuery = ("select EVDatabase, " & _
	"count(EVDatabase) AS Count " & _
	"from EnabledMailboxes " & _
	"Where RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00'" & _
	"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'" & _
	"Group By EVDatabase " & _
	"Order by EVDatabase")

Set objCmd = CreateObject("adodb.command")
objCmd.activeconnection = myconn
objCmd.commandtimeout = strSQLExecute
objCmd.commandtype = adCmdText
objCmd.commandtext = strSQLQuery

Set result = CreateObject("ADODB.Recordset")
result.open objCmd

While not result.EOF		
	strXML = strXML & "<set name='" & (result("EVDatabase")) & "' value='" & (result("Count")) & "' color='" & getFCColor() & "'/>"
	TotalUsers = TotalUsers + (result("Count"))
	result.movenext()
Wend

strXML = strXML & "</graph>"
%>