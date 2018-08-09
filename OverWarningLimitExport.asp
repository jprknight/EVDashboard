<!--#include file="./config.asp"-->
<%
Styr = Year(Now)
Stmo = Month(Now)
if Stmo < 10 Then
	Stmo = "0" & Stmo
end if
MonthFN = MonthName(Stmo, False)
Stdt = Day(Now)
StdtEmail = Day(Now)
if Stdt < 10 Then
	Stdt = "0" & Stdt
end if
DayNo = WeekDay(Now)
DayName = WeekDayName(DayNo,False)
Sthr = Hour(Now)
if Sthr < 10 Then
	Sthr = "0" & Sthr
end if
Stmin = Minute(Now)
if Stmin < 10 Then
	Stmin = "0" & Stmin
end if
Stsec = Second(Now)
if Stsec < 10 Then
	Stsec = "0" & Stsec
end if
DateTime = styr & "." & stmo & "." & stdt & "-" & sthr & "." & stmin & "." & stsec

Response.Buffer = False
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=EV-OverWarningLimit-" & DateTime & ".xls"
%>
<CENTER>
<TABLE cellspacing="0" cellpadding="1" rules="cols" border="1" bordercolor="black" id="MyDataGrid" style="border-color:Black;border-width:1px;border-style:solid;font-family:Verdana;font-size:8pt;border-collapse:collapse;">

<tr>
	<td BGCOLOR="#C0C0C0">Mailbox</td>
	<td BGCOLOR="#C0C0C0">Over Warning (MB)</td>
	<td BGCOLOR="#C0C0C0">Warning (MB)</td>
	<td BGCOLOR="#C0C0C0">Send (MB)</td>
	<td BGCOLOR="#C0C0C0">Receive (MB)</td>
	<td BGCOLOR="#C0C0C0">Mbx Size (MB)</td>
	<td BGCOLOR="#C0C0C0">#Items (Mailbox)</td>
	<td BGCOLOR="#C0C0C0">Exchange Server</td>
	<td BGCOLOR="#C0C0C0">Last Archived</td>
	<td BGCOLOR="#C0C0C0">Exchange State</td>
	<td BGCOLOR="#C0C0C0">Policy Used</td>
</tr>

<%

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

strSQLQuery1 = "Select Mailbox,OverWarningMB,WarningMB,SendMB,ReceiveMB," & _
	"MailboxSizeMB,NumItemsMailbox,ExchangeServer,LastArchived,ExchangeState,PolicyUsed from OverWarningLimit " & _
	"WHERE RecordCreateTimestamp > '" & styr & "/" & stmo & "/" & stdt & " 00:00:00'"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF		
	Response.Write("<tr><td>" & (result1("Mailbox")) & "</td><td>" & (result1("OverWarningMB")) & _
	"</td><td>" & (result1("WarningMB")) & "</td><td>" & (result1("SendMB")) & _
	"</td><td>" & (result1("ReceiveMB")) & "</td><td>" & (result1("MailboxSizeMB")) & _
	"</td><td>" & (result1("NumItemsMailbox")) & "</td><td>" & (result1("ExchangeServer")) & _
	"</td><td>" & (result1("LastArchived")) & "</td><td>" & (result1("ExchangeState")) & _
	"</td><td>" & (result1("PolicyUsed")) & "</td></tr>")
	result1.movenext()
Wend
%>
</table><br>