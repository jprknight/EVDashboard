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
Response.AddHeader "Content-Disposition", "attachment; filename=EV-DisabledUsers-" & DateTime & ".xls"
%>
<CENTER>
<TABLE cellspacing="0" cellpadding="1" rules="cols" border="1" bordercolor="black" id="MyDataGrid" style="border-color:Black;border-width:1px;border-style:solid;font-family:Verdana;font-size:8pt;border-collapse:collapse;">

<tr>
	<td BGCOLOR="#C0C0C0">Mailbox</td>
	<td BGCOLOR="#C0C0C0">Warning (MB)</td>
	<td BGCOLOR="#C0C0C0">Send (MB)</td>
	<td BGCOLOR="#C0C0C0">Receive (MB)</td>
	<td BGCOLOR="#C0C0C0">Mbx Size (MB)</td>
	<td BGCOLOR="#C0C0C0">#Items (Mailbox)</td>
</tr>
<%
'Retrieve data from EVDashboard database.
If EVDConnectionString = 1 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* Open database connection
Set EVDmyconn = CreateObject("adodb.connection")
EVDmyconn.open (EVDconnection)

If cstr(err.number) <> 0 Then
	wscript.echo (" - Error creating connection to " _
   & "database server " & EVDSQLServer & " / " & EVDSQLDatabase & ".  Check your connection string " _
   & "or database server name/IP and try again.")
   wscript.quit
End if

Set result = CreateObject("adodb.recordset")
If err.number <> 0 then msgbox err.description

'* Execute the query
Set result = EVDmyconn.execute("Select Mailbox,WarningMB,SendMB,ReceiveMB,MailboxSizeMB,NumItemsMailbox From DisabledUsers " & _	
	"WHERE RecordCreateTimestamp > '" & styr & "/" & stmo & "/" & stdt & " 00:00:00'")

If err.number <> 0 then msgbox err.description

While not result.EOF
	Response.Write ("<tr><td>" & (result("Mailbox")) & _
		"</td><td>" & (result("WarningMB")) & _
		"</td><td>" & (result("SendMB")) & _
		"</td><td>" & (result("ReceiveMB")) & _
		"</td><td>" & (result("MailboxSizeMB")) & _
		"</td><td>" & (result("NumItemsMailbox")) & _
		"</td></tr>")
	result.movenext()
Wend
%>
</table>
</center>