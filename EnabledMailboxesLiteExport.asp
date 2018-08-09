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
Response.AddHeader "Content-Disposition", "attachment; filename=EV-EnabledMailboxes-" & DateTime & ".xls"
%>
<CENTER>
<TABLE cellspacing="0" cellpadding="1" rules="cols" border="1" bordercolor="black" id="MyDataGrid" style="border-color:Black;border-width:1px;border-style:solid;font-family:Verdana;font-size:8pt;border-collapse:collapse;">

<tr>
	<td BGCOLOR="#C0C0C0">Mailbox</td>
	<td BGCOLOR="#C0C0C0">EV Server</td>
	<td BGCOLOR="#C0C0C0">Num Items Mailbox</td>
	<td BGCOLOR="#C0C0C0">Mailbox Size (MB)</td>
	<td BGCOLOR="#C0C0C0">Num Items Archive</td>
	<td BGCOLOR="#C0C0C0">Archive Size (MB)</td>
	<td BGCOLOR="#C0C0C0">Total Size (MB)</td>
	<td BGCOLOR="#C0C0C0">Archive Created</td>
	<td BGCOLOR="#C0C0C0">Archive Updated</td>
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

On Error Goto 0
	
Set result = CreateObject("adodb.recordset")
If err.number <> 0 then msgbox err.description

'* Execute the query
Set result = EVDmyconn.execute("Select Mailbox,ExchangeServer," & _
	"EVServer,DBServer,EVDatabase,NumItemsMailbox,MailboxSizeMB," & _
	"NumItemsArchive,ArchiveSizeMB,TotalSizeMB,ArchiveCreated,ArchiveUpdated " & _
	"From EnabledMailboxes " & _
	"WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'")

If err.number <> 0 then msgbox err.description

While not result.EOF
	Response.Write ("<tr><td>" & (result("Mailbox")) & _
		"</td><td>" & (result("EVServer")) & _
		"</td><td>" & (result("NumItemsMailbox")) & _
		"</td><td>" & (result("MailboxSizeMB")) & _
		"</td><td>" & (result("NumItemsArchive")) & _
		"</td><td>" & (result("ArchiveSizeMB")) & _
		"</td><td>" & (result("TotalSizeMB")) & _
		"</td><td>" & (result("ArchiveCreated")) & _
		"</td><td>" & (result("ArchiveUpdated")) & "</td></tr>")
	result.movenext()
Wend

%>
</table>
</center>