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
Response.AddHeader "Content-Disposition", "attachment; filename=EV-EnabledUsers-" & DateTime & ".xls"
%>
<CENTER>
<TABLE cellspacing="0" cellpadding="1" rules="cols" border="1" bordercolor="black" id="MyDataGrid" style="border-color:Black;border-width:1px;border-style:solid;font-family:Verdana;font-size:8pt;border-collapse:collapse;">

<tr>
<td BGCOLOR="#C0C0C0">Date</td>
<td BGCOLOR="#C0C0C0">Count</td>
</tr>
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

	
Set result = EVDmyconn.execute("select CONVERT(nvarchar(10), RecordCreateTimestamp, 120) AS Date, count(*) AS Count from enabledusers " & _
	"group by CONVERT(nvarchar(10), RecordCreateTimestamp, 120) " & _
	"order by CONVERT(nvarchar(10), RecordCreateTimestamp, 120) DESC")

If err.number <> 0 then msgbox err.description

'************************************
'* Write data.
While not result.EOF
	Response.Write ("<tr><td>" & (result("Date")) & _
		"</td><td>" & (result("Count")) & _
		"</td></tr>")
	result.movenext()
Wend

%>
</tbody>
</table><br>