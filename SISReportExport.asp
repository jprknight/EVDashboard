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
Response.AddHeader "Content-Disposition", "attachment; filename=EV-SISReport-" & DateTime & ".xls"

%>
<CENTER>
<TABLE cellspacing="0" cellpadding="1" rules="cols" border="1" bordercolor="black" id="MyDataGrid" style="border-color:Black;border-width:1px;border-style:solid;font-family:Verdana;font-size:8pt;border-collapse:collapse;">

<tr>
	<th nowrap align="left">Archive Name</th>
	<th nowrap align="left">EV Database</th>
	<th nowrap align="left">Items Shared</th>
	<th nowrap align="left">Archived Items</th>
	<th nowrap align="left">Percentage Shared</th>
	<th nowrap align="left">Archive Size (MB)</th>
</tr>

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

strSQLQuery1 = ("Select ArchiveName,EVDatabase,ItemsShared,ItemsArchived,ArchiveSizeMB from sisreport " & _
		"WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'")
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF

	Response.Write("<tr><td>" & (result1("ArchiveName")) & _
		"</td><td>" & (result1("EVDatabase")) & _
		"</td><td>" & (result1("ItemsShared")) & _
		"</td><td>" & (result1("ItemsArchived")))
	
	
	if (result1("ItemsShared")) > 0 AND (result1("ItemsArchived")) > 0 then
		strPercentageShared = (result1("ItemsShared")) / (result1("ItemsArchived")) * 100
	else
		strPercentageShared = "0"
	end if

	Response.Write("</td><td>" & LEFT(strPercentageShared,5) & _
		"</td><td>" & (result1("ArchiveSizeMB")) & "</td></tr>")
	result1.movenext()
Wend

%>
</table>
</center>