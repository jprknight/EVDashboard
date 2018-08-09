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
	<td BGCOLOR="#C0C0C0">PST FileName</td>
	<td BGCOLOR="#C0C0C0">Mailbox Name</td>
	<td BGCOLOR="#C0C0C0">Size (KB)</td>
	<td BGCOLOR="#C0C0C0">Current Status</td>
	<td BGCOLOR="#C0C0C0">Total Items</td>
	<td BGCOLOR="#C0C0C0">Items Archived</td>
	<td BGCOLOR="#C0C0C0">Items Stored</td>
	<td BGCOLOR="#C0C0C0">Start Time</td>
	<td BGCOLOR="#C0C0C0">Completion Time</td>
</tr>
<%
'Retrieve data from EnterpriseVaultDirectory database.
strDatabaseName = "EnterpriseVaultDirectory"
		
If ConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
		strDatabase & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
Elseif ConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
		strDatabase & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

' ********************************************************************
' * Revised to include items not previously attempted (or in progress)
' ********************************************************************

strSQLQuery1 = _
	"SELECT PST.FileSpecification as PSTPath, " & _
		"PST.FileSize as PSTSize, " & _
		"EME.MbxDisplayName AS MailboxName, " & _
		"PST.MigrationStatus as LastStatus, " & _
		"PSTMH.MigratedMsgCount as ImportedItems, " & _
		"PSTMH.MigratedMsgCount-PSTMH.NotEligibleForArchMsgCount as ArchivedItems, " & _
		"PSTMH.NotEligibleForArchMsgCount as StoredItems, " & _
		"PSTMH.MigrationStartTime as StartTime, " & _
		"PSTMH.MigrationEndTime as EndTime " & _
	"FROM EnterpriseVaultDirectory.dbo.PstFile PST " & _
	    "INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId " & _
		"INNER JOIN EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry EME ON PST.ExchangeMailboxEntryId = EME.ExchangeMailboxEntryId " & _
		"INNER JOIN ( " & _
		"    SELECT PST.FileSpecification as PSTF, Max(PSTMH.MigrationEndTime) as PSTT " & _
		"    FROM EnterpriseVaultDirectory.dbo.PstFile PST " & _
		"         INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId " & _
		"    GROUP BY PST.FileSpecification " & _
		"         ) AS PSTMaxTime ON PST.FileSpecification = PSTMaxTime.PSTF AND PSTMH.MigrationEndTime = PSTMaxTime.PSTT " & _
	SQLWhereClause & _
	" ORDER BY PST.FileSpecification"

if ProcessDate = Today then
	strSQLQuery1 = _
		"SELECT PSTPath, PSTSize, MailboxName, LastStatus, ImportedItems, ArchivedItems, StoredItems, StartTime, EndTime " & _
		"FROM " & _
		"( " & _
			"SELECT  PST.FileSpecification as PSTPath,  " & _
			"		PST.FileSize as PSTSize,  " & _
			"		EME.MbxDisplayName AS MailboxName,  " & _
			"		PST.MigrationStatus as LastStatus,  " & _
			"		PSTMH.MigratedMsgCount as ImportedItems,  " & _
			"		PSTMH.MigratedMsgCount-PSTMH.NotEligibleForArchMsgCount as ArchivedItems,  " & _
			"		PSTMH.NotEligibleForArchMsgCount as StoredItems,  " & _
			"		PSTMH.MigrationStartTime as StartTime,  " & _
			"		PSTMH.MigrationEndTime as EndTime  " & _
			"FROM EnterpriseVaultDirectory.dbo.PstFile PST  " & _
			"		INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId  " & _
			"		INNER JOIN EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry EME ON PST.ExchangeMailboxEntryId = EME.ExchangeMailboxEntryId  " & _
			"		INNER JOIN (  " & _
			"			SELECT PST.FileSpecification as PSTF, Max(PSTMH.MigrationEndTime) as PSTT  " & _
			"			FROM EnterpriseVaultDirectory.dbo.PstFile PST  " & _
			"				INNER JOIN EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH ON PSTMH.PstFileEntryId = PST.PstFileEntryId  " & _
			"			GROUP BY PST.FileSpecification  " & _
			"			) AS PSTMaxTime ON PST.FileSpecification = PSTMaxTime.PSTF AND PSTMH.MigrationEndTime = PSTMaxTime.PSTT  " & _
			SQLWhereClause & _
			"UNION ALL " & _
			"SELECT  PST.FileSpecification as PSTPath,  " & _
			"		PST.FileSize as PSTSize,  " & _
			"		EME.MbxDisplayName AS MailboxName,  " & _
			"		PST.MigrationStatus as JobResult, " & _
			"		0 as ImportedItems, " & _
			"		0 as ArchivedItems, " & _
			"		0 as StoredItems, " & _
			"		PST.LastAccessedTime as StartTime, " & _
			"		PST.LastCopiedTime as EndTime " & _
			"FROM EnterpriseVaultDirectory.dbo.PstFile PST  " & _
			"		INNER JOIN EnterpriseVaultDirectory.dbo.ExchangeMailboxEntry EME ON PST.ExchangeMailboxEntryId = EME.ExchangeMailboxEntryId  " & _
			"WHERE NOT PST.PstFileEntryId IN (SELECT DISTINCT PSTMH.PstFileEntryId FROM EnterpriseVaultDirectory.dbo.PstMigrationHistory PSTMH) " & _
			"  AND PST.MigrationStatus <> 100 " & _
		") Data " & _
		"ORDER BY PSTPath "
End If

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	Response.Write("<tr><td>" & result1("PSTPath") & _
		"</td><td>" & (result1("MailboxName")) & _
		"</td><td align=""right"">" & (FormatNumber(result1("PSTSize"),0)) & _
		"</td><td>")
	Select Case result1("LastStatus")
		Case "0"		: Response.Write "Not Ready"
		Case "100"		: Response.Write "Do Not Migrate"
		Case "200"		: Response.Write "Ready to Copy"
		Case "300"		: Response.Write "Copying"
		Case "400"		: Response.Write "Copy Failed"
		Case "500"		: Response.Write "Ready to Backup"
		Case "600"		: Response.Write "Backing Up"
		Case "700"		: Response.Write "Backup Failed"
		Case "800"		: Response.Write "Ready to Migrate"
		Case "900"		: Response.Write "Migrating"
		Case "1000"		: Response.Write "Failed"
		Case "1100"		: Response.Write "Ready for Post Processing"
		Case "1200"		: Response.Write "Completing"
		Case "1300"		: Response.Write "Completion Failed"
		Case "1400"		: Response.Write "Complete"
		Case Else		: Response.Write result1("LastStatus")
	End Select

	Response.Write(	"</td><td align=""right"">" & (FormatNumber(result1("ImportedItems"),0)) & _
		"</td><td align=""right"">" & (FormatNumber(result1("ArchivedItems"),0)) & _
		"</td><td align=""right"">" & (FormatNumber(result1("StoredItems"),0)) & _
		"</td><td>" & (result1("StartTime")) &_
		"</td><td>" & (result1("EndTime")) & "</td></tr>" & vbCrLf)
	result1.movenext()
Wend

%>
</table><br>