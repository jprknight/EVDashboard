<span>
<h1>User Indexes (Live Data)</h1><br><br>
<table id="table4" class="insidetable" style="width:300px;">
<thead>
<tr>
	<th nowrap width="80px" align="left">Name</th>
	<th nowrap width="80px" align="left">Status</th>
</tr>
</thead>
<tbody>
<%

Dim LineCount: LineCount = 0

strDatabaseName = "EnterpriseVaultDirectory"

If ConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
		strDatabaseName & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif ConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
		strDatabaseName & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if
		
'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "Use EnterpriseVaultDirectory Select RootIdentity,Failed,Offline,Rebuilding " & _
	"from IndexVolume where failed <> 0 or Offline <> 0 or Rebuilding <> 0"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	FailedUserIndexesData = FailedUserIndexesData & (result1("RootIdentity")) & "-" & _
		(result1("Failed")) & "+" & (result1("Offline")) & "*" & (result1("Rebuilding")) & vbcrlf	
	result1.movenext()
Wend

FailedUserIndexesDataArray = Split(FailedUserIndexesData,vbcrlf)

For each Line in FailedUserIndexesDataArray
	if Len(Line) > 0 then
		
		LineCount = LineCount + 1
		
		Q1 = Instr(Line,"-")
		RootIdentity = mid(Line,1,Q1 - 1)
	
		Q2 = Instr(Line,"+")
		Q3 = Instr(Line,"*")
		
		StatusCount = 0
		
		'* Failed
		FailedLength = Q2 - Q1 - 1
		Failed = mid(Line,Q1 + 1,FailedLength)
		'msgbox "Failed = " & Failed
		If Failed = "True" then
			StatusCount = StatusCount + 1
			Failed = vbcrlf & "Failed"
		Else
			Failed = ""
		End if
		
		'*Offline
		OfflineLength = Q3 - Q2 - 1
		Offline = mid(Line,Q2 + 1,OfflineLength)
		'msgbox "Offline = " & Offline
		If Offline = "True" then
			StatusCount = StatusCount + 1
			If StatusCount > 1 then
				Offline = vbcrlf & ", Offline"
			Else
				Offline = "Offline"
			End if
		Else
			Offline = ""
		End if
		
		'* Rebuilding
		Rebuilding = mid(Line,Q3 + 1)
		'msgbox "Rebuilding = " & rebuilding
		If Rebuilding = "True" then
			StatusCount = StatusCount + 1
			If StatusCount > 1 then
				Rebuilding = ", Rebuilding"
			Else
				Rebuilding = vbcrlf & "Rebuilding"
			End if
		Else
			Rebuilding = ""
		End if
		
		'* Open database connection
		Set myconn2 = CreateObject("adodb.connection")
		myconn2.ConnectionTimeout = strSQLConnect
		myconn2.open(Connection1)

		strSQLQuery2 = "Use EnterpriseVaultDirectory Select MbxDisplayName from ExchangeMailboxEntry " & _
			"Inner Join Root On Root.VaultEntryId=ExchangemailboxEntry.DefaultVaultid Where Root.RootIdentity = " & RootIdentity

		Set objCmd2 = CreateObject("adodb.command")
		objCmd2.activeconnection = myconn2
		objCmd2.commandtimeout = strSQLExecute
		objCmd2.commandtype = adCmdText
		objCmd2.commandtext = strSQLQuery2

		Set result2 = CreateObject("adodb.Recordset")
		result2.open objCmd2
		
		Response.Write ("<tr><td>" & result2("MbxDisplayName") & "</td><td>" & Failed & Offline & Rebuilding & "</td></tr>")
	End if
Next
If LineCount = 0 then
	Response.Write ("<tr><td>No indexes marked as</td><td>Failed, Offline or Rebuilding</td></tr></tbody></table>")
Else
	Response.Write ("</tbody></table>")
End if
%>
</span>