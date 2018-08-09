<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br><br>
<h1><img src="images/server32x32.png"> File Server Archive Configuration</h1>
<br><br>
<table id="table1" class="mytable" align="center">
<thead>
<tr>
<th>EV Server</th>
<th>Vault Store</th>
<th>File Server</th>
<th>Volume</th>
<th>Path</th>
<th>Type</th>
<th>Archive Folder</th>
<th>Archive Children</th>
<th>Volume Policy</th>
<th>Folder Policy</th>
</tr>
</thead>
<tbody>
<%
'on error resume next

HTML_SPAN_GREEN = "<span style=""color: #006600; font-weight: bold;"">"
HTML_SPAN_GREY  = "<span style=""color: #666666; font-weight: bold; font-style: italic;"">"
HTML_SPAN_RED   = "<span style=""color: #ff3333; font-weight: bold;"">"
HTML_SPAN_END   = "</span>"

adCmdText = 1
strSQLExecute = 10

'Retrieve data from EnterpriseVaultDirectory database.
strDatabase = "EnterpriseVaultDirectory"

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

If Err <> 0 Then
	Response.Write "<!-- After Create Connection -->" & vbCrLf
	Response.Write "<!-- " & strSQLConnect & " -->" & vbCrLf
	Response.Write "<!-- " & Err.Number & ": " & Err.Description & " -->" & vbCrLf
	Response.Write "<!-- " & Connection1 & " -->" & vbCrLf
	Response.Flush
	Response.End
End If

strSQLQuery = "SELECT " & vbCrLf & _
"	LEFT(CE.ComputerName, CHARINDEX('.', CE.ComputerName)-1) as [EVServer], " & vbCrLf & _
"	VSE.VaultStoreName as [Vault], " & vbCrLf & _
"	FSE.DnsName as [FileServer], " & vbCrLf & _
"	FSVE.VolumeName as [Volume], " & vbCrLf & _
"	FSFE.FolderPath as [FolderPath], " & vbCrLf & _
"	FSFE.ArchiveThisFolder as [ArchiveFolder], " & vbCrLf & _
"	FSFE.ArchiveSubFolders as [ArchiveChildren], " & vbCrLf & _
"	PE1.poName as [VolumePolicy], " & vbCrLf & _
"	PE2.poName as [FolderPolicy] " & vbCrLf & _
"FROM FileServerEntry FSE " & vbCrLf & _
"	INNER JOIN FileServerVolumeEntry FSVE on FSE.FileServerEntryId = FSVE.FileServerEntryId  " & vbCrLf & _
"	INNER JOIN FileServerFolderEntry FSFE on FSVE.VolumeEntryId = FSFE.VolumeEntryId  " & vbCrLf & _
"	INNER JOIN FileServerVolumeArchiveEntry FSVAE on FSVE.VolumeEntryId = FSVAE.VolumeEntryId  " & vbCrLf & _
"	INNER JOIN VolumePolicyEntry VPE on FSVE.VolumePolicyEntryId = VPE.VolumePolicyEntryId " & vbCrLf & _
"	INNER JOIN FolderPolicyEntry FPE on FPE.FolderPolicyEntryId = FSFE.FolderPolicyEntryId  " & vbCrLf & _
"	INNER JOIN PolicyEntry PE1 on VPE.poPolicyEntryId = PE1.poPolicyEntryId  " & vbCrLf & _
"	INNER JOIN PolicyEntry PE2 on FPE.poPolicyEntryId = PE2.poPolicyEntryId  " & vbCrLf & _
"	INNER JOIN VaultStoreEntry VSE on VSE.VaultStoreEntryId = FSVE.VaultStoreEntryId  " & vbCrLf & _
"	INNER JOIN StorageServiceEntry SSE on VSE.StorageServiceEntryId = SSE.ServiceEntryId   " & vbCrLf & _
"	INNER JOIN ComputerEntry CE on SSE.ComputerEntryId = CE.ComputerEntryId " & vbCrLf & _
"ORDER BY [EVServer], [Vault], [FileServer], [Volume], [FolderPath]"

Response.Write vbCrLf & vbCrLf & "<!-- " & strSQLQuery & " -->" & vbCrLf & vbCrLf
Response.Flush

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery

If Err <> 0 Then
	Response.Write "<!-- After Create Command -->" & vbCrLf
	Response.Write "<!-- " & objCmd1.commandtimeout & " -->" & vbCrLf
	Response.Write "<!-- " & objCmd1.commandtype & " -->" & vbCrLf
	Response.Write "<!-- " & strSQLQuery & " -->" & vbCrLf
	Response.Write "<!-- " & Err.Number & ": " & Err.Description & " -->" & vbCrLf
	Response.Write "<!-- " & Connection1 & " -->" & vbCrLf
	Response.Flush
	Response.End
End If

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

If Err <> 0 Then
	Response.Write "<!-- After Create RecordSet -->" & vbCrLf
	Response.Write "<!-- " & strSQLExecute & " -->" & vbCrLf
	Response.Write "<!-- " & adCmdText & " -->" & vbCrLf
	Response.Write "<!-- " & strSQLQuery & " -->" & vbCrLf
	Response.Write "<!-- " & Err.Number & ": " & Err.Description & " -->" & vbCrLf
	Response.Write "<!-- " & Connection1 & " -->" & vbCrLf
	Response.Flush
	Response.End
End If

While not result1.EOF

	Response.Write	"<tr><td>" & result1("EVServer") & _
			"</td><td>" & result1("Vault") & _
			"</td><td>" & result1("FileServer") & _
			"</td><td>" & result1("Volume") & _
			"</td><td>" & result1("FolderPath")
	FullPath = "\\" & result1("FileServer") & "\" & result1("Volume") & "\" & result1("FolderPath")

	Select Case ADSType(FullPath)
		Case 1: Response.Write "</td><td>" & HTML_SPAN_GREEN & "Archive Point" & HTML_SPAN_END
		Case 2: Response.Write "</td><td>" & HTML_SPAN_GREY & "Folder" & HTML_SPAN_END
		Case Else: Response.Write "</td><td>" & HTML_SPAN_RED & "Unknown" & HTML_SPAN_END
	End Select

	If result1("ArchiveFolder") = "1" Then
		Response.Write "</td><td>" & HTML_SPAN_GREEN & "Enabled" & HTML_SPAN_END
	Else
		Response.Write "</td><td>" & HTML_SPAN_RED & "Disabled" & HTML_SPAN_END
	End If

	If result1("ArchiveChildren") = "1" Then
		Response.Write "</td><td>" & HTML_SPAN_GREEN & "Enabled" & HTML_SPAN_END
	Else
		Response.Write "</td><td>" & HTML_SPAN_RED & "Disabled" & HTML_SPAN_END
	End If

	Response.Write	"</td><td>" & (result1("VolumePolicy")) & _
			"</td><td>" & (result1("FolderPolicy")) & "</td></tr>" & vbCrLf
	result1.movenext()
	Response.Flush
Wend

function ADSType (sFullPath)
	Set fso = CreateObject ("Scripting.FileSystemObject")
	If fso.FileExists(sFullPath & ":EVArchivePoint.xml") Then
		ADSType = 1
	ElseIf fso.FileExists(sFullPath & ":EVFolderPoint.xml") Then
		ADSType = 2
	Else
		ADSType = -1
	End If
end function

%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		filters_row_index: 1,
		sort: true,
		sort_config: {
			sort_types:['String','String','String','String','String','String','String','String','String','String']
		},
		col_0: "select",
		col_1: "select",
		col_2: "select",
		col_3: "select",
		col_5: "select",
		col_6: "select",
		col_7: "select",
		col_8: "select",
		col_9: "select",
		display_all_text: " [ Show all ] ",  
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->
