<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<h1>Archive States (Live Data)</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th align="left"># Mailboxes</th>
<th align="left">Archive State</th>
</tr>
</thead>
<tbody>
<%

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

strSQLQuery1 = "Use EnterpriseVaultDirectory " & _
	"SELECT count(MbxArchivingState) as '# Mailboxes', " & _
	"MbxArchivingState as 'Archive State' " & _
	"FROM ExchangeMailboxEntry " & _
	"GROUP BY MbxArchivingState"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	Select Case (result1("Archive State"))
		Case 0
			ArchiveState = "Not Enabled"
		Case 1
			ArchiveState = "Enabled"
		Case 2
			ArchiveState = "Disabled"
		Case 3
			ArchiveState = "Re-Link"
	End Select

	Response.Write("<tr><td>" & (result1("# Mailboxes")) & "</td><td>" & ArchiveState & "</td></tr>")
	result1.movenext()
Wend

%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number','Number','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->