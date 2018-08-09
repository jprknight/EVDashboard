<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<h1><img src="images/blockdevice32x32.png"> Orphaned Archives</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th>Displayname</th>
<th>Company</th>
<th>State</th>
<th>Last Modified</th>
<th>Legacy DN</th>
</tr>
<thead>
<tbody>
<%
If ConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
Elseif ConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
End if
		
'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "Select MbxDisplayName, Company, State, LastModified, LegacyMbxDN " & _
	"from ExchangeMailboxEntry WHERE LegacyMbxDn LIKE '%marked%' " & _
	"AND LegacyMbxDn like '%@%'"

Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	response.write "<tr><td>" & (result1("MbxDisplayName")) & "</td><td>" & (result1("Company")) & _
		"</td><td>" & (result1("State")) & "</td><td>" & dbDate((result1("LastModified"))) & _
		"</td><td>" & (result1("LegacyMbxDN")) & "</td></tr>"
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
			sort_types:['String','String','String','String','String']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->