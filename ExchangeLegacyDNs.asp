<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<form id="frmconfig" method="post" action="ExchangeLegacyDNs.asp">
<h1><img src="images/blockdevice32x32.png"> Exchange Legacy DNs</h1>
<br><br>
<h3>Enter part or all of a user displayname and press submit.</h3><br>
<input id="txtDisplayName" name="txtDisplayName" type="text"/>
<span><input type="submit" id="btnSubmit" name="btnSubmit" value="Submit"/></button></span>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th>Displayname</th>
<th>Exchange Server</th>
<th>Mbx Alias</th>
<th>Legacy DN</th>
</tr>
<thead>
<tbody>
<%
If Request.Form("btnSubmit") = "Submit" Then
	UserDisplayName = request.form("txtDisplayName")
	'* Included if statement. I found that if UserDisplayName is null and the submit button is pressed, all records are returned.
	'* This can be slow to run on larger installations (retrieving all records from the ExchangeMailboxEntry table.
	if UserDisplayName <> "" then
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

		strSQLQuery1 = "SELECT ExchangeMailboxEntry.MbxDisplayname, " & _
			"ExchangeMailboxEntry.MbxAlias, " & _
			"ExchangeServerEntry.ExchangeComputer, " & _
			"ExchangeMailboxEntry.LegacyMbxDN " & _
			"FROM ExchangeMailboxEntry " & _
			"join ExchangeServerEntry " & _
			"on ExchangeMailboxEntry.ExchangeServerIdentity = ExchangeServerEntry.ExchangeServerIdentity " & _
			"WHERE ExchangeMailboxEntry.MbxDisplayname like '%" & UserDisplayName & "%'"

		Set objCmd1 = CreateObject("adodb.command")
		objCmd1.activeconnection = myconn1
		objCmd1.commandtimeout = strSQLExecute
		objCmd1.commandtype = adCmdText
		objCmd1.commandtext = strSQLQuery1

		Set result1 = CreateObject("adodb.Recordset")
		result1.open objCmd1

		While not result1.EOF
			response.write "<tr><td>" & (result1("MbxDisplayName")) & _
				"</td><td>" & (result1("ExchangeComputer")) & _
				"</td><td>" & (result1("MbxAlias")) & _
				"</td><td>" & (result1("LegacyMbxDN")) & _
				"</td></tr>"
			result1.movenext()
		Wend
	End if
End if
%>
</tbody>
</table><br>
</form>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		sort: true,
		col_1: "select",    
		display_all_text: " [ Show all ] ",  
		sort_config: {
			sort_types:['String','String','String','String']
		},
		alternate_rows: true,
		loader: true,  
		loader_html: '<img src="./images/loader.gif" alt="" ' +
							'style="vertical-align:middle; margin:0 5px 0 5px" />' +
							'<span>Loading...</span>',
		loader_css_class: 'myLoader',
		alternate_rows: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->