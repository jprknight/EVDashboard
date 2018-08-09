<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<br><br>
<h1>Exchange Mailbox States (Live Data)</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th># Mailboxes</th>
<th>Exchange State</th>
</tr>
<thead>
<tbody>
<%

strDatabaseName = "EnterpriseVaultDirectory"

	strSQLPath = (Server.MapPath("./sql/"))
	strSQLFile = strSQLPath & "\MailboxArchiveState.sql"

	Set SQLQueryFile = fso.OpenTextFile(strSQLFile, ForReading)
	SQLQuery = SQLQueryFile.ReadAll

	If ConnectionString = 1 Then
		connection = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
			strDatabaseName & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
	Elseif ConnectionString = 2 Then
		connection = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
			strDatabaseName & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
	End if
			
	'* Open database connection
	Set myconn = CreateObject("adodb.connection")
	myconn.open (connection)

	If cstr(err.number) <> 0 Then
		wscript.echo (" - Error creating connection to " _
	   & "database server " & DirectoryServer & " / " & strDatabaseName & ".  Check your connection string " _
	   & "or database server name/IP and try again.")
	   wscript.quit
	End if
		
	On Error Goto 0
		
	Set result = CreateObject("adodb.recordset")
	If err.number <> 0 then msgbox err.description

	'* Execute the query
	Set result = myconn.execute(SQLQuery)
	If err.number <> 0 then msgbox err.description

	While not result.EOF
		Select Case (result("Exchange State"))
			Case 0
				ExchangeState = "Normal"
			Case 1
				ExchangeState = "Hidden"
			Case 2
				ExchangeState = "Deleted"
		End Select
	
		Response.Write("<tr><td>" & (result("# Mailboxes")) & "</td><td>" & ExchangeState & "</td></tr>")
		result.movenext()
	Wend

%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		sort: true,
		sort_config: {
			sort_types:['Number','String']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script>
<!--#include file="footer.asp"-->