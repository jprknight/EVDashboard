<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<html>
<head>
<br><br>
<h1>Test Database Connectivity</h1><br>
<head>
<body>
<div class="wrapper" id="savediv" style="margin:auto">
  <div class="content-top">
  </div>
  <div class="content">
    <form id="frmconfig" method="post" action="TestDBConnectivity.asp">
		<table>
		<tr>
		<td>
		<span><h3>Database server:</h3></span>
		</td>
		<td>
        <span><input style="width:150px;border:1px dotted gray;padding-bottom:8px;" value=""
            title="The name or IP address of your SQL server." name="txtDBHost" id="txtDBHOST" type="text">
            </span>
		</td>
		</tr>
		<tr>
		<td>
		<span><h3>SQL DB port (optional):</h3></span>
		</td>
		<td>
		<span><input style="width:150px;border:1px dotted gray;padding-bottom:8px;" value="1433" title="SQL Port." name="txtDBPort" id="txtDBPort" type="text"></span>
		</td>
		</tr>
		<tr>
		<td>
		<span><h3>Database Name:</h3></span>
		</td>
		<td>
		<span><input style="width:150px;border:1px dotted gray;padding-bottom:8px;" value="" title="The name of the database." name="txtDBNAME" id="txtDBNAME" type="text"></span>
		</td>
		</tr>
		<tr>
		<td>
		<span><h3>Connection Type:</h3></span>
		</td>
		<td>
		<span>
			<select size="1" name="txtConnectionString">
				<option value="1">SQL Authentication</option>
				<option value="2">Trusted Connection</option>
			</select>
		</span>
		</td>
		</tr>
		<tr>
		<td>
		<span><h3>SQL DB user:</h3></span>
		</td>
		<td>
		<span><input style="width:150px;border:1px dotted gray;padding-bottom:8px;" value="" title="Name of SQL user account." name="txtDBUser" id="txtDBUser" type="text"></span>
		</td>
		</tr>
		<tr>
		<td>
		<span><h3>SQL DB user password:</h3></span>
		</td>
		<td>
		<span><input style="width:150px;border:1px dotted gray;padding-bottom:8px;" value="" title="Password for above SQL user account." name="txtDBUSERPW" id="txtDBUSERPW" type="password"></span>
		</td>
		</tr>
		</table>
		<br>
       </div>
      <span style="text-align:center;width:100%;">
        <input type="submit" id="btnSubmit" name="btnTest" value="Test connection" />
      <span>
  </div>
  <div class="content-bottom">
  </div>
</div>
<%
On Error Resume Next
If Request.Form("btnTest") <> "" Then
	If txtConnectionString.value = 1 then
		Set Conn = Server.CreateObject("ADODB.Connection")
		if request.form("txtDBPort") <> "" then sDBPort = "," & request.form("txtDBPort")
		Conn.Open "Driver={SQL Server}; Server=" & request.form("txtDBHost") & _
			"; Database=" & request.form("txtDBNAME") & sDBPort & "; UID=" & request.form("txtDBUSER") & "; PWD=" & request.form("txtDBUSERPW")
	Elseif txtConnectionString.value = 2 then
		Set Conn = Server.CreateObject("ADODB.Connection")
		if request.form("txtDBPort") <> "" then sDBPort = "," & request.form("txtDBPort")
		Conn.Open "Driver={SQL Server}; Server=" & request.form("txtDBHost") & _
			"; Database=" & request.form("txtDBNAME") & sDBPort & "; Trusted_Connection=yes;"
	End if
	If err.number <> 0 then
		Response.write "<br><br><img src='./images/ico_error.gif'>" & _
			"<span> Error: Cannot connect to server / database.<span><br>" & _
			"<br>"
		For Each objError in Conn.Errors
			Response.Write "Error Number " & objError.Number & ": " & objError.Description & "<br>"
		Next
		Response.write "<br>Next step. Try <b><i>telnet localhost 1433</b></i> on your SQL server."
	Else
		Response.write "<br><br><img src='./images/ico_success.gif'>" & _
			"<span> Connection successful!<span><br>"
	End if
End If
%>
<br>
</body>
</html>
<!--#include file="./footer.asp"-->