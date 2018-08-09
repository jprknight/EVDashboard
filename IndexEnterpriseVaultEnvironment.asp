<span>
<h1>Enterprise Vault Environment (<%=SiteName%>)</h1><br><br>
<table id="table1" class="insidetable">
<thead>
	<tr>
		<th nowrap width="80px" align="left">DNS Alias</th>
		<th nowrap width="70px" align="left">EV Server</th>
		<th nowrap width="80px" align="left">Role</th>
		<th nowrap width="70px" align="left">Datacenter</th>
		<th nowrap width="80px" align="left">DB Server</th>
		<th nowrap width="80px" align="left">Database</th>
	</tr>
</thead>
<tbody>
<%
For i = 0 to Ubound(EVServersList,2)
	Response.Write("<tr><td>" & UCase(EVServersList(0,i))	& "</td><td>" & UCase(EVServersList(1,i)) & "</td><td>" & _
		EVServersList(5,i) & "</td><td>" & EVServersList(2,i) & "</td><td>" & UCase(EVServersList(3,i)) & _
		"</td><td>" & EVServersList(4,i) & "</td></tr>")
Next
%>
</tbody>
</table><br><br>
</span>