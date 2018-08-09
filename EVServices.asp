<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br><br>
<h1><img src="images/server32x32.png"> EV Services (Live Data)</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th align="left" width="100px">DNS Alias</th>
<th align="left" width="100px">Real Name</th>
<th align="left" width="100px">Role</th>
<th align="left" width="120px">Service</th>
<th align="left" width="100px">Status</th>
</tr>
</thead>
<tbody>
<%

On Error Resume Next

For i = 0 to Ubound(EVServersList,2)
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & EVServersList(0,i) & "\root\cimv2")
	Set colRunningServices = objWMIService.ExecQuery _
		("Select * from Win32_Service")
	For Each objService in colRunningServices
		Select case objService.DisplayName 
			Case "Enterprise Vault Admin Service"
				If objService.State <> "Running" then	
					Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" &_
						EVServersList(1,i) & "</h5></td><td><h5>" &_
						EVServersList(5,i) & "</h5></td><td><h5>" & _
						objService.DisplayName & "</h5></td><td><h5>" & objService.State & "</h5></td></tr>")
				Else
					Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & EVServersList(1,i) & _
						"</td><td>" & EVServersList(5,i) & "</td><td>" & objService.DisplayName & _
						"</td><td>" & objService.State & "</td></tr>")
				End if
			Case "Enterprise Vault Directory Service"
				If objService.State <> "Running" then	
					Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" &_
						EVServersList(1,i) & "</h5></td><td><h5>" &_
						EVServersList(5,i) & "</h5></td><td><h5>" & _
						objService.DisplayName & "</h5></td><td><h5>" & objService.State & "</h5></td></tr>")
				Else
					Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & EVServersList(1,i) & _
						"</td><td>" & EVServersList(5,i) & "</td><td>" & objService.DisplayName & _
						"</td><td>" & objService.State & "</td></tr>")
				End if
			Case "Enterprise Vault Indexing Service"
				If objService.State <> "Running" then	
					Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" &_
						EVServersList(1,i) & "</h5></td><td><h5>" &_
						EVServersList(5,i) & "</h5></td><td><h5>" & _
						objService.DisplayName & "</h5></td><td><h5>" & objService.State & "</h5></td></tr>")
				Else
					Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & EVServersList(1,i) & _
						"</td><td>" & EVServersList(5,i) & "</td><td>" & objService.DisplayName & _
						"</td><td>" & objService.State & "</td></tr>")
				End if
			Case "Enterprise Vault Shopping Service"
				If objService.State <> "Running" then	
					Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" &_
						EVServersList(1,i) & "</h5></td><td><h5>" &_
						EVServersList(5,i) & "</h5></td><td><h5>" & _
						objService.DisplayName & "</h5></td><td><h5>" & objService.State & "</h5></td></tr>")
				Else
					Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & EVServersList(1,i) & _
						"</td><td>" & EVServersList(5,i) & "</td><td>" & objService.DisplayName & _
						"</td><td>" & objService.State & "</td></tr>")
				End if
			Case "Enterprise Vault Storage Service"
				If objService.State <> "Running" then	
					Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" &_
						EVServersList(1,i) & "</h5></td><td><h5>" &_
						EVServersList(5,i) & "</h5></td><td><h5>" & _
						objService.DisplayName & "</h5></td><td><h5>" & objService.State & "</h5></td></tr>")
				Else
					Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & EVServersList(1,i) & _
						"</td><td>" & EVServersList(5,i) & "</td><td>" & objService.DisplayName & _
						"</td><td>" & objService.State & "</td></tr>")
			End if
			Case "Enterprise Vault Task Controller Service"
				If objService.State <> "Running" then	
					Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" &_
						EVServersList(1,i) & "</h5></td><td><h5>" &_
						EVServersList(5,i) & "</h5></td><td><h5>" & _
						objService.DisplayName & "</h5></td><td><h5>" & objService.State & "</h5></td></tr>")
				Else
					Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & EVServersList(1,i) & _
						"</td><td>" & EVServersList(5,i) & "</td><td>" & objService.DisplayName & _
						"</td><td>" & objService.State & "</td></tr>")
				End if
			Case "Message Queuing"
				If objService.State <> "Running" then	
					Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" &_
						EVServersList(1,i) & "</h5></td><td><h5>" &_
						EVServersList(5,i) & "</h5></td><td><h5>" & _
						objService.DisplayName & "</h5></td><td><h5>" & objService.State & "</h5></td></tr>")
				Else
					Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & EVServersList(1,i) & _
						"</td><td>" & EVServersList(5,i) & "</td><td>" & objService.DisplayName & _
						"</td><td>" & objService.State & "</td></tr>")
				End if
		End Select
	Next
Next
%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		col_0: "select",
		col_1: "select",
		col_2: "select",
		col_3: "select",
		col_4: "select",
		display_all_text: " [ Show all ] ",  
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