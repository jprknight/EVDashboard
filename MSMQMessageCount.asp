<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<!--#include file="./includes/styles.asp"-->
<form id="frmconfig" method="post" action="MSMQMessageCount.asp">
<br><br>
<h1><img src="images/queue.png"> MSMQ - Message Counts (Live Data)</h1>
<br><br>
<SELECT name="Servers">
<%
For i = 0 to Ubound(EVServersList,2)
	If InStr(EVServersList(5,i),"Email") then
		Response.Write "<OPTION VALUE = '" & EVServersList(1,i) & "'>" & EVServersList(1,i) & "</Option>"
	End if
Next
%>
</select>
<span><input type="submit" id="btnSubmit" name="btnSubmit" value="Selected Server"/></button></span>
<span><input type="submit" id="btnSubmit" name="btnSubmit" value="All Servers"/></button></span>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th nowrap width="90px" align="left">Server</th>
<th nowrap align="left">Server \ MSMQ Path</th>
<th nowrap width="80px" align="left"># Messages</th>
<th nowrap width="80px" align="left">Size (KB)</th>
</tr>
</thead>
<tbody>
<%

Dim MSMQApp ' As MsmqApplication
Set MSMQApp = CreateObject("MSMQ.MSMQApplication")
Dim qFormat ' As String

Dim Mgmt ' As new MSMQManagement


If Request.Form("btnSubmit") = "All Servers" Then
	On Error Resume Next

	For i = 0 to Ubound(EVServersList,2)

		MSMQApp.Machine = EVServersList(0,i)

		'Dim qFormat ' As String
		For each qFormat in MSMQApp.PrivateQueues
			
			Set Mgmt = CreateObject("MSMQ.MSMQManagement")
			Mgmt.Init MSMQApp.Machine,,"DIRECT=OS:" & qFormat
			
			If CLng(Mgmt.MessageCount) >= 1000 then
				Response.write ("<tr><td><h5>" & EVServersList(0,i) & "</td></h5><td><h5>" & _
					qFormat & "</h5></td><td><h5>" & _
					CLng(Mgmt.MessageCount) & "</h5></td><td><h5>" & _
					int(CLng(Mgmt.BytesInQueue) / 1024) & "</h5></td></tr>")
			Else
				Response.write ("<tr><td>" & EVServersList(0,i) & "</td><td>" & qFormat & "</td><td>" & CLng(Mgmt.MessageCount) & _
				"</td><td>" & int(CLng(Mgmt.BytesInQueue) / 1024) & "</td></tr>")
			End if
		Next
	Next
Elseif Request.Form("btnSubmit") = "Selected Server" Then
	'Dim MSMQApp ' As MsmqApplication
	'Set MSMQApp = CreateObject("MSMQ.MSMQApplication")

'	On Error Resume Next
	MSMQApp.Machine = Request.Form("Servers")

	'Dim qFormat ' As String
	For each qFormat in MSMQApp.PrivateQueues
		'Dim Mgmt ' As new MSMQManagement
		Set Mgmt = CreateObject("MSMQ.MSMQManagement")
		Mgmt.Init MSMQApp.Machine,,"DIRECT=OS:" & qFormat
		
		If CLng(Mgmt.MessageCount) >= 1000 then
			Response.write ("<tr><td><h5>" & Request.Form("Servers") & "</td></h5><td><h5>" & _
				qFormat & "</h5></td><td><h5>" & _
				CLng(Mgmt.MessageCount) & "</h5></td><td><h5>" & _
				int(CLng(Mgmt.BytesInQueue) / 1024) & "</h5></td></tr>")
		Else
			Response.write ("<tr><td>" & Request.Form("Servers") & "</td><td>" & qFormat & "</td><td>" & CLng(Mgmt.MessageCount) & _
			"</td><td>" & int(CLng(Mgmt.BytesInQueue) / 1024) & "</td></tr>")
		End if
	Next
End if
%>
</table><br>
</form>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		sort: true,
		col_0: "select",    
		display_all_text: " [ Show all ] ",  
		sort_config: {
			sort_types:['String','String','Number','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script>
<!--#include file="footer.asp"-->