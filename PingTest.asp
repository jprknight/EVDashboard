<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br><br>
<h1><img src="images/server32x32.png"> Ping Test (Live Data)</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th>Server</th>
<th>Role</th>
<th>Status</th>
<th>Bytes</th>
<th>Time (ms)</th>
<th>TTL (s)</th>
</tr>
</thead>
<tbody>
<%

Dim SuccessList, FailureList

For i = 0 to Ubound(EVServersList,2)
	Dim objPing, objRetStatus
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
		("select * from Win32_PingStatus where address = '" & EVServersList(0,i) & "'")
	For Each objRetStatus in objPing
		If IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 Then
			Response.write ("<tr><td><h5>" & UCase(EVServersList(0,i)) & "</h5></td><td><h5>" & _
				EVServersList(5,i) & "</h5></td><td><h5>Failure</h5></td>")
		Else
			Response.write ("<tr><td>" & UCase(EVServersList(0,i)) & "</td><td>" & _
				EVServersList(5,i) & "</td><td>Success</td><td>" & _
				objRetStatus.BufferSize & "</td><td>" & objRetStatus.ResponseTime & _
				"</td><td>" & objRetStatus.ResponseTimeToLive & "</td>")
		End if
	Next
Next

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
