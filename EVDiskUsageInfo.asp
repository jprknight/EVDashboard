<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<head>
	<script type="text/javascript">
		window.addEvent('domready', function() { myCal = new Calendar({ date: 'Y/m/d' }); });
	</script>
</head>
<%
'Get today's date.
Styr = Year(Now)
Stmo = Month(Now)
if Stmo < 10 Then
	Stmo = "0" & Stmo
end if
Stdt = Day(Now)
if Stdt < 10 Then
	Stdt = "0" & Stdt
end if

'React differently depending on whether no buttons have been pressed, refresh data or today buttons have been pressed.
If Request.Form("btn") = "Refresh Data" Then
	If request.form("date") <> "Live Data" then
		SQLWhereClause = "WHERE RecordCreateTimestamp between '" & request.form("date") & " 00:00:00' AND '" & request.form("date") & " 23:59:59'"
		ProcessDate = request.form("date")
	else
		SQLWhereClause = "WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
			"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
		ProcessDate = styr & "/" & stmo & "/" & stdt
	End if
Elseif Request.Form("btn") = "Today" then
	SQLWhereClause = "WHERE RecordCreateTimestamp between '" & styr & "/" & stmo & "/" & stdt & " 00:00:00' " & _ 
		"AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	ProcessDate = styr & "/" & stmo & "/" & stdt
Else
	SQLWhereClause = "WHERE RecordCreateTimestamp between '" & _
		styr & "/" & stmo & "/" & stdt & " 00:00:00' AND '" & styr & "/" & stmo & "/" & stdt & " 23:59:59'"
	ProcessDate = styr & "/" & stmo & "/" & stdt
End if
%>
<!-- Main Form -->
<form id="frmconfig" method="post" action="EVDiskUsageInfo.asp">
<!-- Table for page heading -->
<table width="100%" class="Main">
<tr>
<td nowrap="nowrap" width="350px">
<h1><img src="images/server32x32.png"> EV Disk Usage Information</h1>
</td>
<td nowrap="nowrap" width="400px" align="right">
<input id="date" name="date" type="text" value="<%=ProcessDate%>" />
<span style="text-align:left;width:100%;"><input type="submit" name="btn" value="Refresh Data" class="pgInp"/></button></span>
<span style="text-align:left;width:100%;"><input type="submit" name="btn" value="Today" class="pgInp"/></button></span>
<span style="text-align:left;width:100%;"><input type="submit" name="btn" value="Live Data" class="pgInp"/></button></span>
</td>
</tr>
</table>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th nowrap width="100px" align="left">Server</th>
<th nowrap align="left">Role</th>
<th nowrap align="left">Drive</th>
<th nowrap align="left">Volume Name</th>
<th nowrap align="left">Size (GB)</th>
<th nowrap align="left">Used (GB)</th>
<th nowrap align="left">Free (GB)</th>
<th nowrap align="left">Free (%)</th>
</tr>
</thead>
<tbody>
<%
If Request.Form("btn") = "Refresh Data" Then
	Call HistoryData
Elseif Request.Form("btn") = "Today" then
	Call HistoryData
Elseif Request.Form("btn") = "Live Data" Then
	Call LiveData
else
	Call HistoryData
end if

Function HistoryData
	'Retrieve data from EVDashboard database.
	If EVDConnectionString = 1 Then
		EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
			EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
	Elseif EVDConnectionString = 2 Then
		EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
			EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
	End if

	'* Open database connection
	Set EVDmyconn = CreateObject("adodb.connection")
	EVDmyconn.open (EVDconnection)

	If cstr(err.number) <> 0 Then
		wscript.echo (" - Error creating connection to " _
	   & "database server " & EVDSQLServer & " / " & EVDSQLDatabase & ".  Check your connection string " _
	   & "or database server name/IP and try again.")
	   wscript.quit
	End if

	On Error Goto 0
		
	Set result = CreateObject("adodb.recordset")
	If err.number <> 0 then msgbox err.description

	SQLQuery = ("Select EVServer,EVRole,Drive,VolumeName,SizeGB,UsedGB," & _
		"FreeGB,PercentageFree from DiskUsageInfo " & SQLWhereClause & " AND EVRole <> 'Email' ORDER BY EVServer, Drive")

	'* Execute the query
	Set result = EVDmyconn.execute(SQLQuery)
	If err.number <> 0 then msgbox err.description

	While not result.EOF
		' For some reason this logic check does not work...???
		'If (result("PercentageFree")) < 25 then
		'	Response.Write("<tr><td><h5>" & (result("EVServer")) & _
		'		"</h5></td><td><h5>" & UCase((result("EVRole"))) & _
		'		"</h5></td><td><h5>" & (result("Drive")) & _
		'		"</h5></td><td><h5>" & (result("VolumeName")) & _
		'		"</h5></td><td><h5>" & (result("SizeGB")) & _
		'		"</h5></td><td><h5>" & (result("UsedGB")) & _
		'		"</h5></td><td><h5>" & (result("FreeGB")) & _
		'		"</h5></td><td><h5>" & (result("PercentageFree")) & "</h5></td></tr>")
		'Else
			Response.Write("<tr><td>" & (result("EVServer")) & _
				"</td><td>" & (result("EVRole")) & _
				"</td><td>" & (result("Drive")) & _
				"</td><td>" & (result("VolumeName")) & _
				"</td><td>" & (result("SizeGB")) & _
				"</td><td>" & (result("UsedGB")) & _
				"</td><td>" & (result("FreeGB")) & _
				"</td><td>" & (result("PercentageFree")) & "</td></tr>")
		'End if		
		
		result.movenext()
	Wend
End Function


'* Function for when LIVE button is pressed.
Function LiveData
	On Error Resume Next
	
	For i = 0 to Ubound(EVServersList,2)
		
		Dim objWMIService, objItem, colItems
		Dim strDriveType, strDiskSize

		Set objWMIService = GetObject("winmgmts:\\" & EVServersList(1,i) & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk WHERE DriveType=3")
		
		For Each objItem in colItems
				
			DIM pctFreeSpace,strFreeSpace,strusedSpace
			
			pctFreeSpace = INT((objItem.FreeSpace / objItem.Size) * 1000)/10' & "%"
			strDiskSize = Int(objItem.Size /1073741824)' & "Gb"
			strFreeSpace = Int(objItem.FreeSpace /1073741824)' & "Gb"
			strUsedSpace = Int((objItem.Size-objItem.FreeSpace)/1073741824)' & "Gb"
			
			'Q1 = Instr(pctFreeSpace,"%")
			'Q2 = mid(pctFreeSpace,"1",Q1 - 1)
			
			'If pctFreeSpace < 25 then
			'	Response.Write("<tr><td><h5>" & UCase(EVServersList(1,i)) & "</h5></td>" & _
			'		"<td><h5>" & _
			'		objItem.Name & "</h5></td><td><h5>" & objItem.VolumeName & _
			'		"</h5></td>" & "<td><h5>" & strDiskSize & "</h5></td><td><h5>" & _
			'		strUsedSpace & "</h5></td><td><h5>" & strFreeSpace & "</h5></td><td><h5>" & _
			'		pctFreeSpace & "</h5></td></tr>")
			'Else
				Response.Write("<tr><td>" & UCase(EVServersList(1,i)) & _
					"</td><td>" & EVServersList(5,i) & _
					"</td><td>" & objItem.Name & _
					"</td><td>" & objItem.VolumeName & _
					"</td><td>" & strDiskSize & _
					"</td><td>" & strUsedSpace & _
					"</td><td>" & strFreeSpace & _
					"</td><td>" & pctFreeSpace & "</td></tr>")
			'End if
		Next
	Next
end function
%>
</tbody>
</table><br>
</form>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		on_keyup: true,
		on_keyup_delay: 1200,
		col_0: "select",    
		col_2: "select",    
		display_all_text: " [ Show all ] ",  
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number','Number','Number','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->