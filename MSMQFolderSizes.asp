<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br><br>
<h1><img src="images/server32x32.png"> MSMQ - Folder Sizes (Live Data)</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th>Server / MSMQ Path</th>
<th>Server</th>
<th>Role</th>
<th>Folder Size (MB)</th>
</tr>
</thead>
<tbody>
<%

'For each strServer in EnterpriseVaultAltServersArray
For i = 0 to Ubound(EVServersList,2)

	'* Replace : symbol with $ in MSMQ Path.
	If Instr(EVServersList(6,i),":") then
		EVServersList(6,i) = replace(EVServersList(6,i),":","$")
	end if
	
	'* If the MSMQ Path does not end with \ append it.
	If Right(EVServersList(6,i),1) <> "\" then
		EVServersList(6,i) = EVServersList(6,i) & "\"
	End if
	
	set oFolder = fso.GetFolder("\\" & EVServersList(1,i) & "\" & EVServersList(6,i))
	FolderSize = (FormatNumber(oFolder.Size / 1024 / 1024,2))
	Q1 = Instr(FolderSize,".")
	FolderSizeEvaluator = mid(FolderSize,1,Q1 - 1)
	
	If FolderSizeEvaluator > 1000 then
		Response.write ("<tr><td><h5>" & oFolder.path & "</h5></td><td><h5>" & _
			EVServersList(3,i) & "</h5></td><td><h5>" & EVServersList(5,i) & _
			"</h5></td><td><h5>" & FolderSize & "MB</h5></td></tr>")
	Else
		Response.write ("<tr><td>" & oFolder.path & "</td><td>" & EVServersList(1,i) & _
			"</td><td>" & EVServersList(5,i) & "</td><td>" & FolderSize & "</td></tr>")
	End if
	
Next
%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->