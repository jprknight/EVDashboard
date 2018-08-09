<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br><br>
<h1>Version 7.5 and earlier KVS Storage Registry Keys (Live Data)</h1>
<br><br>
<table id="table1" class="mytable">
<thead>
<tr>
<th align="left">Server</th>
<th align="left">Enable Archive</th>
<th align="left">Enable Expiry</th>
<th align="left">Enable File Watch</th>
<th align="left">Enable Replay Index</th>
<th align="left">Enable Restore</th>
<th align="left">Ignore Archive Bit</th>
</tr>
</thead>
<tbody>
<%
const HKEY_LOCAL_MACHINE = &H80000002
	
For i = 0 to Ubound(EVServersList,2)

	Response.write ("<tr><td>" & EVServersList(1,i) & "</td>")
	
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
	EVServersList(1,i) & "\root\default:StdRegProv")
	
	strKeyPath = "SOFTWARE\KVS\Enterprise Vault\Storage"
	strValueName = "EnableArchive"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	If dwValue = 0 then
		Response.write ("<td><h5>" & dwValue & "</td>")
	Elseif dwValue > 0 then
		Response.write ("<td>" & dwValue & "</td>")
	Else
		Response.write ("<td><h5>Key not set</h5></td>")
	End if
	
	strKeyPath = "SOFTWARE\KVS\Enterprise Vault\Storage"
	strValueName = "EnableExpiry"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	If dwValue = 0 then
		Response.write ("<td><h5>" & dwValue & "</h5></td>")
	Elseif dwValue > 0 then
		Response.write ("<td>" & dwValue & "</td>")
	Else
		Response.write ("<td><h5>Key not set</h5></td>")
	End if
	
	strKeyPath = "SOFTWARE\KVS\Enterprise Vault\Storage"
	strValueName = "EnableFileWatch"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	If dwValue = 0 then
		Response.write ("<td><h5>" & dwValue & "</h5></td>")
	Elseif dwValue > 0 then
		Response.write ("<td>" & dwValue & "</td>")
	Else
		Response.write ("<td><h5>Key not set</h5></td>")
	End if
	
	strKeyPath = "SOFTWARE\KVS\Enterprise Vault\Storage"
	strValueName = "EnableReplayIndex"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	If dwValue = 0 then
		Response.write ("<td><h5>" & dwValue & "</h5></td>")
	Elseif dwValue > 0 then
		Response.write ("<td>" & dwValue & "</td>")
	Else
		Response.write ("<td><h5>Key not set</h5></td>")
	End if
	
	strKeyPath = "SOFTWARE\KVS\Enterprise Vault\Storage"
	strValueName = "EnableRestore"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	If dwValue = 0 then
		Response.write ("<td><h5>" & dwValue & "</h5></td>")
	Elseif dwValue > 0 then
		Response.write ("<td>" & dwValue & "</td>")
	Else
		Response.write ("<td><h5>Key not set</h5></td>")
	End if
	
	strKeyPath = "SOFTWARE\KVS\Enterprise Vault\Storage"
	strValueName = "FileWatchEnableIgnoreArchiveBit"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
	If dwValue = 0 then
		Response.write ("<td><h5>" & dwValue & "</h5></td>")
	Elseif dwValue > 0 then
		Response.write ("<td>" & dwValue & "</td>")
	Else
		Response.write ("<td><h5>Key not set</h5></td>")
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
			sort_types:['String','String','String','String','String','String','String']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->