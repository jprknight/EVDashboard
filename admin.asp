<!--#include file="./config.asp"-->
<!--#include file="./header.asp"-->
<!--#include file="./includes/styles.asp"-->
<head>
<!--#include file="./includes/AdminStyles.asp"-->
<head>
<br><br>
<h1><img src="images/configure.png"> Configure EV Dashboard</h1><br><br>
<div class="content">
	<div id="savediv" style="width:100%">
			<h2>XML Config File Version</h2><br>
		<form id="frmconfig" method="post" action="admin.asp">
			<table class="Main">
			<tr><td>
			<div>
			<span><h3>XML Config File Version:</h3></span>
			</td>
			<td>
			<span><h3><%=ConfigVersion%></h3></span>
			</td>
			<tr><td>
			<span><h3>System Check XML Version:</h3></span>
			</td>
			<td>
			<span><h3><%=CheckXMLConfigVersion%></h3></span>
			</td>
			<td>
			<a href="#" class="hintanchor" onMouseover="showhint('This is a validation check between the version of the config.xml file expected and the version actually loaded.', this, event, '300px')">[?]</a>
			</td>
			</tr>
			</table>
			<br>
			<hr size="1" color="6297a6">
			<br>
			<h2>Enterprise Vault SQL Configuration</h2><br>
			<table class="Main">
			<tr>
			<td>
			<span><h3>Enterprise Vault Version</h3></span>
			</td>
			</tr>
			<tr>
			<td>
			<span>
				<select size="1" name="txtEVVersion">
					<option value="1">Pre 9.0</option>
					<option value="2">9.0 and newer</option>
				</select>
			</span>
			<a href="#" class="hintanchor" onMouseover="showhint('Enterprise Vault Version - A schema change in 9.0 means different SQL joins need to be used depending on the version of EV you use in your environment.', this, event, '650px')">[?]</a><br>
			</td>
			</tr>
			<tr>
			<td>
			<span><h3>SQL Directory Server:</h3></span>
			</td>
			<td>
			<span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtDirectoryServer" id="txtDirectoryServer" type="text">
			</span><a href="#" class="hintanchor" onMouseover="showhint('The name or IP address of your SQL server hosting the directory database.', this, event, '300px')">[?]</a><br>
			</td>
			</tr>
			<tr>
			<td>
			<span><h3>Connection String:</h3></span>
			</td>
			<td>
			<span>
				<select size="1" name="txtConnectionString">
					<option value="1">SQL Authentication</option>
					<option value="2">Trusted Connection</option>
				</select>
			</span>
			<a href="#" class="hintanchor" onMouseover="showhint('SQL Authentication - Use sql username and password for database connections.<br>Driver={SQL Server};Server=<i>SQL_Server</i>;Database=<i>SQL_Database</i>;Uid=<i>SQL_User</i>;Pwd=<i>SQL_Password</i>; <br>Trusted Connection - Use trusted connection for database connections.<br>Driver={SQL Server};Server=<i>SQL_Server</i>;Database=<i>SQL_Database</i>;Trusted_Connection=yes;', this, event, '650px')">[?]</a><br>
			</td>
			</tr>
			<tr>
			<td>
			<span><h3>SQL User Name:</h3></span>
			</td>
			<td>
			<span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtSQLUserName" id="txtSQLUserName" type="text">
			</span><a href="#" class="hintanchor" onMouseover="showhint('Use if you are using SQL Authentication. Note: this information is stored in the plain text file /config/config.xml; this login should be read-only.', this, event, '300px')">[?]</a><br>
			</td>
			<td>
			<span><h3>SQL Password:</h3></span>
			</td>
			<td>
			<span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtSQLPassword" id="txtSQLPassword" type="Password">
			</span>
			<a href="#" class="hintanchor" onMouseover="showhint('Use if you are using SQL Authentication. Note: this information is stored in the plain text file /config/config.xml; this login should be read-only.', this, event, '300px')">[?]</a>
			</td>
			</tr>
			<tr>
			<td><span><h3>Alternative SQL Port (Normally 1433)</h3></span></td>
			<td>
			
				<select size="1" name="txtSQLUseAltPort">
					<option value="1">Yes</option>
					<option value="2">No</option>
				</select>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtAltPortNumber" id="txtAltPortNumber" type="text">
			<a href="#" class="hintanchor" onMouseover="showhint('SQL Alternative Port - Do not change this unless you are sure you need to!', this, event, '650px')">[?]</a>
			</td>
			</tr>
			</table>
			<br>
			<hr size="1" color="6297a6">
			<br>
			<h2>EV Dashboard SQL Configuration</h2><br>
			<table class="Main">
			<tr>
			<td>
			<span><h3>SQL Server:</h3></span>
			</td>
			<td>
			<span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtEVDSQLServer" id="txtEVDSQLServer" type="text">
			</span><a href="#" class="hintanchor" onMouseover="showhint('The name or IP address of your SQL server hosting the directory database.', this, event, '300px')">[?]</a><br>
			</td>
			</tr>
			<tr>
			<td>
			<span><h3>SQL Database:</h3></span>
			</td>
			<td>
			<span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtEVDSQLDatabase" id="txtEVDSQLDatabase" type="text">
			</span><a href="#" class="hintanchor" onMouseover="showhint('The name or IP address of your SQL server hosting the directory database.', this, event, '300px')">[?]</a><br>
			</td>
			</tr>
			<tr>
			<td>
			<span><h3>Connection String:</h3></span>
			</td>
			<td>
			<span>
				<select size="1" name="txtEVDConnectionString">
					<option value="1">SQL Authentication</option>
					<option value="2">Trusted Connection</option>
				</select>
			</span>
			<a href="#" class="hintanchor" onMouseover="showhint('SQL Authentication - Use sql username and password for database connections.<br>Driver={SQL Server};Server=<i>SQL_Server</i>;Database=<i>SQL_Database</i>;Uid=<i>SQL_User</i>;Pwd=<i>SQL_Password</i>; <br>Trusted Connection - Use trusted connection for database connections.<br>Driver={SQL Server};Server=<i>SQL_Server</i>;Database=<i>SQL_Database</i>;Trusted_Connection=yes;', this, event, '650px')">[?]</a><br>
			</td>
			</tr>
			<tr>
			<td>
			<span><h3>SQL User Name:</h3></span>
			</td>
			<td>
			<span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtEVDSQLUserName" id="txtEVDSQLUserName" type="text">
			</span><a href="#" class="hintanchor" onMouseover="showhint('Use if you are using SQL Authentication. Note: this information is stored in the plain text file /config/config.xml; this login should be read-only.', this, event, '300px')">[?]</a><br>
			</td>
			<td>
			<span><h3>SQL Password:</h3></span>
			</td>
			<td>
			<span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtEVDSQLPassword" id="txtEVDSQLPassword" type="Password">
			</span>
			<a href="#" class="hintanchor" onMouseover="showhint('Use if you are using SQL Authentication. Note: this information is stored in the plain text file /config/config.xml; this login should be read-only.', this, event, '300px')">[?]</a>
			</td>
			</tr>
			<tr>
			<td>
			<span><h3>Alternative SQL Port (Normally 1433)</h3></span>
			</td>
			<td>
			<span>
				<select size="1" name="txtEVDSQLUseAltPort">
					<option value="1">Yes</option>
					<option value="2">No</option>
				</select>
			</span>
			<input style="width:140px;border:1px dotted gray;padding-bottom:8px;" name="txtEVDSQLAltPortNumber" id="txtEVDSQLAltPortNumber" type="text" text="1433">
			<a href="#" class="hintanchor" onMouseover="showhint('SQL Alternative Port - Do not change this unless you are sure you need to!', this, event, '650px')">[?]</a><br>
			</td>
			</tr>
			</table>
			<br>
			<hr size="1" color="6297a6">
			<br>
			<table class="Main">
			<tr>
			<td>
			<h2>EV Dashboard SQL Information</h2><br>
			</td>
			</tr>
			<tr>
			<td>
			<h3>Data last gathered between <b><i><%=StartTime%></i></b> and <b><i><%=Finishtime%></i></b>.</h3>
			</td>
			</tr>
			<tr>
			<td>
			<h3>EV Dashboard SQL Information - Edition: <b><i><%=EVDSQLEdition%></i></b>, Version: <b><i><%=EVDSQLVersion%></i></b>, Service Pack: <b><i><%=EVDSQLServicePack%></i></b>.</h3>
			</td>
			</tr>
			</table>
			<br><br>
			<hr size="1" color="6297a6">
			<br>
			<h2>EV Infrastructure</h2><br><br>
			<table class="Main">
			<tr>
				<td>
					<h3>#</h3>
				</td>
				<td>
					<h3>DNS Alias<a href="#" class="hintanchor" onMouseover="showhint('Building blocks DNS alias.', this, event, '150px')">[?]</a></h3>
				</td>
				<td>
					<h3>Server Real Name</h3>
				</td>
				<td>
					<h3>Datacenter</h3>
				</td>
				<td>
					<h3>Database</a></h3>
				</td>
				<td>
					<h3>DB Server</a></h3>
				</td>
				<td>
					<h3>Role</a></h3>
				</td>
				<td>
					<h3>MSMQ Path (i.e. d:\msmq\)</h3>
				</td>
				<td>
					<h3>Trigger File Path</h3>
				</td>
			</tr>
			<%
			' Setup loop to create 20 configuration entries for EV servers.
			For i = 1 to 20
				%>
				<tr>
				<td>
				<span><h3><%=i%></h3></span>
				</td>
				<td>
				<span>
				<input style="width:110px;border:1px dotted gray;padding-bottom:8px;" title="EV Server Building Blocks DNS Alias." name="txtDNSAlias<%=i%>" id="txtDNSAlias<%=i%>" type="text">
				</span>
				</td>
				<td>
				<span>
				<input style="width:110px;border:1px dotted gray;padding-bottom:8px;" title="Real EV Server Name." name="txtEVServer<%=i%>" id="txtEVServer<%=i%>" type="text">
				</span>
				</td>
				<td>
				<span>
				<input style="width:110px;border:1px dotted gray;padding-bottom:8px;" title="EV Server's datacenter, for information purposes only." name="txtDatacenter<%=i%>" id="txtDatacenter<%=i%>" type="text">
				</span>
				</td>
				<td>
				<span>
				<input style="width:110px;border:1px dotted gray;padding-bottom:8px;" title="EV SQL database name." name="txtEVDatabase<%=i%>" id="txtEVDatabase<%=i%>" type="text">
				</span>
				</td>
				<td>
				<span>
				<input style="width:110px;border:1px dotted gray;padding-bottom:8px;" title="EV SQL Database server name." name="txtDBServer<%=i%>" id="txtDBServer<%=i%>" type="text">
				</span>
				</td>
				<td>
				<span>
					<select size="1" name="txtRole<%=i%>">
						<option value="Email Archiving (Physical)">Email Archiving (Physical)</option>
						<option value="Email Archiving (Virtual)">Email Archiving (Virtual)</option>
						<option value="File Archiving (Physical)">File Archiving (Physical)</option>
						<option value="File Archiving (Virtual)">File Archiving (Virtual)</option>
						<option value="Standby (Physical)">Standby (Physical)</option>
						<option value="Standby (Virtual)">Standby (Virtual)</option>
					</select>
				</span>
				</td>
				<td>
				<span>
				<input style="width:150px;border:1px dotted gray;padding-bottom:8px;" title="EV Server MSMQ Path." name="txtMSMQPath<%=i%>" id="txtMSMQPath<%=i%>" type="text">
				</span>
				</td>
				<td>
				<span>
				<input style="width:275px;border:1px dotted gray;padding-bottom:8px;" title="EV Server Trigger File Path." name="txtTriggerPath<%=i%>" id="txtTriggerPath<%=i%>" type="text">
				</span>
				</td>
				</tr>				
			<%Next%>	
		</table>
	<br>
	<span style="text-align:center;width:100%;"><input type="submit" id="btnSubmit" name="btnSubmit" value="Save Changes" class="pgInp" /></button></span>
    </div>
</div>
  <script language="VBSCRIPT">
Function LoadValues
	'Load environment variables into the form...

	document.GetElementByID("txtDirectoryServer").value = "<%=DirectoryServer%>"

	document.GetElementByID("txtEVVersion").value = "<%=EVVersion%>"
	document.GetElementByID("txtConnectionString").value = "<%=ConnectionString%>"
	document.GetElementByID("txtSQLUserName").value = "<%=SQLUserName%>"
	document.GetElementByID("txtSQLPassword").value = "<%=SQLPassword%>"
	document.GetElementByID("txtSQLUseAltPort").value = "<%=SQLUseAltPort%>"
	document.GetElementByID("txtAltPortNumber").value = "<%=SQLAltPortNumber%>"
	
	document.GetElementByID("txtEVDSQLServer").value = "<%=EVDSQLServer%>"
	document.GetElementByID("txtEVDSQLDatabase").value = "<%=EVDSQLDatabase%>"
	
	document.GetElementByID("txtEVDConnectionString").value = "<%=EVDConnectionString%>"
	document.GetElementByID("txtEVDSQLUserName").value = "<%=EVDSQLUserName%>"
	document.GetElementByID("txtEVDSQLPassword").value = "<%=EVDSQLPassword%>"
	document.GetElementByID("txtEVDSQLUseAltPort").value = "<%=EVDSQLUseAltPort%>"
	document.GetElementByID("txtEVDSQLAltPortNumber").value = "<%=EVDSQLAltPortNumber%>"
		
	document.GetElementByID("txtDNSAlias1").value = "<%=DNSAlias1%>"
	document.GetElementByID("txtEVServer1").value = "<%=EVServer1%>"
	document.GetElementByID("txtDatacenter1").value = "<%=Datacenter1%>"
	document.GetElementByID("txtEVDatabase1").value = "<%=EVDatabase1%>"
	document.GetElementByID("txtDBServer1").value = "<%=DBServer1%>"
	document.GetElementByID("txtRole1").value = "<%=Role1%>"
	document.GetElementByID("txtMSMQPath1").value = "<%=MSMQPath1%>"
	document.GetElementByID("txtTriggerPath1").value = "<%=TriggerPath1%>"

	document.GetElementByID("txtDNSAlias2").value = "<%=DNSAlias2%>"
	document.GetElementByID("txtEVServer2").value = "<%=EVServer2%>"
	document.GetElementByID("txtDatacenter2").value = "<%=Datacenter2%>"
	document.GetElementByID("txtEVDatabase2").value = "<%=EVDatabase2%>"
	document.GetElementByID("txtDBServer2").value = "<%=DBServer2%>"
	document.GetElementByID("txtRole2").value = "<%=Role2%>"
	document.GetElementByID("txtMSMQPath2").value = "<%=MSMQPath2%>"
	document.GetElementByID("txtTriggerPath2").value = "<%=TriggerPath2%>"

	document.GetElementByID("txtDNSAlias3").value = "<%=DNSAlias3%>"
	document.GetElementByID("txtEVServer3").value = "<%=EVServer3%>"
	document.GetElementByID("txtDatacenter3").value = "<%=Datacenter3%>"
	document.GetElementByID("txtEVDatabase3").value = "<%=EVDatabase3%>"
	document.GetElementByID("txtDBServer3").value = "<%=DBServer3%>"
	document.GetElementByID("txtRole3").value = "<%=Role3%>"
	document.GetElementByID("txtMSMQPath3").value = "<%=MSMQPath3%>"
	document.GetElementByID("txtTriggerPath3").value = "<%=TriggerPath3%>"

	document.GetElementByID("txtDNSAlias4").value = "<%=DNSAlias4%>"
	document.GetElementByID("txtEVServer4").value = "<%=EVServer4%>"
	document.GetElementByID("txtDatacenter4").value = "<%=Datacenter4%>"
	document.GetElementByID("txtEVDatabase4").value = "<%=EVDatabase4%>"
	document.GetElementByID("txtDBServer4").value = "<%=DBServer4%>"
	document.GetElementByID("txtRole4").value = "<%=Role4%>"
	document.GetElementByID("txtMSMQPath4").value = "<%=MSMQPath4%>"
	document.GetElementByID("txtTriggerPath4").value = "<%=TriggerPath4%>"

	document.GetElementByID("txtDNSAlias5").value = "<%=DNSAlias5%>"
	document.GetElementByID("txtEVServer5").value = "<%=EVServer5%>"
	document.GetElementByID("txtDatacenter5").value = "<%=Datacenter5%>"
	document.GetElementByID("txtEVDatabase5").value = "<%=EVDatabase5%>"
	document.GetElementByID("txtDBServer5").value = "<%=DBServer5%>"
	document.GetElementByID("txtRole5").value = "<%=Role5%>"
	document.GetElementByID("txtMSMQPath5").value = "<%=MSMQPath5%>"
	document.GetElementByID("txtTriggerPath5").value = "<%=TriggerPath5%>"

	document.GetElementByID("txtDNSAlias6").value = "<%=DNSAlias6%>"
	document.GetElementByID("txtEVServer6").value = "<%=EVServer6%>"
	document.GetElementByID("txtDatacenter6").value = "<%=Datacenter6%>"
	document.GetElementByID("txtEVDatabase6").value = "<%=EVDatabase6%>"
	document.GetElementByID("txtDBServer6").value = "<%=DBServer6%>"
	document.GetElementByID("txtRole6").value = "<%=Role6%>"
	document.GetElementByID("txtMSMQPath6").value = "<%=MSMQPath6%>"
	document.GetElementByID("txtTriggerPath6").value = "<%=TriggerPath6%>"

	document.GetElementByID("txtDNSAlias7").value = "<%=DNSAlias7%>"
	document.GetElementByID("txtEVServer7").value = "<%=EVServer7%>"
	document.GetElementByID("txtDatacenter7").value = "<%=Datacenter7%>"
	document.GetElementByID("txtEVDatabase7").value = "<%=EVDatabase7%>"
	document.GetElementByID("txtDBServer7").value = "<%=DBServer7%>"
	document.GetElementByID("txtRole7").value = "<%=Role7%>"
	document.GetElementByID("txtMSMQPath7").value = "<%=MSMQPath7%>"
	document.GetElementByID("txtTriggerPath7").value = "<%=TriggerPath7%>"

	document.GetElementByID("txtDNSAlias8").value = "<%=DNSAlias8%>"
	document.GetElementByID("txtEVServer8").value = "<%=EVServer8%>"
	document.GetElementByID("txtDatacenter8").value = "<%=Datacenter8%>"
	document.GetElementByID("txtEVDatabase8").value = "<%=EVDatabase8%>"
	document.GetElementByID("txtDBServer8").value = "<%=DBServer8%>"
	document.GetElementByID("txtRole8").value = "<%=Role8%>"
	document.GetElementByID("txtMSMQPath8").value = "<%=MSMQPath8%>"
	document.GetElementByID("txtTriggerPath8").value = "<%=TriggerPath8%>"

	document.GetElementByID("txtDNSAlias9").value = "<%=DNSAlias9%>"
	document.GetElementByID("txtEVServer9").value = "<%=EVServer9%>"
	document.GetElementByID("txtDatacenter9").value = "<%=Datacenter9%>"
	document.GetElementByID("txtEVDatabase9").value = "<%=EVDatabase9%>"
	document.GetElementByID("txtDBServer9").value = "<%=DBServer9%>"
	document.GetElementByID("txtRole9").value = "<%=Role9%>"
	document.GetElementByID("txtMSMQPath9").value = "<%=MSMQPath9%>"
	document.GetElementByID("txtTriggerPath9").value = "<%=TriggerPath9%>"

	document.GetElementByID("txtDNSAlias10").value = "<%=DNSAlias10%>"
	document.GetElementByID("txtEVServer10").value = "<%=EVServer10%>"
	document.GetElementByID("txtDatacenter10").value = "<%=Datacenter10%>"
	document.GetElementByID("txtEVDatabase10").value = "<%=EVDatabase10%>"
	document.GetElementByID("txtDBServer10").value = "<%=DBServer10%>"
	document.GetElementByID("txtRole10").value = "<%=Role10%>"
	document.GetElementByID("txtMSMQPath10").value = "<%=MSMQPath10%>"
	document.GetElementByID("txtTriggerPath10").value = "<%=TriggerPath10%>"

	document.GetElementByID("txtDNSAlias11").value = "<%=DNSAlias11%>"
	document.GetElementByID("txtEVServer11").value = "<%=EVServer11%>"
	document.GetElementByID("txtDatacenter11").value = "<%=Datacenter11%>"
	document.GetElementByID("txtEVDatabase11").value = "<%=EVDatabase11%>"
	document.GetElementByID("txtDBServer11").value = "<%=DBServer11%>"
	document.GetElementByID("txtRole11").value = "<%=Role11%>"
	document.GetElementByID("txtMSMQPath11").value = "<%=MSMQPath11%>"
	document.GetElementByID("txtTriggerPath11").value = "<%=TriggerPath11%>"

	document.GetElementByID("txtDNSAlias12").value = "<%=DNSAlias12%>"
	document.GetElementByID("txtEVServer12").value = "<%=EVServer12%>"
	document.GetElementByID("txtDatacenter12").value = "<%=Datacenter12%>"
	document.GetElementByID("txtEVDatabase12").value = "<%=EVDatabase12%>"
	document.GetElementByID("txtDBServer12").value = "<%=DBServer12%>"
	document.GetElementByID("txtRole12").value = "<%=Role12%>"
	document.GetElementByID("txtMSMQPath12").value = "<%=MSMQPath12%>"
	document.GetElementByID("txtTriggerPath12").value = "<%=TriggerPath12%>"

	document.GetElementByID("txtDNSAlias13").value = "<%=DNSAlias13%>"
	document.GetElementByID("txtEVServer13").value = "<%=EVServer13%>"
	document.GetElementByID("txtDatacenter13").value = "<%=Datacenter13%>"
	document.GetElementByID("txtEVDatabase13").value = "<%=EVDatabase13%>"
	document.GetElementByID("txtDBServer13").value = "<%=DBServer13%>"
	document.GetElementByID("txtRole13").value = "<%=Role13%>"
	document.GetElementByID("txtMSMQPath13").value = "<%=MSMQPath13%>"
	document.GetElementByID("txtTriggerPath13").value = "<%=TriggerPath13%>"
	
	document.GetElementByID("txtDNSAlias14").value = "<%=DNSAlias14%>"
	document.GetElementByID("txtEVServer14").value = "<%=EVServer14%>"
	document.GetElementByID("txtDatacenter14").value = "<%=Datacenter14%>"
	document.GetElementByID("txtEVDatabase14").value = "<%=EVDatabase14%>"
	document.GetElementByID("txtDBServer14").value = "<%=DBServer14%>"
	document.GetElementByID("txtRole14").value = "<%=Role14%>"
	document.GetElementByID("txtMSMQPath14").value = "<%=MSMQPath14%>"
	document.GetElementByID("txtTriggerPath14").value = "<%=TriggerPath14%>"
	
	document.GetElementByID("txtDNSAlias15").value = "<%=DNSAlias15%>"
	document.GetElementByID("txtEVServer15").value = "<%=EVServer15%>"
	document.GetElementByID("txtDatacenter15").value = "<%=Datacenter15%>"
	document.GetElementByID("txtEVDatabase15").value = "<%=EVDatabase15%>"
	document.GetElementByID("txtDBServer15").value = "<%=DBServer15%>"
	document.GetElementByID("txtRole15").value = "<%=Role15%>"
	document.GetElementByID("txtMSMQPath15").value = "<%=MSMQPath15%>"
	document.GetElementByID("txtTriggerPath15").value = "<%=TriggerPath15%>"
	
	document.GetElementByID("txtDNSAlias16").value = "<%=DNSAlias16%>"
	document.GetElementByID("txtEVServer16").value = "<%=EVServer16%>"
	document.GetElementByID("txtDatacenter16").value = "<%=Datacenter16%>"
	document.GetElementByID("txtEVDatabase16").value = "<%=EVDatabase16%>"
	document.GetElementByID("txtDBServer16").value = "<%=DBServer16%>"
	document.GetElementByID("txtRole16").value = "<%=Role16%>"
	document.GetElementByID("txtMSMQPath16").value = "<%=MSMQPath16%>"
	document.GetElementByID("txtTriggerPath16").value = "<%=TriggerPath16%>"
	
	document.GetElementByID("txtDNSAlias17").value = "<%=DNSAlias17%>"
	document.GetElementByID("txtEVServer17").value = "<%=EVServer17%>"
	document.GetElementByID("txtDatacenter17").value = "<%=Datacenter17%>"
	document.GetElementByID("txtEVDatabase17").value = "<%=EVDatabase17%>"
	document.GetElementByID("txtDBServer17").value = "<%=DBServer17%>"
	document.GetElementByID("txtRole17").value = "<%=Role17%>"
	document.GetElementByID("txtMSMQPath17").value = "<%=MSMQPath17%>"
	document.GetElementByID("txtTriggerPath17").value = "<%=TriggerPath17%>"
	
	document.GetElementByID("txtDNSAlias18").value = "<%=DNSAlias18%>"
	document.GetElementByID("txtEVServer18").value = "<%=EVServer18%>"
	document.GetElementByID("txtDatacenter18").value = "<%=Datacenter18%>"
	document.GetElementByID("txtEVDatabase18").value = "<%=EVDatabase18%>"
	document.GetElementByID("txtDBServer18").value = "<%=DBServer18%>"
	document.GetElementByID("txtRole18").value = "<%=Role18%>"
	document.GetElementByID("txtMSMQPath18").value = "<%=MSMQPath18%>"
	document.GetElementByID("txtTriggerPath18").value = "<%=TriggerPath18%>"
	
	document.GetElementByID("txtDNSAlias19").value = "<%=DNSAlias19%>"
	document.GetElementByID("txtEVServer19").value = "<%=EVServer19%>"
	document.GetElementByID("txtDatacenter19").value = "<%=Datacenter19%>"
	document.GetElementByID("txtEVDatabase19").value = "<%=EVDatabase19%>"
	document.GetElementByID("txtDBServer19").value = "<%=DBServer19%>"
	document.GetElementByID("txtRole19").value = "<%=Role19%>"
	document.GetElementByID("txtMSMQPath19").value = "<%=MSMQPath19%>"
	document.GetElementByID("txtTriggerPath19").value = "<%=TriggerPath19%>"
	
	document.GetElementByID("txtDNSAlias20").value = "<%=DNSAlias20%>"
	document.GetElementByID("txtEVServer20").value = "<%=EVServer20%>"
	document.GetElementByID("txtDatacenter20").value = "<%=Datacenter20%>"
	document.GetElementByID("txtEVDatabase20").value = "<%=EVDatabase20%>"
	document.GetElementByID("txtDBServer20").value = "<%=DBServer20%>"
	document.GetElementByID("txtRole20").value = "<%=Role20%>"
	document.GetElementByID("txtMSMQPath20").value = "<%=MSMQPath20%>"
	document.GetElementByID("txtTriggerPath20").value = "<%=TriggerPath20%>"
	
	If err.number <> 0 then 
		Call LogEvent("Error occured in LoadValues - Error number: " & err.number & vbcrlf & "Description " & err.description,1)
	End If

End Function

  </script>
<%
'Test to see if the form has been submitted. If it has,
'update the XML file. If not, transform the XML file for
'editing.

If Request.Form("btnSubmit") <> "" Then
  on error goto 0
  call xmlsave %>
  <script language="vbscript">
    Call LoadValues
    document.GetElementByID("savediv").style.display = "none"
  </script><%
End If
%>
  <script language="vbscript">
  Call LoadValues

  </script>

  </form>
</div>
<!--#include file="./footer.asp"-->
<%

Function XMLSave

  set xmlDoc = server.createobject("Microsoft.XMLDOM")
  xmlDoc.async = false
  xmlDoc.load strConfigurationFile
  
  If err.number <> 0 then 
    Call LogEvent("Error occured in XMLSave - Error number: " & err.number & vbcrlf & "Description " & err.description,1)
  End IF

  set ServerConfig = xmlDoc.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")

  ServerConfig.setAttribute("DirectoryServer") = request.form("txtDirectoryServer")
  DirectoryServer = request.form("txtDirectoryServer")

  ServerConfig.setAttribute("EVVersion") = request.form("txtEVVersion")
  EVVersion = request.form("txtEVVersion")
  
  ServerConfig.setAttribute("ConnectionString") = request.form("txtConnectionString")
  ConnectionString = request.form("txtConnectionString")

  ServerConfig.setAttribute("SQLUserName") = request.form("txtSQLUserName")
  SQLUserName = request.form("txtSQLUserName")

  ServerConfig.setAttribute("SQLPassword") = request.form("txtSQLPassword")
  SQLPassword = request.form("txtSQLPassword")
  
  ServerConfig.setAttribute("SQLUseAltPort") = request.form("txtSQLUseAltPort")
  SQLUseAltPort = request.form("txtSQLUseAltPort")  
  
  ServerConfig.setAttribute("SQLAltPortNumber") = request.form("txtAltPortNumber")
  SQLAltPortNumber = request.form("txtAltPortNumber") 

  ServerConfig.setAttribute("Setup") = "false"
  
  '***************************************************************  
  
  ServerConfig.setAttribute("EVDSQLServer") = request.form("txtEVDSQLServer")
  EVDSQLServer = request.form("txtEVDSQLServer")
  
  ServerConfig.setAttribute("EVDSQLDatabase") = request.form("txtEVDSQLDatabase")
  EVDSQLDatabase = request.form("txtEVDSQLDatabase")
  
  ServerConfig.setAttribute("EVDConnectionString") = request.form("txtEVDConnectionString")
  EVDConnectionString = request.form("txtEVDConnectionString")

  ServerConfig.setAttribute("EVDSQLUserName") = request.form("txtEVDSQLUserName")
  EVDSQLUserName = request.form("txtEVDSQLUserName")

  ServerConfig.setAttribute("EVDSQLPassword") = request.form("txtEVDSQLPassword")
  EVDSQLPassword = request.form("txtEVDSQLPassword")
  
  ServerConfig.setAttribute("EVDSQLUseAltPort") = request.form("txtEVDSQLUseAltPort")
  EVDSQLUseAltPort = request.form("txtEVDSQLUseAltPort")  
  
  ServerConfig.setAttribute("EVDSQLAltPortNumber") = request.form("txtEVDSQLAltPortNumber")
  EVDSQLAltPortNumber = request.form("txtEVDSQLAltPortNumber") 
   
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias1") = request.form("txtDNSAlias1")
  DNSAlias1 = request.form("txtDNSAlias1")
  
  ServerConfig.setAttribute("EVServer1") = request.form("txtEVServer1")
  EVServer1 = request.form("txtEVServer1")
  
  ServerConfig.setAttribute("Datacenter1") = request.form("txtDatacenter1")
  Datacenter1 = request.form("txtDatacenter1")
    
  ServerConfig.setAttribute("EVDatabase1") = request.form("txtEVDatabase1")
  EVDatabase1 = request.form("txtEVDatabase1")
  
  ServerConfig.setAttribute("DBServer1") = request.form("txtDBServer1")
  DBServer1 = request.form("txtDBServer1")
  
  ServerConfig.setAttribute("Role1") = request.form("txtRole1")
  Role1 = request.form("txtRole1")
  
  ServerConfig.setAttribute("MSMQPath1") = request.form("txtMSMQPath1")
  MSMQPath1 = request.form("txtMSMQPath1")
  
  ServerConfig.setAttribute("TriggerPath1") = request.form("txtTriggerPath1")
  TriggerPath1 = request.form("txtTriggerPath1")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias2") = request.form("txtDNSAlias2")
  DNSAlias2 = request.form("txtDNSAlias2")
  
  ServerConfig.setAttribute("EVServer2") = request.form("txtEVServer2")
  EVServer2 = request.form("txtEVServer2")
  
  ServerConfig.setAttribute("Datacenter2") = request.form("txtDatacenter2")
  Datacenter2 = request.form("txtDatacenter2")
  
  ServerConfig.setAttribute("EVDatabase2") = request.form("txtEVDatabase2")
  EVDatabase2 = request.form("txtEVDatabase2")
  
  ServerConfig.setAttribute("DBServer2") = request.form("txtDBServer2")
  DBServer2 = request.form("txtDBServer2")
  
  ServerConfig.setAttribute("Role2") = request.form("txtRole2")
  Role2 = request.form("txtRole2")
  
  ServerConfig.setAttribute("MSMQPath2") = request.form("txtMSMQPath2")
  MSMQPath2 = request.form("txtMSMQPath2")
 
  ServerConfig.setAttribute("TriggerPath2") = request.form("txtTriggerPath2")
  TriggerPath2 = request.form("txtTriggerPath2")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias3") = request.form("txtDNSAlias3")
  DNSAlias3 = request.form("txtDNSAlias3")
  
  ServerConfig.setAttribute("EVServer3") = request.form("txtEVServer3")
  EVServer3 = request.form("txtEVServer3")
  
  ServerConfig.setAttribute("Datacenter3") = request.form("txtDatacenter3")
  Datacenter3 = request.form("txtDatacenter3")
  
  ServerConfig.setAttribute("EVDatabase3") = request.form("txtEVDatabase3")
  EVDatabase3 = request.form("txtEVDatabase3")
  
  ServerConfig.setAttribute("DBServer3") = request.form("txtDBServer3")
  DBServer3 = request.form("txtDBServer3")
  
  ServerConfig.setAttribute("Role3") = request.form("txtRole3")
  Role3 = request.form("txtRole3")
  
  ServerConfig.setAttribute("MSMQPath3") = request.form("txtMSMQPath3")
  MSMQPath3 = request.form("txtMSMQPath3")
  
  ServerConfig.setAttribute("TriggerPath3") = request.form("txtTriggerPath3")
  TriggerPath3 = request.form("txtTriggerPath3")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias4") = request.form("txtDNSAlias4")
  DNSAlias4 = request.form("txtDNSAlias4")
  
  ServerConfig.setAttribute("EVServer4") = request.form("txtEVServer4")
  EVServer4 = request.form("txtEVServer4")
  
  ServerConfig.setAttribute("Datacenter4") = request.form("txtDatacenter4")
  Datacenter4 = request.form("txtDatacenter4")
  
  ServerConfig.setAttribute("EVDatabase4") = request.form("txtEVDatabase4")
  EVDatabase4 = request.form("txtEVDatabase4")
  
  ServerConfig.setAttribute("DBServer4") = request.form("txtDBServer4")
  DBServer4 = request.form("txtDBServer4")
  
  ServerConfig.setAttribute("Role4") = request.form("txtRole4")
  Role4 = request.form("txtRole4")
  
  ServerConfig.setAttribute("MSMQPath4") = request.form("txtMSMQPath4")
  MSMQPath4 = request.form("txtMSMQPath4")
  
  ServerConfig.setAttribute("TriggerPath4") = request.form("txtTriggerPath4")
  TriggerPath4 = request.form("txtTriggerPath4")
  
  '***************************************************************  
  
  ServerConfig.setAttribute("DNSAlias5") = request.form("txtDNSAlias5")
  DNSAlias5 = request.form("txtDNSAlias5")
  
  ServerConfig.setAttribute("EVServer5") = request.form("txtEVServer5")
  EVServer5 = request.form("txtEVServer5")
  
  ServerConfig.setAttribute("Datacenter5") = request.form("txtDatacenter5")
  Datacenter5 = request.form("txtDatacenter5")
  
  ServerConfig.setAttribute("EVDatabase5") = request.form("txtEVDatabase5")
  EVDatabase5 = request.form("txtEVDatabase5")
  
  ServerConfig.setAttribute("DBServer5") = request.form("txtDBServer5")
  DBServer5 = request.form("txtDBServer5")
  
  ServerConfig.setAttribute("Role5") = request.form("txtRole5")
  Role5 = request.form("txtRole5")
  
  ServerConfig.setAttribute("MSMQPath5") = request.form("txtMSMQPath5")
  MSMQPath5 = request.form("txtMSMQPath5")
  
  ServerConfig.setAttribute("TriggerPath5") = request.form("txtTriggerPath5")
  TriggerPath5 = request.form("txtTriggerPath5")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias6") = request.form("txtDNSAlias6")
  DNSAlias6 = request.form("txtDNSAlias6")
  
  ServerConfig.setAttribute("EVServer6") = request.form("txtEVServer6")
  EVServer6 = request.form("txtEVServer6")
  
  ServerConfig.setAttribute("Datacenter6") = request.form("txtDatacenter6")
  Datacenter6 = request.form("txtDatacenter6")
  
  ServerConfig.setAttribute("EVDatabase6") = request.form("txtEVDatabase6")
  EVDatabase6 = request.form("txtEVDatabase6")
  
  ServerConfig.setAttribute("DBServer6") = request.form("txtDBServer6")
  DBServer6 = request.form("txtDBServer6")
  
  ServerConfig.setAttribute("Role6") = request.form("txtRole6")
  Role6 = request.form("txtRole6")
  
  ServerConfig.setAttribute("MSMQPath6") = request.form("txtMSMQPath6")
  MSMQPath6 = request.form("txtMSMQPath6")
  
  ServerConfig.setAttribute("TriggerPath6") = request.form("txtTriggerPath6")
  TriggerPath6 = request.form("txtTriggerPath6")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias7") = request.form("txtDNSAlias7")
  DNSAlias7 = request.form("txtDNSAlias7")
  
  ServerConfig.setAttribute("EVServer7") = request.form("txtEVServer7")
  EVServer7 = request.form("txtEVServer7")
  
  ServerConfig.setAttribute("Datacenter7") = request.form("txtDatacenter7")
  Datacenter7 = request.form("txtDatacenter7")
  
  ServerConfig.setAttribute("EVDatabase7") = request.form("txtEVDatabase7")
  EVDatabase7 = request.form("txtEVDatabase7")
  
  ServerConfig.setAttribute("DBServer7") = request.form("txtDBServer7")
  DBServer7 = request.form("txtDBServer7")
  
  ServerConfig.setAttribute("Role7") = request.form("txtRole7")
  Role7 = request.form("txtRole7")
  
  ServerConfig.setAttribute("MSMQPath7") = request.form("txtMSMQPath7")
  MSMQPath7 = request.form("txtMSMQPath7")
  
  ServerConfig.setAttribute("TriggerPath7") = request.form("txtTriggerPath7")
  TriggerPath7 = request.form("txtTriggerPath7")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias8") = request.form("txtDNSAlias8")
  DNSAlias8 = request.form("txtDNSAlias8")
  
  ServerConfig.setAttribute("EVServer8") = request.form("txtEVServer8")
  EVServer8 = request.form("txtEVServer8")
  
  ServerConfig.setAttribute("Datacenter8") = request.form("txtDatacenter8")
  Datacenter8 = request.form("txtDatacenter8")
  
  ServerConfig.setAttribute("EVDatabase8") = request.form("txtEVDatabase8")
  EVDatabase8 = request.form("txtEVDatabase8")
  
  ServerConfig.setAttribute("DBServer8") = request.form("txtDBServer8")
  DBServer8 = request.form("txtDBServer8")
  
  ServerConfig.setAttribute("Role8") = request.form("txtRole8")
  Role8 = request.form("txtRole8")
  
  ServerConfig.setAttribute("MSMQPath8") = request.form("txtMSMQPath8")
  MSMQPath8 = request.form("txtMSMQPath8")
  
  ServerConfig.setAttribute("TriggerPath8") = request.form("txtTriggerPath8")
  TriggerPath8 = request.form("txtTriggerPath8")
 
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias9") = request.form("txtDNSAlias9")
  DNSAlias9 = request.form("txtDNSAlias9")
  
  ServerConfig.setAttribute("EVServer9") = request.form("txtEVServer9")
  EVServer9 = request.form("txtEVServer9")
  
  ServerConfig.setAttribute("Datacenter9") = request.form("txtDatacenter9")
  Datacenter9 = request.form("txtDatacenter9")
  
  ServerConfig.setAttribute("EVDatabase9") = request.form("txtEVDatabase9")
  EVDatabase9 = request.form("txtEVDatabase9")
  
  ServerConfig.setAttribute("DBServer9") = request.form("txtDBServer9")
  DBServer9 = request.form("txtDBServer9")
  
  ServerConfig.setAttribute("Role9") = request.form("txtRole9")
  Role9 = request.form("txtRole9")
  
  ServerConfig.setAttribute("MSMQPath9") = request.form("txtMSMQPath9")
  MSMQPath9 = request.form("txtMSMQPath9")
  
  ServerConfig.setAttribute("TriggerPath9") = request.form("txtTriggerPath9")
  TriggerPath9 = request.form("txtTriggerPath9")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias10") = request.form("txtDNSAlias10")
  DNSAlias10 = request.form("txtDNSAlias10")
  
  ServerConfig.setAttribute("EVServer10") = request.form("txtEVServer10")
  EVServer10 = request.form("txtEVServer10")
  
  ServerConfig.setAttribute("Datacenter10") = request.form("txtDatacenter10")
  Datacenter10 = request.form("txtDatacenter10")
  
  ServerConfig.setAttribute("EVDatabase10") = request.form("txtEVDatabase10")
  EVDatabase10 = request.form("txtEVDatabase10")
  
  ServerConfig.setAttribute("DBServer10") = request.form("txtDBServer10")
  DBServer10 = request.form("txtDBServer10")
  
  ServerConfig.setAttribute("Role10") = request.form("txtRole10")
  Role10 = request.form("txtRole10")
  
  ServerConfig.setAttribute("MSMQPath10") = request.form("txtMSMQPath10")
  MSMQPath10 = request.form("txtMSMQPath10")
  
  ServerConfig.setAttribute("TriggerPath10") = request.form("txtTriggerPath10")
  TriggerPath10 = request.form("txtTriggerPath10")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias11") = request.form("txtDNSAlias11")
  DNSAlias11 = request.form("txtDNSAlias11")
  
  ServerConfig.setAttribute("EVServer11") = request.form("txtEVServer11")
  EVServer11 = request.form("txtEVServer11")
  
  ServerConfig.setAttribute("Datacenter11") = request.form("txtDatacenter11")
  Datacenter11 = request.form("txtDatacenter11")
  
  ServerConfig.setAttribute("EVDatabase11") = request.form("txtEVDatabase11")
  EVDatabase11 = request.form("txtEVDatabase11")
  
  ServerConfig.setAttribute("DBServer11") = request.form("txtDBServer11")
  DBServer11 = request.form("txtDBServer11")
  
  ServerConfig.setAttribute("Role11") = request.form("txtRole11")
  Role11 = request.form("txtRole11")
  
  ServerConfig.setAttribute("MSMQPath11") = request.form("txtMSMQPath11")
  MSMQPath11 = request.form("txtMSMQPath11")
  
  ServerConfig.setAttribute("TriggerPath11") = request.form("txtTriggerPath11")
  TriggerPath11 = request.form("txtTriggerPath11")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias12") = request.form("txtDNSAlias12")
  DNSAlias12 = request.form("txtDNSAlias12")
  
  ServerConfig.setAttribute("EVServer12") = request.form("txtEVServer12")
  EVServer12 = request.form("txtEVServer12")
  
  ServerConfig.setAttribute("Datacenter12") = request.form("txtDatacenter12")
  Datacenter12 = request.form("txtDatacenter12")
  
  ServerConfig.setAttribute("EVDatabase12") = request.form("txtEVDatabase12")
  EVDatabase12 = request.form("txtEVDatabase12")
  
  ServerConfig.setAttribute("DBServer12") = request.form("txtDBServer12")
  DBServer12 = request.form("txtDBServer12")
  
  ServerConfig.setAttribute("Role12") = request.form("txtRole12")
  Role12 = request.form("txtRole12")
  
  ServerConfig.setAttribute("MSMQPath12") = request.form("txtMSMQPath12")
  MSMQPath12 = request.form("txtMSMQPath12")
  
  ServerConfig.setAttribute("TriggerPath12") = request.form("txtTriggerPath12")
  TriggerPath12 = request.form("txtTriggerPath12")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias13") = request.form("txtDNSAlias13")
  DNSAlias13 = request.form("txtDNSAlias13")
  
  ServerConfig.setAttribute("EVServer13") = request.form("txtEVServer13")
  EVServer13 = request.form("txtEVServer13")
  
  ServerConfig.setAttribute("Datacenter13") = request.form("txtDatacenter13")
  Datacenter13 = request.form("txtDatacenter13")
  
  ServerConfig.setAttribute("EVDatabase13") = request.form("txtEVDatabase13")
  EVDatabase13 = request.form("txtEVDatabase13")
  
  ServerConfig.setAttribute("DBServer13") = request.form("txtDBServer13")
  DBServer13 = request.form("txtDBServer13")
  
  ServerConfig.setAttribute("Role13") = request.form("txtRole13")
  Role13 = request.form("txtRole13")
  
  ServerConfig.setAttribute("MSMQPath13") = request.form("txtMSMQPath13")
  MSMQPath13 = request.form("txtMSMQPath13")
  
  ServerConfig.setAttribute("TriggerPath13") = request.form("txtTriggerPath13")
  TriggerPath13 = request.form("txtTriggerPath13")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias14") = request.form("txtDNSAlias14")
  DNSAlias14 = request.form("txtDNSAlias14")
  
  ServerConfig.setAttribute("EVServer14") = request.form("txtEVServer14")
  EVServer14 = request.form("txtEVServer14")
  
  ServerConfig.setAttribute("Datacenter14") = request.form("txtDatacenter14")
  Datacenter14 = request.form("txtDatacenter14")
  
  ServerConfig.setAttribute("EVDatabase14") = request.form("txtEVDatabase14")
  EVDatabase14 = request.form("txtEVDatabase14")
  
  ServerConfig.setAttribute("DBServer14") = request.form("txtDBServer14")
  DBServer14 = request.form("txtDBServer14")
  
  ServerConfig.setAttribute("Role14") = request.form("txtRole14")
  Role14 = request.form("txtRole14")
  
  ServerConfig.setAttribute("MSMQPath14") = request.form("txtMSMQPath14")
  MSMQPath14 = request.form("txtMSMQPath14")
  
  ServerConfig.setAttribute("TriggerPath14") = request.form("txtTriggerPath14")
  TriggerPath14 = request.form("txtTriggerPath14")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias15") = request.form("txtDNSAlias15")
  DNSAlias15 = request.form("txtDNSAlias15")
  
  ServerConfig.setAttribute("EVServer15") = request.form("txtEVServer15")
  EVServer15 = request.form("txtEVServer15")
  
  ServerConfig.setAttribute("Datacenter15") = request.form("txtDatacenter15")
  Datacenter15 = request.form("txtDatacenter15")
  
  ServerConfig.setAttribute("EVDatabase15") = request.form("txtEVDatabase15")
  EVDatabase15 = request.form("txtEVDatabase15")
  
  ServerConfig.setAttribute("DBServer15") = request.form("txtDBServer15")
  DBServer15 = request.form("txtDBServer15")
  
  ServerConfig.setAttribute("Role15") = request.form("txtRole15")
  Role15 = request.form("txtRole15")
  
  ServerConfig.setAttribute("MSMQPath15") = request.form("txtMSMQPath15")
  MSMQPath15 = request.form("txtMSMQPath15")
  
  ServerConfig.setAttribute("TriggerPath15") = request.form("txtTriggerPath15")
  TriggerPath15 = request.form("txtTriggerPath15")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias16") = request.form("txtDNSAlias16")
  DNSAlias16 = request.form("txtDNSAlias16")
  
  ServerConfig.setAttribute("EVServer16") = request.form("txtEVServer16")
  EVServer16 = request.form("txtEVServer16")
  
  ServerConfig.setAttribute("Datacenter16") = request.form("txtDatacenter16")
  Datacenter16 = request.form("txtDatacenter16")
  
  ServerConfig.setAttribute("EVDatabase16") = request.form("txtEVDatabase16")
  EVDatabase16 = request.form("txtEVDatabase16")
  
  ServerConfig.setAttribute("DBServer16") = request.form("txtDBServer16")
  DBServer16 = request.form("txtDBServer16")
  
  ServerConfig.setAttribute("Role16") = request.form("txtRole16")
  Role16 = request.form("txtRole16")
  
  ServerConfig.setAttribute("MSMQPath16") = request.form("txtMSMQPath16")
  MSMQPath16 = request.form("txtMSMQPath16")
  
  ServerConfig.setAttribute("TriggerPath16") = request.form("txtTriggerPath16")
  TriggerPath16 = request.form("txtTriggerPath16")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias17") = request.form("txtDNSAlias17")
  DNSAlias17 = request.form("txtDNSAlias17")
  
  ServerConfig.setAttribute("EVServer17") = request.form("txtEVServer17")
  EVServer17 = request.form("txtEVServer17")
  
  ServerConfig.setAttribute("Datacenter17") = request.form("txtDatacenter17")
  Datacenter17 = request.form("txtDatacenter17")
  
  ServerConfig.setAttribute("EVDatabase17") = request.form("txtEVDatabase17")
  EVDatabase17 = request.form("txtEVDatabase17")
  
  ServerConfig.setAttribute("DBServer17") = request.form("txtDBServer17")
  DBServer17 = request.form("txtDBServer17")
  
  ServerConfig.setAttribute("Role17") = request.form("txtRole17")
  Role17 = request.form("txtRole17")
  
  ServerConfig.setAttribute("MSMQPath17") = request.form("txtMSMQPath17")
  MSMQPath17 = request.form("txtMSMQPath17")
  
  ServerConfig.setAttribute("TriggerPath17") = request.form("txtTriggerPath17")
  TriggerPath17 = request.form("txtTriggerPath17")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias18") = request.form("txtDNSAlias18")
  DNSAlias18 = request.form("txtDNSAlias18")
  
  ServerConfig.setAttribute("EVServer18") = request.form("txtEVServer18")
  EVServer18 = request.form("txtEVServer18")
  
  ServerConfig.setAttribute("Datacenter18") = request.form("txtDatacenter18")
  Datacenter18 = request.form("txtDatacenter18")
  
  ServerConfig.setAttribute("EVDatabase18") = request.form("txtEVDatabase18")
  EVDatabase18 = request.form("txtEVDatabase18")
  
  ServerConfig.setAttribute("DBServer18") = request.form("txtDBServer18")
  DBServer18 = request.form("txtDBServer18")
  
  ServerConfig.setAttribute("Role18") = request.form("txtRole18")
  Role18 = request.form("txtRole18")
  
  ServerConfig.setAttribute("MSMQPath18") = request.form("txtMSMQPath18")
  MSMQPath18 = request.form("txtMSMQPath18")
  
  ServerConfig.setAttribute("TriggerPath18") = request.form("txtTriggerPath18")
  TriggerPath18 = request.form("txtTriggerPath18")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias19") = request.form("txtDNSAlias19")
  DNSAlias19 = request.form("txtDNSAlias19")
  
  ServerConfig.setAttribute("EVServer19") = request.form("txtEVServer19")
  EVServer19 = request.form("txtEVServer19")
  
  ServerConfig.setAttribute("Datacenter19") = request.form("txtDatacenter19")
  Datacenter19 = request.form("txtDatacenter19")
  
  ServerConfig.setAttribute("EVDatabase19") = request.form("txtEVDatabase19")
  EVDatabase19 = request.form("txtEVDatabase19")
  
  ServerConfig.setAttribute("DBServer19") = request.form("txtDBServer19")
  DBServer19 = request.form("txtDBServer19")
  
  ServerConfig.setAttribute("Role19") = request.form("txtRole19")
  Role19 = request.form("txtRole19")
  
  ServerConfig.setAttribute("MSMQPath19") = request.form("txtMSMQPath19")
  MSMQPath19 = request.form("txtMSMQPath19")
  
  ServerConfig.setAttribute("TriggerPath19") = request.form("txtTriggerPath19")
  TriggerPath19 = request.form("txtTriggerPath19")
  
  '***************************************************************
  
  ServerConfig.setAttribute("DNSAlias20") = request.form("txtDNSAlias20")
  DNSAlias20 = request.form("txtDNSAlias20")
  
  ServerConfig.setAttribute("EVServer20") = request.form("txtEVServer20")
  EVServer20 = request.form("txtEVServer20")
  
  ServerConfig.setAttribute("Datacenter20") = request.form("txtDatacenter20")
  Datacenter20 = request.form("txtDatacenter20")
  
  ServerConfig.setAttribute("EVDatabase20") = request.form("txtEVDatabase20")
  EVDatabase20 = request.form("txtEVDatabase20")
  
  ServerConfig.setAttribute("DBServer20") = request.form("txtDBServer20")
  DBServer20 = request.form("txtDBServer20")
  
  ServerConfig.setAttribute("Role20") = request.form("txtRole20")
  Role20 = request.form("txtRole20")
  
  ServerConfig.setAttribute("MSMQPath20") = request.form("txtMSMQPath20")
  MSMQPath20 = request.form("txtMSMQPath20")
  
  ServerConfig.setAttribute("TriggerPath20") = request.form("txtTriggerPath20")
  TriggerPath20 = request.form("txtTriggerPath20")
  
  '***************************************************************
  
  xmlDoc.save strConfigurationFile

  if err.number <> "0" Then
    response.write "<span style='float:middle;text-align:center;'><img src='./images/ico_error.gif' style='vertical-align:top'>&nbsp;Error saving configuration.  Check your permissions to the configuration.xml file.</span>"
    Call LogEvent("Error occured in XMLSave - Error number: " & err.number & vbcrlf & "Description " & err.description,1)
    exit function
  Else
    response.write "<meta http-equiv='refresh' content=3;url='./admin.asp'/><br><span style='width:100%;float:middle;text-align:center;'><img style='vertical-align:top' src='./images/ico_success.gif'>&nbsp;Configuration saved.</span><br><br>"
    response.write "<span style='font-size:85%;float:middle;text-align:center;'><a href='./admin.asp'>Return to admin settings</a></span>"
  End If

  set xmlDoc = nothing

End Function

%>
