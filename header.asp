<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

<% 
EVDashboardVersion = "EV Dashboard 0.24"
%>

<head>
	<title><%=EVDashboardVersion%> - https://sourceforge.net/projects/evdashboard/</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<link rel="stylesheet" type="text/css" href="./css/main.css" />
	<script type="text/javascript" src="./js/menu.js"></script>
	<script language="JavaScript" src="./js/FusionCharts.js"></script>
</head>

<body>
<div id="header">
	<div class="leftcol">
		<div id="logo1"></div>
	</div>
	<div class="rightcol">
		<div class="head_bg">
			<div id="logo2_right">
				<img src="images/EVDashboardLogo.png" alt="EV Dashboard"/>
			</div>
		</div>
	</div>
</div>

<div class="splitmenu" id="chromemenu">
	<ul>
		<li><a href="index.asp">Home</a><img src="images/nav_item.png"></li>
		<li><a href="" rel="dropmenu1">Health Checks</a><img src="images/nav_item.png"></li>
		<li><a href="" rel="dropmenu2">Mailbox Archive Reports</a><img src="images/nav_item.png"></li>
		<li><a href="" rel="dropmenu3">Exchange Server Reports</a><img src="images/nav_item.png"></li>
		<li><a href="" rel="dropmenu4">File Archive Reports</a><img src="images/nav_item.png"></li>
		<li><a href="" rel="dropmenu5">Registry Keys</a><img src="images/nav_item.png"></li>
		<li><a href="" rel="dropmenu6">Links</a><img src="images/nav_item.png"></li>
		<li><a href="" rel="dropmenu7">Help</a></li>
	</ul>
</div>

<!-- Health Checks Menu -->                                                   
<div id="dropmenu1" class="dropmenudiv">
	<a href="msmqmessagecount.asp"><img src="images/msmq.gif">MSMQs - Message Counts</a>
	<a href="msmqfoldersizes.asp"><img src="images/msmq.gif">MSMQs - Folder Sizes</a>
	<a href="pingtest.asp"><img src="images/server.png">Ping Test</a>
	<a href="evservices.asp"><img src="images/cog.png">EV Services</a>
	<a href="EVdiskusageinfo.asp"><img src="images/server.png">EV Disk Usage Information</a>
	<a href="HourlyArchivingRates.asp"><img src="images/database_refresh.png">Hourly Archiving Rates</a>	
	<a href="OrganisationArchivingHistory.asp"><img src="images/database_key.png">Organisation Archiving History</a>	
	<a href="EVServerEventLogs.asp"><img src="images/magnifier.png">EV Server Event Logs</a>
</div>

<!-- Mailbox Archive Reports Menu -->                                                
<div id="dropmenu2" class="dropmenudiv" style="width:220px">
	<a href="EnabledMailboxes.asp"><img src="images/group.png">Enabled Mailboxes (List)</a>
	<a href="EnabledMailboxesLite.asp"><img src="images/group.png">Enabled Mailboxes Lite (List)</a> 
	<a href="DisabledUsers.asp"><img src="images/group.png">Disabled Users (List)</a>
	<a href="OrphanedArchives.asp"><img src="images/disconnect.png">Orphaned Archives (List)</a>
	<a href="EnabledMailboxesCounts.asp"><img src="images/database_key.png">Enabled Mailboxes (Count)</a>
	<a href="ArchivesGrowth.asp"><img src="images/chart_bar.png">Archives Growth (Count)</a>
	<a href="ArchiveStates.asp"><img src="images/database_go.png">Archive States (Count)</a>
	<a href="UsageReportByArchive.asp"><img src="images/group.png">Usage Report By Archive (List)</a>
	<a href="UsageReportByArchiveByYear.asp"><img src="images/group.png">Usage Report By Archive By Year (List)</a>
	<a href="SISReport.asp"><img src="images/database_gear.png">Single Instance Storage Report</a>
	<a href="PSTReport.asp"><img src="images/database_connect.png">PST Migration Report</a>
</div>

<!-- Exchange Server Reports Menu -->                                                   
<div id="dropmenu3" class="dropmenudiv" style="width:250px">
	<a href="ExchangeMailboxes.asp"><img src="images/chart_bar.png">Exchange Mailboxes (By Server)</a>
	<a href="ExchangeMailstores.asp"><img src="images/email_go.png">Exchange Mailboxes (By Mailstore)</a>
	<a href="Exchangediskusageinfo.asp"><img src="images/server_lightning.png">Exchange Server Disk Usage Information</a>
	<a href="ExchangeLegacyDNs.asp"><img src="images/database_save.png">Exchange Legacy DNs (List)</a>
	<a href="ExchangeMailboxstates.asp"><img src="images/database_edit.png">Exchange Mailbox States (Count)</a>
	<a href="OverWarningLimit.asp"><img src="images/group_key.png">Over Warning Limit (List)</a>
	<a href="OverSendLimit.asp"><img src="images/group_error.png">Over Send Limit (List)</a>
	<a href="OverReceiveLimit.asp"><img src="images/group_delete.png">Over Receive Limit (List)</a>
</div>


<!-- File Archive Reports Menu -->                                                   
<div id="dropmenu4" class="dropmenudiv">
	<a href="FSADailyArchiveRates.asp"><img src="images/database_refresh.png">Daily Archive Rates</a>
	<a href="FSATargets.asp"><img src="images/server.png">File Server Archive Targets</a>
</div>

<!-- Registry Keys Menu -->                                                   
<div id="dropmenu5" class="dropmenudiv">
	<a href="RegistryKeysKVSStorage.asp"><img src="images/keys.png">7.5 KVS Storage Registry Keys</a>
</div>

<!-- Links Menu -->                                                   
<div id="dropmenu6" class="dropmenudiv" style="width:250px">
	<a href="https://sourceforge.net/projects/evdashboard/" target="_blank"><img src="images/package_link.png">Project Webpage</a>
	<a href="https://www-secure.symantec.com/connect" target="_blank"><img src="images/package_link.png">Symantec Connect</a>
	<a href="http://www.publicshareware.com/public-share-outlook-utilities.php" target="_blank"><img src="images/package_link.png">DocMessageClass - Shareware</a>
</div>

<!-- Help Menu -->                                                   
<div id="dropmenu7" class="dropmenudiv">
	<a href="admin.asp"><img src="images/cog.png">Configure EVDashboard</a>
	<a href="TestDBConnectivity.asp"><img src="images/connect.png">Test Database Connectivity</a>
	<a href="DatabaseStats.asp"><img src="images/connect.png">EV Dashboard Database Statistics</a>
	<a href="Upgrade.asp"><img src="images/information.png">EV Dashboard Upgrade Notes</a>
	<a href="ChangeLog.asp"><img src="images/page_code.png">Change Log</a>
	<a href="Feedback.asp"><img src="images/bug.png">Feedback</a>
	<a href="UsefulInfo.asp"><img src="images/information.png">Useful Information</a>
	<a href="FilterOperators.asp"><img src="images/information.png">Filter Operators</a>
	<a href="Notepad++.asp"><img src="images/information.png">Notepad++</a>
	<a href="about.asp"><img src="images/information.png">About</a>
</div>
<script type="text/javascript">
	cssdropdown.startchrome("chromemenu")
</script>
<br>