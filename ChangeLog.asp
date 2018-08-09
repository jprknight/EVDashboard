<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<br><br>
<h1>Change Log</h1>
<br><br>
<br>
<h2>Version 0.24 - 10th March 2011</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Bug Fix:3176839: Updated SISReport to provide apostrophe fix.</li>
<li>Bug Fix:3178061: Updated disk reports to show mounted volumes. Credit to drawling!</li>
<li>Added feature request:3177910: PST Report, with export option. Credit to drawling!</li>
<li>Added feature request:3180190: FSA Report. Credit to drawling!</li>
<li>Added alternative SQL port functionality from support request: 3197820.</li>
<li>Added code to begin EV 9.0 database schema changes support. Selectable option in admin.asp.</li>
<li>Updated OverWarningLimit.asp, OverSendLimt.asp and OverReceiveLimit.asp with new SQL query for above schema change.</li>
</ul>
</h4>
<br>
<h2>Version 0.23 - 10th February 2011</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Removed evdiscuss.net branding and links, replaced with sourceforge links.</li>
<li>Added Eric Rife's updated installation document. This is also available to download from <a href="https://sourceforge.net/projects/evdashboard/" target="_blank">https://sourceforge.net/projects/evdashboard/</a></li>
</ul>
</h4>
<br>
<h2>Version 0.22 - 8th February 2011</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Addition: inc_SEV_EventLogs.vbs - This script scrapes the Symantec Enterprise Vault event log for warnings and errors. These entries are put into the same table for inc_EV_Server_EventLogs.vbs, so also appear on the front page.</li>
<li>Re-shuffled menu items.</li>
<li>Addition: Added Backup Trigger File information to the frontpage. Removed as a seperate menu item.</li>
<li>Addition: Added Journal Archive Count to frontpage.</li>
</ul>
</h4>
<br>
<h2>Version 0.21 - 18th October 2010</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Change: EnabledMailboxesLite.asp & EnabledMailboxes.asp - Added 5000 & 10000 to items per page. Sorting will only sort on the currently viewable records. Added these options for larger sites.</li>
<li>Change: IndexBuildXML.asp - Instead of counting records by EV server, changed them to be counted by EV Vault Storage Database. This reflects archive load better.</li>
<li>Addition: Single Instance Storage Report. You can now work out which archives are using the most storage and what percentage of those archives are sharing with SIS. Thanks go to JesusWept2 for providing the underlining SQL query.</li>
<li>Fix: Minor fixes to the appearance of EVDiskUsageInfo.asp & ExchangeDiskUsageInfo.asp.</li>
<li>Addition: Added export file SISReportExport.asp to allow downloading the new report into Excel.</li>
<li>Change: Removed graph from ArhiveGrowth.asp; once a large amount of data has been recorded into the EVDashboard database the graph becomes really slow to run.</li>
<li>Fix: Added rotateNames='1' to line 14 of IndexBuildXML.asp. So the names of columns are vertical and do not overlap each other on the EV Mailbox Archiving Users graph.</li>
</ul>
</h4>
<br>
<br>
<h2>Version 0.20 - 23rd September 2010</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Added extra output to logging within inc_EV_Server_EventLogs.vbs.</li>
<li>Removed Filter button from UsageReportByArchive.asp, and enabled 'Filter as you type'.</li>
<li>Removed Filter button from UsageReportByArchiveByYear.asp, and enabled 'Filter as you type'.</li>
<li>Fixed IndexDiskAlerts.asp, added exclusion for rows not recorded correctly.</li>
<li>Added date selector to EnabledMailboxesCounts.asp.</li>
<li>Added ON ERROR RESUME NEXT to inc_Enabled_Mailboxes.vbs and inc_Enabled_Mailboxes_Databases.vbs to allow for network disconnections on volume data retrievals.</li>
<li>Added logging levels to the EVDashboard Data Collector script. This can help with diagnosing connection problems if data is not what you expect.</li>
<li>Fixed inc_Enabled_Mailboxes_Databases.asp - Added validation check for DefaultVaultID before running SQL queries. This means some database rows will not have # item in archive, archive size etc, but the whole process is not trashed. The section of code now runs and completes if records without a DefaultVaultID in the ExchangeMailboxEntry table exist.</li>
</ul>
</h4>
<br>
<h2>Version 0.19 - 14th April 2010</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Updated EVDiskUsage.asp - Added 'Filter as you type'.</li>
<li>Updated EXchangeDiskUsage.asp - Added 'Filter as you type'.</li>
<li><u>New table: EVServerEventLogs</u> - Added new table EVServerEventLogs, EVServerEventLogs.asp, inc_EV_Server_EventLogs.vbs. The same data view has also been added to index.asp, by adding inc_EVServerEventLogs.asp. See the file /sql/EVDashboard Database Creation/CreateTables.sql for the code to add this new table into your EV Dashboard database.</li><br>
This new code adds a data view for all EV related events logged within the 24 hours prior to when you run the data collector. This can pick up things like backup alerts etc. Time Written refers to when the data was written into the corresponding event log, not when the data was written into the EV Dashboard database. On the index.asp page only warnings and errors are shown.
<li>Updated IndexDiskAlerts.asp - Changed around code as results could have been inaccurate using SQL script to determine which rows to print to screen. The code now looks for disks which have a percentage free of less than or equal to 25%, and less than 150GB free. Hopefully this should aviod false positives.</li>
</ul>
</h4>
<br>
<h2>Version 0.18 - 26th March 2010</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Added FilterOperators.asp. An information page from <a href="http://tablefilter.free.fr/" target="_blank">http://tablefilter.free.fr/</a> detailing the operators which can be used for filtering the data contained within the reports.</li>
<li>Experimented with column resizer module from tablefilter.free.fr. Seems to be more trouble than it is worth with changing screen resolutions and the format of the tables in use.</li>
<li>Updated inc_Enabled_Mailboxes_Databases.vbs - small change to show count of the number of records when run in interactive mode.</li>
<li>Fixed - EnabledMailboxesLite.asp. The post method went back to EnabledMailboxes.asp, this has been fixed.</li>
<li>Created EnabledMailboxesLiteExport.asp. EnabledMailboxesLite.asp can now be exported to Excel.</li>
</ul>
</h4>
<br>
<h2>Version 0.17 - 15th March 2010</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Fixed inc_Enabled_Mailboxes_Databases.vbs - strMailboxSizeMB was not set, resulting in total mail size being incorrect.</li>
<li>Updated EVDashboard_data_Collector.vbs - Added further output information, written to log and screen. Added If...Then...condition so start and finish time is only written to config.xml if the script was started before 8:30am.</li>
<li>Updated IndexExchangeEnvironment.asp - Added required code for the domain name to be displayed in this section on the index.asp page. Part of this was adding inc_Domain_Name.vbs. Additionally added code to get ServiceAccount; the account the data collector is run as. This is not displayed anywhere yet though.</li>
<li>Added recorded timestamp into the Disk Alerts section of index.asp. Actual changes in IndexDiskAlerts.asp.</li>
<li>Fixed bug 00001: Added new Record TimeStamp formatting function into global.asp. With this dates are now displayed into standard ISO format, regardless of the locale (yyyy/mm/dd hh:mm). Applied this fix to the following files:</li><br>
IndexServiceAlerts.asp<br>
IndexDiskAlerts.asp<br>
OverSendLimit.asp<br>
OverWarningLimit.asp<br>
inc_Enabled_Mailboxes_Databases.vbs<br>
inc_Over_Receive_Limit.vbs<br>
inc_Over_Send_Limit.vbs<br>
inc_Over_Warning_Limit.vbs<br>
inc_Hourly_Archiving_Rates.vbs<br>
<li>Updated EVDashboard_Data_Collector.vbs - Extended strSQLExecute value to 1800 seconds or 30 minutes. Some larger sites may see scripts terminating with a timeout because of high volumes of records retrieved by some SQL queries. Typically inc_Over_Warning_Limit.vbs can retrieve high numbers of records and take more than the previous limit of 5 minutes.</li>
<li>Updated Admin.asp - Moved styles and coding section out to ./includes/AdminStyles.asp for easier reading.</li>
<li>Fixed CSS in Admin.asp - Removed table borders, which were unintentionally set.</li>
<li>Updated inc_Over_Warning_Limit.vbs, inc_Over_Send_Limit.vbs, inc_Over_Receive_Limit.vbs so the count of records and mailbox names are written to screen when the script is run interactively.</li>
<li>Added EnabledMailboxesLite.asp - A cut down version of EnabledMailboxes.asp. Giving slightly less information on screen. Hopefully this will make it easier for people with smaller screen resolutions.</li>
<li>Enabled 'Filter As You Type' within TableFilter.js script. With this change, removed the Filter button which was out of place anyway.</li>
Change applied to:<br>
OverWarningLimit.asp<br>
OverSendLimit.asp<br>
OverReceiveLimit.asp<br>
EnabledMailboxes.asp<br>
EnabledMailboxesLite.asp<br>
DisabledUsers.asp
<li>Updated tablefilter.js and filtergrid.css from version 1.9.4 to 1.9.6.</li>
<li>Added the Column visibility manager and Filter row visibility scripts to:</li><br>
OverWarningLimit.asp<br>
OverSendLimit.asp<br>
OverReceiveLimit.asp<br>
DisabledUsers.asp<br>
<li>Added only Filter row visibility script to:</li><br>
EnabledMailboxes.asp<br>
EnabledMailboxesLite.asp<br>
On larger installations (potentially 5000 EV archives and upwards) enabling the Column visibility manager on these two reports extends the load time of the script to the point where the browser process consumes excessive amounts of memory and CPU power for an excessive amount of time. It appears the script fails to load in this scenario. For this reason these two reports only have the Filter row visibility script turned on.
<li>Fixed number of records per page values to 100,500,1000 on all user report pages.</li>
<li>Enabled 'Load Filters on Demand' for slightly faster loading times on:</li>
EnabledMailboxes.asp<br>
EnabledMailboxesLite.asp<br>
DisabledUsers.asp<br>
OverWarningLimit.asp<br>
OverSendLimit.asp<br>
OverReceiveLimit.asp<br>
</h4>
</ul>
<br>
<h2>Version 0.16 - 17th February 2010</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Updated installation and upgrade documentation.</li>
<li>Re-arranged the order of the scripts which are run by the data collector; So scripts which call data other than from the EVDashboard database run last. For example inc_Exchange_Disk_Info.vbs, which could potentially fail because of network connectivity or WMI problems.</li>
<li>Updated with the application title logo</li>
<li>Updated IndexWatchFileCounts.asp with new SQL connect code.</li>
<li><u>New table: UsageReportsByArchive</u> Report on the number of items which have been archived into each archive. Thanks to JesusWept2 for the underlining SQL query. Added the following ASP pages:</li><br>
UsageReportByArchive.asp<br>
UsageReportByArchiveByYear.asp<br>
The only gotcha with this report is it connects to the ExchangeMailboxEntry table, so only archives with Exchange mailboxes will be listed.<br>
See /sql/EVDashboard Database Creation/CreateTables.sql for the code to add this table to your EV Dashboard database.
<li>Updated Data_Collector.vbs - Updated User Configuration variable based on time of day. This script can now be run within the same day without editing the file. If the script is run after 8:30am the script runs a reduced set of sub-scripts.</li><br>
inc_Build_MultidimentionalArray.vbs<br>
inc_Service_Alerts.vbs<br>
inc_MSMQMessageCountAlerts.vbs<br>
inc_EV_Disk_Info.vbs<br>
inc_Exchange_Disk_Info.vbs
<li>Updated ExchangeMailboxes.asp - Added date selector to page.</li>
<li>Updated ExchangeMailstores.asp - Added date selector to page.</li>
<li>Fixed EnabledMailboxes.asp - Corrected SQL query 2 SQLConnect timeout; it now uses the global strSQLConnect, by default 300 seconds. This should resolve problems of this script stopping half way through, and interupting the main data collector script.</li>
</ul>
</h4>
<br>
<h3>Known Bugs:</h3>
<br>
<ul>
<li><h4>00001 - Date conversation within EVDashboard_Data_Collector.vbs, temporary work-around in place.<br>
For some reason when timestamps are selected with adodb they change according to regional settings.<br>
Currently date format is for UK (DD.MM.YYYY). Code changes will be needed for US style dates (MM.DD.YYYY).</li></h4>
</ul>
<br>
<h2>Version 0.15 - 18th January 2010</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Updated SQL insert statements so data collector can run more than once on the same day in:<br>
inc_Enabled_Users.vbs<br>
inc_Disabled_Users.vbs<br>
inc_EV_Disk_Info.vbs<br>
inc_Exchange_Disk_Info.vbs<br>
inc_Exchange_Mailboxes.vbs<br>
inc_MSMQMessageCountAlerts.vbs<br>
inc_Service_Alerts.vbs<br>
inc_Over_Receive_Limit.vbs<br>
inc_Over_Send_Limit.vbs<br>
inc_Over_Warning_Limit.vbs<br>
inc_Hourly_Archiving_Rates.vbs<br>
inc_Organisation_Archiving_History.vbs<br><br>
</li>
<li>Fixed inc_Enabled_Users.asp to include apostrophe fix in linked servers section. - Thanks to Paul Bendall for noticing this.</li>
<li>Fixed database server name issue which could come up when using SQL authentication within inc_Organisation_Archiving_History.vbs - Timeout Fix</li>
<li>Removed upgrade SQL scripts. From now on any new tables will be annouced in this changelog.</li>
<li>Updated the mechanism for collecting user / mailbox / archive data.</li><br>
DB table EnabledMailboxes, enabledmailboxes.asp, inc_enabled_mailboxes.vbs and inc_enabled_mailboxes_databases.vbs have all been created to replace the 'enabledusers' equivalents.<br><br>
NEW TABLE: <u>EnabledMailboxes</u><br><br>
The SQL script to create EnabledMailboxes if you need it can be found in: \EVDashboard\sql\EVDashboard Database Creation\CreateTables.sql<br>
This new mechanism is designed to work with mailbox archiving databases which are spread across more than one SQL server/instance.<br>
Previous attempts of using linkedserver queries did not work.<br>
Even if you only have one SQL server this new mechanism will work for you as well.<br><br>
<li>Updated EVDashboard_Data_Collector.vbs to reflect the above changes.</li>
<li>Fixed Hourly_Archiving_Rates.asp database connection code including timeout. Thanks to Paul Bendall for this.</li>
<li>Added the same timeout fix to as many other files as possible vbs and asp. If I have missed any please post on the forums.</li>
<li>Setup global string variables for the timeout fix; strSQLConnect & strSQLExecute in global.asp. Also included the same in EVDashboard_Data_Collector.vbs.</li>
<li>Included \EVDashboard\vbscript\EnabledMailboxesReportEmail.vbs which runs a SQL query against the new EnabledMailboxes database table, retrieves all the data form the table, then emails the result to an email address of your choice. You will need to open this script and edit in your SMTP server, email address details etc.</li><br>
This allows you to schedule an export of the report and email it.
<li>Added IndexDiskAlerts.asp. Disks which have a percentage value less than 25% are now reported to the index.asp page.</li>
</ul>
</h4>
<br>
<h3>Known Bugs:</h3>
<br>
<ul>
<li><h4>00001 - Date conversation within EVDashboard_Data_Collector.vbs, temporary work-around in place.<br>
For some reason when timestamps are selected with adodb they change according to regional settings.<br>
Currently date format is for UK (DD.MM.YYYY). Code changes will be needed for US style dates (MM.DD.YYYY).</li></h4>
</ul>


<br>
<h2>Version 0.14 - 4th December 2009</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Added IgnoreArchiveTriggerBit.asp, with associated elements in config.asp and BuildMultiDimensionalArray.asp. Will soon be adding XML file option as well.</li>
<li>Moved Filtergrid code which was duplicated in multiple pages into /includes/styles.asp.</li>
<li>Fixed - Apostrophe fix added to inc_Enabled_Users.vbs, inc_Disabled_Users.vbs, inc_Over_Receive_Limit.vbs, inc_Over_Send_Limit.vbs, inc_Over_Warning_Limit.vbs. Thanks Metatron.</li>
</ul>
</h4>
<br>
<h3>Known Bugs:</h3>
<br>
<ul>
<li><h4>00001 - Date conversation within EVDashboard_Data_Collector.vbs, temporary work-around in place.<br>
For some reason when timestamps are selected with adodb they change according to regional settings.<br>
Currently date format is for UK (DD.MM.YYYY). Code changes will be needed for US style dates (MM.DD.YYYY).</li></h4>
</ul>

</ul>

<br>
<h2>Version 0.13 - 17th November 2009</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Fixed issue within user report pages. The default action for the enter key on these pages is the Refresh Data button.</li><br>
Added the filter button onto these pages, without this filtering is impossible. For the moment the css is not pretty, but it works.
<li>Updated css for buttons on user report pages.</li>
<li>Removed inc_OldestRecord.vbs - This is no longer needed since the introduction of DatabaseStats.asp</li>
<li>Added OrphanedArchives.asp</li>
<li>Added ExchangeLegacyDNs.asp</li>
<li>Added Service Alerts to index.asp</li>
<li>Updated MSMQMessageCount.asp - Now includes options for retrieving info from one or all servers at a time</li>
<li>Added OrganisationArchivingHistory.asp - Gives you totals of the number of items archived per month for each database</li>
<li>Dropped Exchange Transaction Logs from Data Collector vbscript and header.asp. This provided unreliable and is not core to the goals of EV Dashboard. Files left in place in case a fix is found.</li>
<li>Fixed CSS for calendar. Previously the background for the month letters was set the same colour blue as the .th headings in tables.</li>
<li>Seperated out index.asp functions into seperate files for easier reading.</li>
<li>Fixed footer.asp evdiscuss.net link. Thanks raredirt.</li>
</ul>
</h4>
<br>
<h3>Known Bugs:</h3>
<br>
<ul>
<li><h4>00001 - Date conversation within EVDashboard_Data_Collector.vbs, temporary work-around in place.<br>
For some reason when timestamps are selected with adodb they change according to regional settings.<br>
Currently date format is for UK (DD.MM.YYYY). Code changes will be needed for US style dates (MM.DD.YYYY).</li></h4>
</ul>

</ul>
<br>
<h2>Version 0.12 - 9th October 2009</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<h4>
<ul>
<li>Added this change log to track changes and bug fixes.</li>
<li>Fixed Minor enhancements to the EVDashboard_Data_Collector.vbs file.</li>
<li>Fixed connection bug in Mailboxstatus.asp - Connection type selection for Warning Status not present.</li>
<li>Added Number of users over receive limit to MailboxStatus.asp.</li>
<li>Fixed EnabledUsersCounts.asp - Page now retrieves data from EV Dashboard database.</li>
<li>Added ArchivesGrowth.asp (graph & list) & ArchivesGrowthExport.asp - Selects counts from enabledusers table by date.</li>
<li>Fixed EnabledUsers.asp - Date column data now selecting with CONVERT(nvarchar(30), timestamp_field, 120) AS 'timestamp_field'.<br>
Dates are now represented in ISO format and sorting works correctly. This does not fix bug 00001, but makes things better.</li>
<li>Fixed OverReceiveLimit.asp - The fix applied to datetime column(s).</li>
<li>Fixed OverSendLimit.asp - The fix applied to datetime column(s).</li>
<li>Fixed OverWarningLimit.asp - The fix applied to datetime column(s).</li>
<li>Added EV Dashboard database info to config.asp.</li>
<li>Fixed browser page title.</li>
<li>Added ExchangeDiskInfo.asp.</li>
<li>Removed the need for a .xml file for Fusion Chart in ArchivesGrowth.asp.</li>
<li>Moved some .asp files to ./includes.</li>
<li>Removed the need for a .xml file for Fusion Chart in ExchangeMailboxes.asp.</li>
<li>Fixed EVDiskUsageInfo.asp; updated column names and data. Sorting now functioning.</li>
<li>Fixed ExchangeDiskInfo.asp; same as above.</li>
<li>Removed the need for a .xml file for Fusion Chart in Index.asp.</li>
<li>Added Link to DocMessageClass website.</li>
<li>Added ./includes/FC_Colors.asp for Fusion Chart color selections.</li>
<li>Updated CSS in index.asp. Borders are now correctly show on the data tables.</li>
<li>Updated HourlyArchivingRates.asp to include table css, filtering and sorting.</li>
<li>Added MSMQPath to config.asp for each EV Server. Used in MSMQFolderSizes.asp.</li>
<li>Added TestDBConnectivity.asp; it is now possible to test your DB servers for ADODB / Remote access on port 1433 (or other) from within the site.</li>
<li>Removed the need for setup.asp.</li>
<li>Changed the icon set to Silk Icons.</li>
<li>Added Code for EV site name to be included in index.asp.</li>
<li>Fixed MSMQMessageCount.asp. Updated to use new table format.</li>
<li>Added UsefulInfo.asp.</li>
<li>Added Date Selector for viewing historic data from the EVDashboard database to:</li>
OverWarningLimit.asp, OverSendLimit.asp, OverReceiveLimit.asp, EnabledUsers.asp, DisabledUsers.asp.
<li>Added 'Message Queuing' service to EVServices.asp</li>
<li>Seperated funtions in EVDashboard_Data_Collector.vbs into seperate files for easier reading, use.</li>
<li>Added ExchangeTransactionLogs.asp</li>
<li>Updated HourlyArchivingRates.asp - Data is now recorded into the EVDashboard database. ASP page reads from EV Dashboard database, no longer runs against the live EV databases.</li>
<li>Updated EVDiskUsage.asp. Data is now recorded into the EV Dashboard database, retained the ability to view live data.</li>
<li>Updated ExchangeDiskUsage.asp. Data is now recorded into the EV Dashboard database, retained the ability to view live data.</li>
<li>Added EV Dashboard Database Statistics. Displaying Oldest, Newest and Count from each of the EV Dashboard database tables.</li>
</ul>
</h4>
<br>
<h3>Known Bugs:</h3>
<br>
<ul>
<li><h4>00001 - Date conversation within EVDashboard_Data_Collector.vbs, temporary work-around in place.<br>
For some reason when timestamps are selected with adodb they change according to regional settings.<br>
Currently date format is for UK (DD.MM.YYYY). Code changes will be needed for US style dates (MM.DD.YYYY).</li></h4>
</ul>
<br>
<h2>Version 0.11 - 2nd September 2009 - First public release.</h2>
<br>
<h3>Fixes / Additions:</h3>
<br>
<ul>
<h4><li>Added all known scripts and more from EV Toolkit HTA:</li>
<br><br>
<li>MSMQs - Message Counts</li><br>
<li>MSMQs - Folder Sizes</li><br>
<li>Ping Tests</li><br>
<li>EV Services</li><br>
<li>Disk Usage Information</li><br>
<li>Watch File Counts</li><br>
<li>Journal Archive Counts</li><br>
<li>Hourly Archiving Rates</li><br>
<br>
<li>User Index Status</li><br>
<li>Enabled Users</li><br>
<li>Disabled Users</li><br>
<li>Archive States</li><br>
<li>Users Over Warning Limit</li><br>
<li>Users Over Send Limit</li><br>
<li>Users Over Receive Limit</li><br>
<li>Mailbox Quota Status</li><br>
<li>Exchange States</li><br>
<br>
<li>Exchange Mailbox Counts - By Server - Graph</li><br>
<li>Exchange Mailbox Counts - By Mailstore</li><br>
<br><br>
</ul>
<ul>
<li>KVS Registry Keys</li>
</h4>
</ul>
<h3>Known Bugs:</h3>
<br>
<ul>
<li><h4>00001 - Date conversation within EVDashboard_Data_Collector.vbs, temporary work-around in place.<br>
For some reason when timestamps are selected with adodb they change according to regional settings.<br>
Currently date format is for UK (DD.MM.YYYY). Code changes will be needed for US style dates (MM.DD.YYYY).</li></h4>
</ul>
<br>
<!--#include file="footer.asp"-->