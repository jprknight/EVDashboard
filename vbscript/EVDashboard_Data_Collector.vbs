'********************************************************************************
'*
'*	jprknight - EV Dashboard Data Collector
'*
'*	This script reads your pre-setup config.xml, then runs SQL scripts against tables to get lists of
'*	* Enabled Users
'*	* Disabled Users
'*	* Over Warning Limit
'*	* Over Send Limit
'*	* Over Receive Limit
'*
'*	The front end queries one table, in one database, on one database server instead of multiples.
'*	In larger environments, this is essential for performance reasons.
'*
'*	24.08.2009	1.0		* Intial draft
'*	25.08.2009	1.1		* Base code completed for Enabled Users.
'*						* Timestamp select issue with adodb within EnabledUsers query.
'*							Work around in place, until this is properly fixed. Documented
'*							further within the query.
'*	26.08.2009	1.1a	* Added DisabledUsers, OverWarningLimit, OverSendLimit, OverReceiveLimit.
'*	28.08.2009	1.2		* General clean-up and code fixes.
'*	02.09.2009	1.2.1	* Added Log file writing functionality.
'*						* Fixed issue within EnabledUsers; removed explicit if statement referencing
'*							EV databases on test/development environment.
'*						* Fixed issue within EnabledUsers; missing [ ] signifying linked SQL Server.
'*						* Replaced all On Error Goto 0 with On Error Resume Next so error handling is 
'*							used correctly.
'*	03.09.2009	1.2.2	* Minor code checking / fixing.
'*	11.09.2009	1.2.3	* Added section to retrieve oldest record timestamp from enabledusers table.
'*	15.09.2009	1.2.4	* Added EV Site Name.
'*	05.10.2009	1.2.5	* Seperated out functions into seperate files, making it easier to work with the script.
'*	05.01.2010	1.2.6	* Added UserConfiguration code for either enabled users or enabled mailboxes.
'*	01.02.2010	1.2.7	* Updated User Configuration variable based on time of day. This script can now
'*							be run within the same day without editing this file.
'*	18.02.2010	1.2.8	* Added further output information written to log and screen. Added if then else
'*							condition so start and finish time is only written if the script was started
'*							before 8:30am.
'*	11.03.2010	1.2.9	* Extended strSQLExecute value to 1800 seconds or 30 minutes.
'*						* ISO date function moved in from inc_ scripts. No need to repeat the same function
'*							within several different script files.
'*	08.04.2010	1.2.10 	* Added code for writting out the status of the StartTime and FinishTime to screen and to
'*						* log file.
'*	23.08.2010	1.2.11 	* Added LoggingLevel option to control the level of information which is written to the
'*							EVDashboard_Data_Collector.log file.
'*      20.02.2011	1.2.12	* Added parameter "INIT" to allow for collecting all data after 08:30
'*
Version = "v1.2.12"
'********************************************************************************
'* Logging Level
'*
'* 0 = No logging. No entries are written to a log file.
'* 1 = Minimal logging. Tends to only include notifications if jobs starting and stopping.
'* 2 = Normal logging. Also includes information on the objects within the jobs.
'* 3 = verbose logging. Also includes SQL queries running on objects within jobs.

'* Regardless of the selection errors are still written to screen and the log file.
'* Logging levels 2 and above are not written to screen by default. These lines have been commented out.
LoggingLevel = 1
'*
'********************************************************************************
ON ERROR RESUME NEXT

'* Start Time printed on footer of the front end.
StartTime = DateTime()

Sthr = Hour(Now)
if Sthr < 10 Then
	Sthr = "0" & Sthr
end if
Stmin = Minute(Now)
if Stmin < 10 Then
	Stmin = "0" & Stmin
end if
Stsec = Second(Now)
if Stsec < 10 Then
	Stsec = "0" & Stsec
end if

strStartTimeForWrittingXML = Sthr & Stmin & Stsec

Dim oShell
Set oShell = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Const adCmdText = 1

'* Give the strSQLConnect a timeout value of 5 minutes.
strSQLConnect = 300
'* Give the strSQLExecute a timeout value of 30 minutes.
strSQLExecute = 300

'* Force Script to run using CSCRIPT.
forceUseCScript

Sub forceUseCScript()
	Dim RequiredParam
	RequiredParam = ""
	If Not WScript.FullName = WScript.Path & "\cscript.exe" Then
		'oShell.Popup "Launched using wscript. Relaunching...",3,"WSCRIPT"
		If WScript.Arguments.Count > 0 Then
			If UCase(WScript.Arguments(0)) = "INIT" Then
				RequiredParam = " INIT"
			End If
		End If
		oShell.Run "cmd.exe /k " & WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & _
			WScript.scriptFullName & Chr(34) & RequiredParam,1,False
		WScript.Quit 0
	End If

'	If Not WScript.FullName = WScript.Path & "\cscript.exe" Then
'		'oShell.Popup "Launched using wscript. Relaunching...",3,"WSCRIPT"
'		oShell.Run "cmd.exe /k " & WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & _
'			WScript.scriptFullName & Chr(34),1,False
'		WScript.Quit 0
'	End If
End Sub

'********************************************************************************

Dim strConfigurationFile

strConfigurationFile = "../config/config.xml"

'********************************************************************************
'* Log File Configuration.
LogFile = "EVDashboard_Data_Collector.log"

If fso.FileExists(LogFile) Then
	Set OutputLogFile = fso.OpenTextFile(LogFile, ForAppending)
else
	Set OutputLogFile = fso.CreateTextFile(LogFile, ForWriting)
end if

'********************************************************************************
wscript.echo "******************************************************"
wscript.echo "*"
wscript.echo "*  EVDashboard Data Collector - " & Version
wscript.echo "*"
wscript.echo "******************************************************"
wscript.echo ""
OutputLogFile.writeline ""
OutputLogFile.writeline "******************************************************"
OutputLogFile.writeline "*"
OutputLogFile.writeline "*  EVDashboard Data Collector - " & Version
OutputLogFile.writeline "*"
OutputLogFile.writeline "******************************************************"
OutputLogFile.writeline ""
'********************************************************************************

If instr(ReportFileStatus(strConfigurationFile)," doesn't exist:") then
  wscript.echo DateTime() & " - " & ReportFileStatus(strConfigurationFile)
  OutputLogFile.writeline DateTime() & " - " & ReportFileStatus(strConfigurationFile)
End If

'********************************************************************************
Function ReportFileStatus(filespec)
   on error resume next
   Dim fso, msg
   Set fso = Server.CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      msg = filespec & " exists."
	  wscript.echo DateTime() & ": Successfully opened config.xml."
	  OutputLogFile.writeline DateTime() & ": Successfully opened config.xml."
   Else
      msg = filespec & " doesn't exist: " & err.description
   End If
   ReportFileStatus = msg
End Function
'********************************************************************************
Function DateTime
	Styr = Year(Now)
	Stmo = Month(Now)
	if Stmo < 10 Then
		Stmo = "0" & Stmo
	end if
	Stdt = Day(Now)
	if Stdt < 10 Then
		Stdt = "0" & Stdt
	end if
	Sthr = Hour(Now)
	if Sthr < 10 Then
		Sthr = "0" & Sthr
	end if
	Stmin = Minute(Now)
	if Stmin < 10 Then
		Stmin = "0" & Stmin
	end if
	Stsec = Second(Now)
	if Stsec < 10 Then
		Stsec = "0" & Stsec
	end if
	DateTime = styr & "." & stmo & "." & stdt & " " & sthr & ":" & stmin & ":" & stsec
End Function
'********************************************************************************
'* Function to convert date into standard ISO format, regardless of the locale.
Function dbDate(dt)
	'* Included if statement. If the user has no last archived value, the function will not be able to process the date, as it is null and will error out.
	If Len(dt) > 1 then
		dbDate = year(dt) & "/" & right("0" & month(dt), 2) & "/" & right("0" & day(dt),2) & " " & formatdatetime(dt,4)
	End if
End Function
'********************************************************************************
'* Open script files and execute them.

Dim UserConfiguration
Dim TestMode: TestMode = 0

'* Assuming the Data Collector script will first run before 08:30 in the morning,
'* and if the script is run again within the same day only the quicker scripts need
'* to be run to refresh service alerts data etc.

If TestMode = 0 then
	If sthr & stmin & stsec < "083000" then
		UserConfiguration = 1
		wscript.echo DateTime() & ": Running in pre-online day mode."
		OutputLogFile.writeline DateTime() & ": Running in pre-online day mode."
	Elseif sthr & stmin & stsec >= "083000" then
		UserConfiguration = 3
		wscript.echo DateTime() & ": Running in online day mode."
		OutputLogFile.writeline DateTime() & ": Running in online day mode."
	End if
Else
	UserConfiguration = 2
	wscript.echo DateTime() & ": Running in TEST mode."
	OutputLogFile.writeline DateTime() & ": Running in TEST mode."
End if

If WScript.Arguments.Count > 0 Then
	If UCase(WScript.Arguments(0)) = "INIT" Then
		UserConfiguration=1
		WScript.Echo DateTime() & ": Running in explicit INIT mode."
		OutputLogFile.writeline DateTime() & ": Running in explicit INIT mode."
	End If
End If

'* 1 - default - This is the normal setting required.
'* 2 - testing - ignore this setting. It is not for normal use.
'* 3 - daily refresh - All scripts apart from the user reports which take a long time to run.
If UserConfiguration = 1 then
	Files = Array("inc_Build_MultidimentionalArray.vbs",_
		"inc_UsageReportByArchive.vbs",_
		"inc_Disabled_Users.vbs",_
		"inc_Over_Send_Limit.vbs",_
		"inc_Over_Warning_Limit.vbs",_
		"inc_Over_Receive_Limit.vbs",_
		"inc_SQL_Info.vbs",_
		"inc_EV_Site_Name.vbs",_
		"inc_Hourly_Archiving_Rates.vbs",_
		"inc_Organisation_Archiving_History.vbs",_
		"inc_Service_Alerts.vbs",_
		"inc_EV_Disk_Info.vbs",_
		"inc_Exchange_Disk_Info.vbs",_
		"inc_Exchange_Mailboxes.vbs",_
		"inc_Domain_Name.vbs",_
		"inc_EV_Server_EventLogs.vbs",_
		"inc_SEV_EventLogs.vbs",_
		"inc_SIS_Report.vbs",_
		"inc_Enabled_Mailboxes.vbs",_
		"inc_Enabled_Mailboxes_Databases.vbs",_
		"inc_MSMQMessageCountAlerts.vbs")
Elseif UserConfiguration = 2 then
	'* TESTING
	Files = Array("inc_Build_MultidimentionalArray.vbs","inc_SEV_EventLogs.vbs")
	'* TESTING
Elseif UserConfiguration = 3 then
	'* DAILY REFRESH - Run all scripts apart from the user reports which take a long time to run.
	Files = Array("inc_Build_MultidimentionalArray.vbs",_
	"inc_Service_Alerts.vbs",_
	"inc_Hourly_Archiving_Rates.vbs",_
	"inc_Enabled_Mailboxes.vbs",_
	"inc_MSMQMessageCountAlerts.vbs")',_
'	"inc_EV_Server_EventLogs.vbs")
	'* DAILY REFRESH
End if

'* Run through each of the files set in the 'files' array from the above if statement.
For each file in Files
	Dim f: set f = fso.OpenTextFile(file,ForReading)
	Dim s: s = f.ReadAll()
	ExecuteGlobal s
Next
'********************************************************************************
wscript.echo DateTime() & ": Finished."
OutputLogFile.writeline DateTime() & ": Finished."
'* Finish time is writen the footer of the front end.
FinishTime = DateTime()

wscript.sleep 250

'* Only write start and finish time to config.xml if it is pre-8:30am
If strStartTimeForWrittingXML < "083000" then
	'* Clear any previous errors so only XML write errors come up.
	Err.clear

	set xmlDoc = createobject("Microsoft.XMLDOM")
	xmlDoc.async = false
	xmlDoc.load strConfigurationFile

	set ServerConfig = xmlDoc.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")

	'* Attributes written to footer of the front end.
	ServerConfig.setAttribute("StartTime") = StartTime
	ServerConfig.setAttribute("FinishTime") = FinishTime

	xmlDoc.save strConfigurationFile

	If err.number <> 0 then 
		wscript.echo "Error occured in XMLSave - Error number: " & err.number & _
			vbcrlf & "Description " & DateTime() & ": " & Err.Description
		OutputLogFile.writeline "Error occured in XMLSave - Error number: " & _
			err.number & vbcrlf & "Description " & DateTime() & ": " & Err.Description
	Else
		wscript.echo DateTime() & ": Written StartTime as " & StartTime
		wscript.echo DateTime() & ": Written FinishTime as " & FinishTime
		OutputLogFile.writeline DateTime() & ": Written StartTime as " & StartTime
		OutputLogFile.writeline DateTime() & ": Written FinishTime as " & FinishTime
	End if

	set xmlDoc = nothing
End if

OutputLogFile.writeline ""
OutputLogFile.close

wscript.quit
'********************************************************************************
'* SCRIPT END
'********************************************************************************