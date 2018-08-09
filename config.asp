<%
on error goto 0
'*******************************************************************************
'EDIT CONFIGURATION PARAMETER HERE
'------- Variables for Configuration
Dim strConfigurationFile, strConfigurationPath

strConfigurationPath = (Server.MapPath("./config"))
strConfigurationFile = strConfigurationPath & "/config.xml"

CheckXMLConfigVersion = "0.24"
'------- End variables for configuration
'*******************************************************************************

Const ForReading = 1, ForWriting = 2, ForAppending = 8

set fso = CreateObject("Scripting.FileSystemObject")

If instr(ReportFileStatus(strConfigurationFile)," doesn't exist:") then
  response.write ReportFileStatus(strConfigurationFile)
End If

Function ReportFileStatus(filespec)
   on error resume next
   Dim fso, msg
   Set fso = Server.CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      msg = filespec & " exists."
   Else
      msg = filespec & " doesn't exist: " & err.description
   End If
   ReportFileStatus = msg
End Function

'* Function to convert date into standard ISO format, regardless of the locale.
Function dbDate(dt)
	dbDate = year(dt) & "/" & right("0" & month(dt), 2) & "/" & right("0" & day(dt),2) & " " & formatdatetime(dt,4) 
End Function

on error resume next
	
Dim ConfigVersion, DirectoryServer, ConnectionString, SQLUserName, SQLPassword, SQLUseAltPort, SQLAltPortNumber
Dim EnabledUsersReportOption
Dim EVDSQLServer, EVDSQLDatabase, EVDConnectionString, EVDSQLUserName, EVDSQLPassword, EVDSQLUseAltPort, EVDSQLAltPortNumber
Dim StartTime, FinishTime, SiteName, EVVersion
Dim EVDSQLVersion, EVDSQLServicePack, EVDSQLEdition
Dim DNSAlias1, EVServer1, Datacenter1, EVDatabase1, DBServer1, Role1, MSMQPath1, TriggerPath1
Dim DNSAlias2, EVServer2, Datacenter2, EVDatabase2, DBServer2, Role2, MSMQPath2, TriggerPath2
Dim DNSAlias3, EVServer3, Datacenter3, EVDatabase3, DBServer3, Role3, MSMQPath3, TriggerPath3
Dim DNSAlias4, EVServer4, Datacenter4, EVDatabase4, DBServer4, Role4, MSMQPath4, TriggerPath4
Dim DNSAlias5, EVServer5, Datacenter5, EVDatabase5, DBServer5, Role5, MSMQPath5, TriggerPath5
Dim DNSAlias6, EVServer6, Datacenter6, EVDatabase6, DBServer6, Role6, MSMQPath6, TriggerPath6
Dim DNSAlias7, EVServer7, Datacenter7, EVDatabase7, DBServer7, Role7, MSMQPath7, TriggerPath7
Dim DNSAlias8, EVServer8, Datacenter8, EVDatabase8, DBServer8, Role8, MSMQPath8, TriggerPath8
Dim DNSAlias9, EVServer9, Datacenter9, EVDatabase9, DBServer9, Role9, MSMQPath9, TriggerPath9
Dim DNSAlias10, EVServer10, Datacenter10, EVDatabase10, DBServer10, Role10, MSMQPath10, TriggerPath10
Dim DNSAlias11, EVServer11, Datacenter11, EVDatabase11, DBServer11, Role11, MSMQPath11, TriggerPath11
Dim DNSAlias12, EVServer12, Datacenter12, EVDatabase12, DBServer12, Role12, MSMQPath12, TriggerPath12
Dim DNSAlias13, EVServer13, Datacenter13, EVDatabase13, DBServer13, Role13, MSMQPath13, TriggerPath13
Dim DNSAlias14, EVServer14, Datacenter14, EVDatabase14, DBServer14, Role14, MSMQPath14, TriggerPath14
Dim DNSAlias15, EVServer15, Datacenter15, EVDatabase15, DBServer15, Role15, MSMQPath15, TriggerPath15
Dim DNSAlias16, EVServer16, Datacenter16, EVDatabase16, DBServer16, Role16, MSMQPath16, TriggerPath16
Dim DNSAlias17, EVServer17, Datacenter17, EVDatabase17, DBServer17, Role17, MSMQPath17, TriggerPath17
Dim DNSAlias18, EVServer18, Datacenter18, EVDatabase18, DBServer18, Role18, MSMQPath18, TriggerPath18
Dim DNSAlias19, EVServer19, Datacenter19, EVDatabase19, DBServer19, Role19, MSMQPath19, TriggerPath19
Dim DNSAlias20, EVServer20, Datacenter20, EVDatabase20, DBServer20, Role20, MSMQPath20, TriggerPath20
Dim DomainName, ServiceAccount


Call LoadXML
'Dim bAdmin
Public Function LoadXML
	'Set options based upon contents of configuration XML file.

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")

	xmlDoc.async = False
	xmlDoc.load strConfigurationFile
	set strEVDashboard = xmlDoc.documentelement.selectSingleNode("//EVDashboard")
	ConfigVersion = strEVDashboard.getAttribute("ConfigVersion")

	Set strConfig = xmlDoc.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")
	DirectoryServer = strConfig.getAttribute("DirectoryServer")
	
	EVVersion = strConfig.getAttribute("EVVersion")
	
	DomainName = strConfig.getAttribute("DomainName")
	ServiceAccount = strConfig.getAttribute("ServiceAccount")
	
	ConnectionString = strConfig.getAttribute("ConnectionString")
	SQLUserName = strConfig.getAttribute("SQLUserName")
	SQLPassword = strConfig.getAttribute("SQLPassword")

	SQLUseAltPort = strConfig.getAttribute("SQLUseAltPort")
	SQLAltPortNumber = strConfig.getAttribute("SQLAltPortNumber")
	
	EnabledUsersReportOption = strConfig.getAttribute("EnabledUsersReportOption")
	
	EVDSQLServer = strConfig.getAttribute("EVDSQLServer")
	EVDSQLDatabase = strConfig.getAttribute("EVDSQLDatabase")
	
	EVDConnectionString = strConfig.getAttribute("EVDConnectionString")
	EVDSQLUserName = strConfig.getAttribute("EVDSQLUserName")
	EVDSQLPassword = strConfig.getAttribute("EVDSQLPassword")

	EVDSQLUseAltPort = strConfig.getAttribute("EVDSQLUseAltPort")
	EVDSQLAltPortNumber = strConfig.getAttribute("EVDSQLAltPortNumber")
	
	StartTime = strConfig.getAttribute("StartTime")
	FinishTime = strConfig.getAttribute("FinishTime")
	SiteName = strConfig.getAttribute("SiteName")
	
	EVDSQLVersion = strConfig.getAttribute("EVDSQLVersion")
	EVDSQLServicePack = strConfig.getAttribute("EVDSQLServicePack")
	EVDSQLEdition = strConfig.getAttribute("EVDSQLEdition")
	
	DNSAlias1 = strConfig.getAttribute("DNSAlias1")
	EVServer1 = strConfig.getAttribute("EVServer1")
	Datacenter1 = strConfig.getAttribute("Datacenter1")
	EVDatabase1 = strConfig.getAttribute("EVDatabase1")
	DBServer1 = strConfig.getAttribute("DBServer1")
	Role1 = strConfig.getAttribute("Role1")
	MSMQPath1 = strConfig.getAttribute("MSMQPath1")
	TriggerPath1 = strConfig.getAttribute("TriggerPath1")
	
	DNSAlias2 = strConfig.getAttribute("DNSAlias2")
	EVServer2 = strConfig.getAttribute("EVServer2")
	Datacenter2 = strConfig.getAttribute("Datacenter2")
	EVDatabase2 = strConfig.getAttribute("EVDatabase2")
	DBServer2 = strConfig.getAttribute("DBServer2")
	Role2 = strConfig.getAttribute("Role2")
	MSMQPath2 = strConfig.getAttribute("MSMQPath2")
	TriggerPath2 = strConfig.getAttribute("TriggerPath2")
	
	DNSAlias3 = strConfig.getAttribute("DNSAlias3")
	EVServer3 = strConfig.getAttribute("EVServer3")
	Datacenter3 = strConfig.getAttribute("Datacenter3")
	EVDatabase3 = strConfig.getAttribute("EVDatabase3")
	DBServer3 = strConfig.getAttribute("DBServer3")
	Role3 = strConfig.getAttribute("Role3")
	MSMQPath3 = strConfig.getAttribute("MSMQPath3")
	TriggerPath3 = strConfig.getAttribute("TriggerPath3")
	
	DNSAlias4 = strConfig.getAttribute("DNSAlias4")
	EVServer4 = strConfig.getAttribute("EVServer4")
	Datacenter4 = strConfig.getAttribute("Datacenter4")
	EVDatabase4 = strConfig.getAttribute("EVDatabase4")
	DBServer4 = strConfig.getAttribute("DBServer4")
	Role4 = strConfig.getAttribute("Role4")
	MSMQPath4 = strConfig.getAttribute("MSMQPath4")
	TriggerPath4 = strConfig.getAttribute("TriggerPath4")
	
	DNSAlias5 = strConfig.getAttribute("DNSAlias5")
	EVServer5 = strConfig.getAttribute("EVServer5")
	Datacenter5 = strConfig.getAttribute("Datacenter5")
	EVDatabase5 = strConfig.getAttribute("EVDatabase5")
	DBServer5 = strConfig.getAttribute("DBServer5")
	Role5 = strConfig.getAttribute("Role5")
	MSMQPath5 = strConfig.getAttribute("MSMQPath5")
	TriggerPath5 = strConfig.getAttribute("TriggerPath5")
	
	DNSAlias6 = strConfig.getAttribute("DNSAlias6")
	EVServer6 = strConfig.getAttribute("EVServer6")
	Datacenter6 = strConfig.getAttribute("Datacenter6")
	EVDatabase6 = strConfig.getAttribute("EVDatabase6")
	DBServer6 = strConfig.getAttribute("DBServer6")
	Role6 = strConfig.getAttribute("Role6")
	MSMQPath6 = strConfig.getAttribute("MSMQPath6")
	TriggerPath6 = strConfig.getAttribute("TriggerPath6")
	
	DNSAlias7 = strConfig.getAttribute("DNSAlias7")
	EVServer7 = strConfig.getAttribute("EVServer7")
	Datacenter7 = strConfig.getAttribute("Datacenter7")
	EVDatabase7 = strConfig.getAttribute("EVDatabase7")
	DBServer7 = strConfig.getAttribute("DBServer7")
	Role7 = strConfig.getAttribute("Role7")
	MSMQPath7 = strConfig.getAttribute("MSMQPath7")
	TriggerPath7 = strConfig.getAttribute("TriggerPath7")
	
	DNSAlias8 = strConfig.getAttribute("DNSAlias8")
	EVServer8 = strConfig.getAttribute("EVServer8")
	Datacenter8 = strConfig.getAttribute("Datacenter8")
	EVDatabase8 = strConfig.getAttribute("EVDatabase8")
	DBServer8 = strConfig.getAttribute("DBServer8")
	Role8 = strConfig.getAttribute("Role8")
	MSMQPath8 = strConfig.getAttribute("MSMQPath8")
	TriggerPath8 = strConfig.getAttribute("TriggerPath8")
	
	DNSAlias9 = strConfig.getAttribute("DNSAlias9")
	EVServer9 = strConfig.getAttribute("EVServer9")
	Datacenter9 = strConfig.getAttribute("Datacenter9")
	EVDatabase9 = strConfig.getAttribute("EVDatabase9")
	DBServer9 = strConfig.getAttribute("DBServer9")
	Role9 = strConfig.getAttribute("Role9")
	MSMQPath9 = strConfig.getAttribute("MSMQPath9")
	TriggerPath9 = strConfig.getAttribute("TriggerPath9")
	
	DNSAlias10 = strConfig.getAttribute("DNSAlias10")
	EVServer10 = strConfig.getAttribute("EVServer10")
	Datacenter10 = strConfig.getAttribute("Datacenter10")
	EVDatabase10 = strConfig.getAttribute("EVDatabase10")
	DBServer10 = strConfig.getAttribute("DBServer10")
	Role10 = strConfig.getAttribute("Role10")
	MSMQPath10 = strConfig.getAttribute("MSMQPath10")
	TriggerPath10 = strConfig.getAttribute("TriggerPath10")
	
	DNSAlias11 = strConfig.getAttribute("DNSAlias11")
	EVServer11 = strConfig.getAttribute("EVServer11")
	Datacenter11 = strConfig.getAttribute("Datacenter11")
	EVDatabase11 = strConfig.getAttribute("EVDatabase11")
	DBServer11 = strConfig.getAttribute("DBServer11")
	Role11 = strConfig.getAttribute("Role11")
	MSMQPath11 = strConfig.getAttribute("MSMQPath11")
	TriggerPath11 = strConfig.getAttribute("TriggerPath11")
	
	DNSAlias12 = strConfig.getAttribute("DNSAlias12")
	EVServer12 = strConfig.getAttribute("EVServer12")
	Datacenter12 = strConfig.getAttribute("Datacenter12")
	EVDatabase12 = strConfig.getAttribute("EVDatabase12")
	DBServer12 = strConfig.getAttribute("DBServer12")
	Role12 = strConfig.getAttribute("Role12")
	MSMQPath12 = strConfig.getAttribute("MSMQPath12")
	TriggerPath12 = strConfig.getAttribute("TriggerPath12")
	
	DNSAlias13 = strConfig.getAttribute("DNSAlias13")
	EVServer13 = strConfig.getAttribute("EVServer13")
	Datacenter13 = strConfig.getAttribute("Datacenter13")
	EVDatabase13 = strConfig.getAttribute("EVDatabase13")
	DBServer13 = strConfig.getAttribute("DBServer13")
	Role13 = strConfig.getAttribute("Role13")
	MSMQPath13 = strConfig.getAttribute("MSMQPath13")
	TriggerPath13 = strConfig.getAttribute("TriggerPath13")
	
	DNSAlias14 = strConfig.getAttribute("DNSAlias14")
	EVServer14 = strConfig.getAttribute("EVServer14")
	Datacenter14 = strConfig.getAttribute("Datacenter14")
	EVDatabase14 = strConfig.getAttribute("EVDatabase14")
	DBServer14 = strConfig.getAttribute("DBServer14")
	Role14 = strConfig.getAttribute("Role14")
	MSMQPath14 = strConfig.getAttribute("MSMQPath14")
	TriggerPath14 = strConfig.getAttribute("TriggerPath14")
	
	DNSAlias15 = strConfig.getAttribute("DNSAlias15")
	EVServer15 = strConfig.getAttribute("EVServer15")
	Datacenter15 = strConfig.getAttribute("Datacenter15")
	EVDatabase15 = strConfig.getAttribute("EVDatabase15")
	DBServer15 = strConfig.getAttribute("DBServer15")
	Role15 = strConfig.getAttribute("Role15")
	MSMQPath15 = strConfig.getAttribute("MSMQPath15")
	TriggerPath15 = strConfig.getAttribute("TriggerPath15")

	DNSAlias16 = strConfig.getAttribute("DNSAlias16")
	EVServer16 = strConfig.getAttribute("EVServer16")
	Datacenter16 = strConfig.getAttribute("Datacenter16")
	EVDatabase16 = strConfig.getAttribute("EVDatabase16")
	DBServer16 = strConfig.getAttribute("DBServer16")
	Role16 = strConfig.getAttribute("Role16")
	MSMQPath16 = strConfig.getAttribute("MSMQPath16")
	TriggerPath16 = strConfig.getAttribute("TriggerPath16")

	DNSAlias17 = strConfig.getAttribute("DNSAlias17")
	EVServer17 = strConfig.getAttribute("EVServer17")
	Datacenter17 = strConfig.getAttribute("Datacenter17")
	EVDatabase17 = strConfig.getAttribute("EVDatabase17")
	DBServer17 = strConfig.getAttribute("DBServer17")
	Role17 = strConfig.getAttribute("Role17")
	MSMQPath17 = strConfig.getAttribute("MSMQPath17")
	TriggerPath17 = strConfig.getAttribute("TriggerPath17")

	DNSAlias18 = strConfig.getAttribute("DNSAlias18")
	EVServer18 = strConfig.getAttribute("EVServer18")
	Datacenter18 = strConfig.getAttribute("Datacenter18")
	EVDatabase18 = strConfig.getAttribute("EVDatabase18")
	DBServer18 = strConfig.getAttribute("DBServer18")
	Role18 = strConfig.getAttribute("Role18")
	MSMQPath18 = strConfig.getAttribute("MSMQPath18")
	TriggerPath18 = strConfig.getAttribute("TriggerPath18")

	DNSAlias19 = strConfig.getAttribute("DNSAlias19")
	EVServer19 = strConfig.getAttribute("EVServer19")
	Datacenter19 = strConfig.getAttribute("Datacenter19")
	EVDatabase19 = strConfig.getAttribute("EVDatabase19")
	DBServer19 = strConfig.getAttribute("DBServer19")
	Role19 = strConfig.getAttribute("Role19")
	MSMQPath19 = strConfig.getAttribute("MSMQPath19")
	TriggerPath19 = strConfig.getAttribute("TriggerPath19")

	DNSAlias20 = strConfig.getAttribute("DNSAlias20")
	EVServer20 = strConfig.getAttribute("EVServer20")
	Datacenter20 = strConfig.getAttribute("Datacenter20")
	EVDatabase20 = strConfig.getAttribute("EVDatabase20")
	DBServer20 = strConfig.getAttribute("DBServer20")
	Role20 = strConfig.getAttribute("Role20")
	MSMQPath20 = strConfig.getAttribute("MSMQPath20")
	TriggerPath20 = strConfig.getAttribute("TriggerPath20")

	Set xmldoc = Nothing
		
End Function

%>