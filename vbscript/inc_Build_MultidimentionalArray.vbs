'* Start build multidimential array.
Dim ConfigVersion, DirectoryServer, ConnectionString, SQLUserName, SQLPassword, SQLUseAltPort, SQLAltPortNumber
Dim EnabledUsersReportOption
Dim EVDSQLServer, EVDSQLDatabase, EVDConnectionString, EVDSQLUserName, EVDSQLPassword, SQLEVDUseAltPort, SQLEVDAltPortNumber
Dim StartTime, FinishTime, SiteName
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

'Set options based upon contents of configuration XML file.

wscript.echo DateTime() & ": Loading data from config.xml."
OutputLogFile.writeline DateTime() & ": Loading data from config.xml."

Set xmlDoc = CreateObject("Microsoft.XMLDOM")

xmlDoc.async = False
xmlDoc.load strConfigurationFile
set strEVDashboard = xmlDoc.documentelement.selectSingleNode("//EVDashboard")
ConfigVersion = strEVDashboard.getAttribute("ConfigVersion")

Set strConfig = xmlDoc.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")
DirectoryServer = strConfig.getAttribute("DirectoryServer")

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

SQLEVDUseAltPort = strConfig.getAttribute("SQLEVDUseAltPort")
SQLEVDAltPortNumber = strConfig.getAttribute("SQLEVDAltPortNumber")

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
'* END LOAD DATA FROM XML FILE.
'********************************************************************************
'	Build dynamic multidimential array of EV Servers, DB Servers, DB Databases and Roles.

'	EVServersList(0,i)	DNS Alias
'	EVServersList(1,i)	EV Server
'	EVServersList(2,i)	Data Center
'	EVServersList(3,i)	DB Server
'	EVServersList(4,i)	EV Database
'	EVServersList(5,i)	EV Server Role
wscript.echo DateTime() & ": Creating multi-dimentional array."
OutputLogFile.writeline DateTime() & ": Creating multi-dimentional array."

If EVServer1 <> "" then
	Redim EVServersList(7,0)
	EVServersList(0,0) = DNSAlias1
	EVServersList(1,0) = EVServer1
	EVServersList(2,0) = Datacenter1
	EVServersList(3,0) = DBServer1
	EVServersList(4,0) = EVDatabase1
	EVServersList(5,0) = Role1
	EVServersList(6,0) = MSMQPath1
	EVServersList(7,0) = TriggerPath1
	
End If

If EVServer2 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,1) = DNSAlias2
	EVServersList(1,1) = EVServer2
	EVServersList(2,1) = Datacenter2
	EVServersList(3,1) = DBServer2
	EVServersList(4,1) = EVDatabase2
	EVServersList(5,1) = Role2
	EVServersList(6,1) = MSMQPath2
	EVServersList(7,1) = TriggerPath2
End If

If EVServer3 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,2) = DNSAlias3
	EVServersList(1,2) = EVServer3
	EVServersList(2,2) = Datacenter3
	EVServersList(3,2) = DBServer3
	EVServersList(4,2) = EVDatabase3
	EVServersList(5,2) = Role3
	EVServersList(6,2) = MSMQPath3
	EVServersList(7,2) = TriggerPath3
End If

If EVServer4 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,3) = DNSAlias4
	EVServersList(1,3) = EVServer4
	EVServersList(2,3) = Datacenter4
	EVServersList(3,3) = DBServer4
	EVServersList(4,3) = EVDatabase4
	EVServersList(5,3) = Role4
	EVServersList(6,3) = MSMQPath4
	EVServersList(7,3) = TriggerPath4
End If
	
If EVServer5 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,4) = DNSAlias5
	EVServersList(1,4) = EVServer5
	EVServersList(2,4) = Datacenter5
	EVServersList(3,4) = DBServer5
	EVServersList(4,4) = EVDatabase5
	EVServersList(5,4) = Role5
	EVServersList(6,4) = MSMQPath5
	EVServersList(7,4) = TriggerPath5
End If
	
If EVServer6 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,5) = DNSAlias6
	EVServersList(1,5) = EVServer6
	EVServersList(2,5) = Datacenter6
	EVServersList(3,5) = DBServer6
	EVServersList(4,5) = EVDatabase6
	EVServersList(5,5) = Role6
	EVServersList(6,5) = MSMQPath6
	EVServersList(7,5) = TriggerPath6
End If

If EVServer7 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,6) = DNSAlias7
	EVServersList(1,6) = EVServer7
	EVServersList(2,6) = Datacenter7
	EVServersList(3,6) = DBServer7
	EVServersList(4,6) = EVDatabase7
	EVServersList(5,6) = Role7
	EVServersList(6,6) = MSMQPath7
	EVServersList(7,6) = TriggerPath7
End If

If EVServer8 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,7) = DNSAlias8
	EVServersList(1,7) = EVServer8
	EVServersList(2,7) = Datacenter8
	EVServersList(3,7) = DBServer8
	EVServersList(4,7) = EVDatabase8
	EVServersList(5,7) = Role8
	EVServersList(6,7) = MSMQPath8
	EVServersList(7,7) = TriggerPath8
End If

If EVServer9 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,8) = DNSAlias9
	EVServersList(1,8) = EVServer9
	EVServersList(2,8) = Datacenter9
	EVServersList(3,8) = DBServer9
	EVServersList(4,8) = EVDatabase9
	EVServersList(5,8) = Role9
	EVServersList(6,8) = MSMQPath9
	EVServersList(7,8) = TriggerPath9
End If

If EVServer10 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,9) = DNSAlias10
	EVServersList(1,9) = EVServer10
	EVServersList(2,9) = Datacenter10
	EVServersList(3,9) = DBServer10
	EVServersList(4,9) = EVDatabase10
	EVServersList(5,9) = Role10
	EVServersList(6,9) = MSMQPath10
	EVServersList(7,9) = TriggerPath10
End If

If EVServer11 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,10) = DNSAlias11
	EVServersList(1,10) = EVServer11
	EVServersList(2,10) = Datacenter11
	EVServersList(3,10) = DBServer11
	EVServersList(4,10) = EVDatabase11
	EVServersList(5,10) = Role11
	EVServersList(6,10) = MSMQPath11
	EVServersList(7,10) = TriggerPath11
End If

If EVServer12 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,11) = DNSAlias12
	EVServersList(1,11) = EVServer12
	EVServersList(2,11) = Datacenter12
	EVServersList(3,11) = DBServer12
	EVServersList(4,11) = EVDatabase12
	EVServersList(5,11) = Role12
	EVServersList(6,11) = MSMQPath12
	EVServersList(7,11) = TriggerPath12
End If

If EVServer13 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,12) = DNSAlias13
	EVServersList(1,12) = EVServer13
	EVServersList(2,12) = Datacenter13
	EVServersList(3,12) = DBServer13
	EVServersList(4,12) = EVDatabase13
	EVServersList(5,12) = Role13
	EVServersList(6,12) = MSMQPath13
	EVServersList(7,12) = TriggerPath13
End If

If EVServer14 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,13) = DNSAlias14
	EVServersList(1,13) = EVServer14
	EVServersList(2,13) = Datacenter14
	EVServersList(3,13) = DBServer14
	EVServersList(4,13) = EVDatabase14
	EVServersList(5,13) = Role14
	EVServersList(6,13) = MSMQPath14
	EVServersList(7,13) = TriggerPath14
End If

If EVServer15 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,14) = DNSAlias15
	EVServersList(1,14) = EVServer15
	EVServersList(2,14) = Datacenter15
	EVServersList(3,14) = DBServer15
	EVServersList(4,14) = EVDatabase15
	EVServersList(5,14) = Role15
	EVServersList(6,14) = MSMQPath15
	EVServersList(7,14) = TriggerPath15
End If

If EVServer16 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,15) = DNSAlias16
	EVServersList(1,15) = EVServer16
	EVServersList(2,15) = Datacenter16
	EVServersList(3,15) = DBServer16
	EVServersList(4,15) = EVDatabase16
	EVServersList(5,15) = Role16
	EVServersList(6,15) = MSMQPath16
	EVServersList(7,15) = TriggerPath16
End If

If EVServer17 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,16) = DNSAlias17
	EVServersList(1,16) = EVServer17
	EVServersList(2,16) = Datacenter17
	EVServersList(3,16) = DBServer17
	EVServersList(4,16) = EVDatabase17
	EVServersList(5,16) = Role17
	EVServersList(6,16) = MSMQPath17
	EVServersList(7,16) = TriggerPath17
End If

If EVServer18 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,17) = DNSAlias18
	EVServersList(1,17) = EVServer18
	EVServersList(2,17) = Datacenter18
	EVServersList(3,17) = DBServer18
	EVServersList(4,17) = EVDatabase18
	EVServersList(5,17) = Role18
	EVServersList(6,17) = MSMQPath18
	EVServersList(7,17) = TriggerPath18
End If

If EVServer19 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,18) = DNSAlias19
	EVServersList(1,18) = EVServer19
	EVServersList(2,18) = Datacenter19
	EVServersList(3,18) = DBServer19
	EVServersList(4,18) = EVDatabase19
	EVServersList(5,18) = Role19
	EVServersList(6,18) = MSMQPath19
	EVServersList(7,18) = TriggerPath19
End If

If EVServer20 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,19) = DNSAlias20
	EVServersList(1,19) = EVServer20
	EVServersList(2,19) = Datacenter20
	EVServersList(3,19) = DBServer20
	EVServersList(4,19) = EVDatabase20
	EVServersList(5,19) = Role20
	EVServersList(6,19) = MSMQPath20
	EVServersList(7,19) = TriggerPath20
End If
'wscript.echo ubound(EVServersList,2) 'Output number of rows in the MD Array. 

If EVServer15 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,14) = DNSAlias15
	EVServersList(1,14) = EVServer15
	EVServersList(2,14) = Datacenter15
	EVServersList(3,14) = DBServer15
	EVServersList(4,14) = EVDatabase15
	EVServersList(5,14) = Role15
End If

If EVServer16 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,15) = DNSAlias16
	EVServersList(1,15) = EVServer16
	EVServersList(2,15) = Datacenter16
	EVServersList(3,15) = DBServer16
	EVServersList(4,15) = EVDatabase16
	EVServersList(5,15) = Role16
End If

If EVServer17 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,16) = DNSAlias17
	EVServersList(1,16) = EVServer17
	EVServersList(2,16) = Datacenter17
	EVServersList(3,16) = DBServer17
	EVServersList(4,16) = EVDatabase17
	EVServersList(5,16) = Role17
End If

If EVServer18 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,17) = DNSAlias18
	EVServersList(1,17) = EVServer18
	EVServersList(2,17) = Datacenter18
	EVServersList(3,17) = DBServer18
	EVServersList(4,17) = EVDatabase18
	EVServersList(5,17) = Role18
End If

If EVServer19 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,18) = DNSAlias19
	EVServersList(1,18) = EVServer19
	EVServersList(2,18) = Datacenter19
	EVServersList(3,18) = DBServer19
	EVServersList(4,18) = EVDatabase19
	EVServersList(5,18) = Role19
End If

If EVServer20 <> "" then
	Redim preserve EVServersList(ubound (EVServersList, 1), ubound (EVServersList, 2) + 1)
	EVServersList(0,19) = DNSAlias20
	EVServersList(1,19) = EVServer20
	EVServersList(2,19) = Datacenter20
	EVServersList(3,19) = DBServer20
	EVServersList(4,19) = EVDatabase20
	EVServersList(5,19) = Role20
End If
'wscript.echo ubound(EVServersList,2) 'Output number of rows in the MD Array. 
'* End build multidimential array.