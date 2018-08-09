<%
'	Build dynamic multidimential array of EV Servers, DB Servers, DB Databases and Roles.

'	EVServersList(0,i)	DNS Alias
'	EVServersList(1,i)	EV Server
'	EVServersList(2,i)	Data Center
'	EVServersList(3,i)	DB Server
'	EVServersList(4,i)	EV Database
'	EVServersList(5,i)	EV Server Role
'	EVServersList(6,i)	EV MSMQ Path
'	EVServersList(7,i)	TriggerPath


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
%>