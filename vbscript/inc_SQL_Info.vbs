'********************************************************************************
'* EV DASHBOARD SQL INFO START
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing SQL info."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing SQL info."
End if

If EVDConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
Elseif EVDConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
		EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (SQL info) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (SQL info) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn1 = CreateObject("adodb.connection")
myconn1.ConnectionTimeout = strSQLConnect
myconn1.open(Connection1)

strSQLQuery1 = "SELECT SERVERPROPERTY('productversion') AS Version," & _
	"SERVERPROPERTY ('productlevel') AS ServicePack," & _
	"SERVERPROPERTY ('edition') AS Edition"

'* VERBOSE LOGGING.
If LoggingLevel >= 3 then
	'wscript.echo DateTime() & ": (SQL info) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (SQL info) - strSQLQuery1" & vbcrlf & strSQLQuery1
End if
	
Set objCmd1 = CreateObject("adodb.command")
objCmd1.activeconnection = myconn1
objCmd1.commandtimeout = strSQLExecute
objCmd1.commandtype = adCmdText
objCmd1.commandtext = strSQLQuery1

Set result1 = CreateObject("adodb.Recordset")
result1.open objCmd1

While not result1.EOF
	strEVDSQLVersion = (result1("Version"))
	strEVDSQLServicePack = (result1("ServicePack"))
	strEVDSQLEdition = (result1("Edition"))
	result1.movenext()
Wend

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (SQL info) - strSQLQuery1" & vbcrlf & strSQLQuery1
	OutputLogFile.writeline DateTime() & ": (SQL info) - strEVDSQLVersion - " & strEVDSQLVersion
	OutputLogFile.writeline DateTime() & ": (SQL info) - strEVDSQLServicePack - " & strEVDSQLServicePack
	OutputLogFile.writeline DateTime() & ": (SQL info) - strEVDSQLEdition - " & strEVDSQLEdition
End if

Err.clear

set xmlDoc = createobject("Microsoft.XMLDOM")
xmlDoc.async = false
xmlDoc.load strConfigurationFile

'* Save data to config.xml

set ServerConfig = xmlDoc.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")

ServerConfig.setAttribute("EVDSQLVersion") = strEVDSQLVersion
ServerConfig.setAttribute("EVDSQLServicePack") = strEVDSQLServicePack
ServerConfig.setAttribute("EVDSQLEdition") = strEVDSQLEdition

xmlDoc.save strConfigurationFile

If err.number <> 0 then 
	wscript.echo "Error occured in XMLSave - Error number: " & err.number & " Error Description: " & err.description
End IF

set xmlDoc = nothing


'********************************************************************************
'* EV DASHBOARD SQL INFO END
'********************************************************************************
