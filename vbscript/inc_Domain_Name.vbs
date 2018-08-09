'********************************************************************************
'* DOMAIN NAMESTART
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Domain Name."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Domain Name."
End if

Set WshNetwork = WScript.CreateObject("WScript.Network")
'WScript.Echo "Domain = " & WshNetwork.UserDomain
'WScript.Echo "Computer Name = " & WshNetwork.ComputerName
'WScript.Echo "User Name = " & WshNetwork.UserName

wscript.echo strConfigurationFile

set xmlDoc = createobject("Microsoft.XMLDOM")
xmlDoc.async = false
xmlDoc.load strConfigurationFile

set ServerConfig = xmlDoc.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")

ServerConfig.setAttribute("DomainName") = WshNetwork.UserDomain
ServerConfig.setAttribute("ServiceAccount") = WshNetwork.UserName

xmlDoc.save strConfigurationFile

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (Domain Name) - Domain name set to " & WshNetwork.UserDomain
	OutputLogFile.writeline DateTime() & ": (Domain Name) - Domain name set to " & WshNetwork.UserDomain
End if

'********************************************************************************
'* DOMAIN NAME FINISH
'********************************************************************************