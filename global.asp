<!--#include file="./config.asp"-->
<% 
set fso = CreateObject("Scripting.FileSystemObject")
Session.Timeout = 60

on error goto 0

Dim xmlDoc1, ServerConfig

set xmlDoc1 = server.createobject("Microsoft.XMLDOM")

xmlDoc1.async = false
xmlDoc1.load strConfigurationFile

set ServerConfig = xmlDoc1.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")

Const adCmdText = 1
strSQLConnect = "300"
strSQLExecute = "300"

if lcase(ServerConfig.getAttribute("Setup")) = "true" then
	set xmlDoc1 = nothing
	response.redirect "admin.asp"
End If
%>