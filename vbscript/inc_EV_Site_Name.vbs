'********************************************************************************
'* EV SITE NAME START
'********************************************************************************
ON ERROR RESUME NEXT

If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing EV Site Name."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing EV Site Name."
End if

If ConnectionString = 1 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
Elseif ConnectionString = 2 Then
	Connection1 = "Driver={SQL Server};Server=" & DirectoryServer & _
		";Database=EnterpriseVaultDirectory;Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
End if

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (EV Site Name) - Connection1" & vbcrlf & Connection1
	OutputLogFile.writeline DateTime() & ": (EV Site Name) - Connection1" & vbcrlf & Connection1
End if

'* Open database connection
Set myconn = CreateObject("adodb.connection")
myconn.open (Connection1)

If cstr(err.number) <> 0 Then
	wscript.echo (DateTime() & ": Error creating connection to " _
	   & "database server " & DirectoryServer & " / EnterpriseVaultDirectory.  Check your connection string " _
	   & "or database server name/IP and try again.")
	OutputLogFile.writeline (DateTime() & ": Error creating connection to " _
	   & "database server " & DirectoryServer & " / EnterpriseVaultDirectory.  Check your connection string " _
	   & "or database server name/IP and try again.")
   wscript.quit
End if
	
On Error Goto 0
	
Set result = CreateObject("adodb.recordset")
If err.number <> 0 then 
	wscript.echo DateTime() & ": " & Err.Description
	OutputLogFile.writeline DateTime() & ": " & Err.Description
End if

SQLQuery = "Use EnterpriseVaultDirectory Select sitename AS SiteName from siteentry"

'* Execute the query
Set result = myconn.execute(SQLQuery)
If err.number <> 0 then 
	wscript.echo DateTime() & ": " & Err.Description
	OutputLogFile.writeline DateTime() & ": " & Err.Description
End if

While not result.EOF

	SiteName = (result("SiteName"))

	set xmlDoc = createobject("Microsoft.XMLDOM")
	xmlDoc.async = false
	xmlDoc.load strConfigurationFile

	If err.number <> 0 then 
	Call LogEvent("Error occured in XMLSave - Error number: " & err.number & vbcrlf & "Description " & err.description,1)
	End IF

	'* Save data to config.xml

	set ServerConfig = xmlDoc.documentelement.selectSingleNode("//EVDashboard/ServerConfiguration")

	ServerConfig.setAttribute("SiteName") = SiteName

	xmlDoc.save strConfigurationFile

	set xmlDoc = nothing

	result.movenext()
Wend

'* NORMAL LOGGING.
If LoggingLevel >= 2 then
	'wscript.echo DateTime() & ": (EV Site Name) - EV Site name set to " & SiteName
	OutputLogFile.writeline DateTime() & ": (EV Site Name) - EV Site name set to " & SiteName
End if

'********************************************************************************
'* EV SITE NAME END
'********************************************************************************
