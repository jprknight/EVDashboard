'********************************************************************************
'* Organisation Archiving History Start
'********************************************************************************
If LoggingLevel >= 1 then
	wscript.echo DateTime() & ": Processing Organisation Archiving History."
	OutputLogFile.writeline ""
	OutputLogFile.writeline "*************************************************************************"
	OutputLogFile.writeline ""
	OutputLogFile.writeline DateTime() & ": Processing Organisation Archiving History."
End if

On Error Resume Next

For i = 0 to Ubound(EVServersList,2)
	LinkedServerError = 0
	If Len(EVServersList(4,i)) > 4 then
		If instr(EVServersList(5,i),"Email") then
			'If EVServersList(3,i) = DirectoryServer then
				' Database exists on the same SQL server as the EnterpriseVaultDirectory database / DirectoryService.
				wscript.echo DateTime() & ": Processing Organisation Archiving History - " & EVServersList(4,i) & "."
				OutputLogFile.writeline DateTime() & ": Processing Organisation Archiving History - " & EVServersList(4,i) & "."
				sSQL = "SELECT left (convert (nvarchar(30), s.archiveddate,112),6) AS ArchivedDate, count (*) AS RecordCount " & _
					"FROM Saveset s " & _
					"GROUP BY left (convert (nvarchar(30), s.archiveddate,112),6) " & _
					"ORDER BY left (convert (nvarchar(30), s.archiveddate,112),6) desc"
				
				If ConnectionString = 1 Then
					connection = "Driver={SQL Server};Server=" & DirectoryServer & ";Database=" & _
						EVServersList(4,i) & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
				Elseif ConnectionString = 2 Then
					connection = "Driver={SQL Server};Server=" & EVServersList(3,i) & ";Database=" & EVServersList(4,i) & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
				End if
				
				'* Open database connection
				Set myconn = CreateObject("adodb.connection")
				myconn.open (connection)

				If cstr(err.number) <> 0 Then
					wscript.echo (DateTime () & ": Error creating connection to " _
						& "database server " & EVServersList(3,i) & " / " & EVServersList(4,i) & ".  Check your connection string " _
						& "or database server name/IP and try again.")
					OutputLogFile.writeline (DateTime() & ": Error creating connection to " _
						& "database server " & EVServersList(3,i) & " / " & EVServersList(4,i) & ".  Check your connection string " _
						& "or database server name/IP and try again.")
					wscript.quit
				End if
					
				On Error Resume Next
					
				Set result = CreateObject("adodb.recordset")
				If err.number <> 0 then 
					wscript.echo DateTime() & ": " & Err.Description
					OutputLogFile.writeline DateTime() & ": " & Err.Description
				End if

				'************************************
				'* Execute the query
				Set result = myconn.execute(sSQL)
				If err.number <> 0 then 
					wscript.echo DateTime() & ": " & Err.Description
					OutputLogFile.writeline DateTime() & ": " & Err.Description
				End if
				
				'************************************
				'* Write data.
				While not result.EOF
					
					If EVDConnectionString = 1 Then
						EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
							EVDSQLDatabase & ";Uid=" & EVDSQLUserName & ";Pwd=" & EVDSQLPassword & ";Port=" & SQLEVDAltPortNumber & ";"
					Elseif EVDConnectionString = 2 Then
						EVDconnection = "Driver={SQL Server};Server=" & EVDSQLServer & ";Database=" & _
							EVDSQLDatabase & ";Trusted_Connection=yes;Port=" & SQLEVDAltPortNumber & ";"
					End if
					
					'* Open database connection
					Set EVDmyconn = CreateObject("adodb.connection")
					EVDmyconn.open (EVDconnection)

					If cstr(err.number) <> 0 Then
						wscript.echo (DateTime() & ": Error creating connection to " _
						   & "database server " & EVDSQLServer & " / " & EVDSQLDatabase & ".  Check your connection string " _
						   & "or database server name/IP and try again.")
					   OutputLogFile.writeline (DateTime() & ": Error creating connection to " _
						   & "database server " & EVDSQLServer & " / " & EVDSQLDatabase & ".  Check your connection string " _
						   & "or database server name/IP and try again.")
					   wscript.quit
					End if
						
					On Error Resume Next
						
					Set EVDSQL = CreateObject("adodb.recordset")
					If err.number <> 0 then 
						wscript.echo DateTime() & ": " & Err.Description
						OutputLogFile.writeline DateTime() & ": " & Err.Description
					End if

					EVDSQLQuery = "UPDATE dbo.OrganisationArchivingHistory " & _
						"SET ArchivedDate='" & (result("ArchivedDate")) & "'," & _
						"DNSAlias='" & EVServersList(0,i) & "'," & _
						"EVServer='" & EVServersList(1,i) & "'," & _
						"EVDatabase='" & EVServersList(4,i) & "'," & _
						"RecordCount='" & (result("RecordCount")) & "' " & _
						"WHERE ArchivedDate='" & (result("ArchivedDate")) & "' AND EVDatabase='" & EVServersList(4,i) & "' " &  _
						"IF @@ROWCOUNT=0 " & _
						"INSERT INTO dbo.OrganisationArchivingHistory (ArchivedDate,DNSAlias,EVServer,EVDatabase,RecordCount) VALUES ('" & _
						(result("ArchivedDate")) & "','" & EVServersList(0,i) & "','" & _
						EVServersList(1,i) & "','" & EVServersList(4,i) & "','" & _
						(result("RecordCount")) & "')"
					
					'* Execute the query
					Set EVDSQL = EVDmyconn.execute(EVDSQLQuery)
					
					If err.number <> 0 then 
						wscript.echo DateTime() & ": " & Err.Description
						OutputLogFile.writeline DateTime() & ": " & Err.Description
					End if

					EVDmyconn.close
					
					result.movenext()
				Wend
			'end if
		end if
	End if
Next
'********************************************************************************
'* Organisation Archiving History End
'********************************************************************************