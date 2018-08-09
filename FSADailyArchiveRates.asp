<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<br><br>
<h1>FSA Daily Archive Rates</h1>
<br><br>
<table id="table1" class="mytable" align="center">
<thead>
<tr>
<th>EV Server</th>
<th>DB Server</th>
<th>EV Database</th>
<th>Archive Date</th>
<th>Daily Rate</th>
<th>Average File Size</th>
<th>Total File Size</th>
</tr>
</thead>
<tbody>
<%

For i = 0 to Ubound(EVServersList,2)
	If Left(EVServersList(5,i),4) = "File" then
		'Retrieve data from FSA database.
		If ConnectionString = 1 Then
			Connection1 = "Driver={SQL Server};Server=" & EVServersList(3,i) & ";Database=" & _
				EVServersList(4,i) & ";Uid=" & SQLUserName & ";Pwd=" & SQLPassword & ";Port=" & SQLAltPortNumber & ";"
		Elseif ConnectionString = 2 Then
			Connection1 = "Driver={SQL Server};Server=" & EVServersList(3,i) & ";Database=" & _
				EVServersList(4,i) & ";Trusted_Connection=yes;Port=" & SQLAltPortNumber & ";"
		End if

		'* Open database connection
		Set myconn1 = CreateObject("adodb.connection")
		myconn1.ConnectionTimeout = strSQLConnect
		myconn1.open(Connection1)

		strSQLQuery1 = _
			"select Dates.TheDate as 'ArchDate', " & _
			"	LEFT(CONVERT(VarChar, CAST(count (SaveSet.IndexSeqNo) as Money), 1), LEN(CONVERT(VarChar, CAST(count (*) as Money), 1))-3) AS 'DailyRate', " & _
			"	ISNULL(LEFT(CONVERT(VarChar, CAST((sum (itemsize)/count (*)) as Money), 1), LEN(CONVERT(VarChar, CAST((sum (itemsize)/count (*)) as Money), 1))-3), 0) AS 'AvSizeKB', " & _
			"	ISNULL(LEFT(CONVERT(VarChar, CAST(sum (itemsize) as Money), 1), LEN(CONVERT(VarChar, CAST(sum (itemsize) as Money), 1))-3), 0) AS 'DailyKBs' " & _
			"from ( " & _
			"		select LEFT(Convert(VarChar, dateadd(day, 0-number, GETDATE()), 20), 10) as 'TheDate' " & _
			"		from  " & _
			"		    (select distinct number from master..spt_values " & _
			"			 where name is null and number < 31 " & _
			"		    ) n " & _
			"	) AS Dates  " & _
			"	LEFT OUTER JOIN SaveSet on LEFT(CONVERT (varchar, archiveddate, 20), 10) = Dates.TheDate " & _
			"group by Dates.TheDate  " & _
			"order by ArchDate desc"

		Set objCmd1 = CreateObject("adodb.command")
		objCmd1.activeconnection = myconn1
		objCmd1.commandtimeout = strSQLExecute
		objCmd1.commandtype = adCmdText
		objCmd1.commandtext = strSQLQuery1

		Set result1 = CreateObject("adodb.Recordset")
		result1.open objCmd1
		If Err.Number <> 0 Then
			Response.Write "<!-- Error [" & Err.Number & "] : " & Err.Description & " -->"
		End If		
		While not result1.EOF
			Response.write ("<tr><td>" & EVServersList(0,i) & _
					"</td><td>" & EVServersList(3,i) & _
					"</td><td>" & EVServersList(4,i) & _
					"</td><td align=""center"">" & (result1("ArchDate")) & _
					"</td><td align=""right"">" & (result1("DailyRate")) & _
					"</td><td align=""right"">" & (result1("AvSizeKB")) & _
					"</td><td align=""right"">" & (result1("DailyKBs")) & _
					"</td></tr>")
			result1.movenext()
		Wend
	End if
Next

%>
</tbody>
</table><br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		on_keyup: true,
		on_keyup_delay: 1200,
		filters_row_index: 1,
		col_0: "select",  
		col_1: "select",  
		col_2: "select",  
		sort: true,
		sort_config: {
			sort_types:['String','String','String','String','Number','Number','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->
