<!--#include file="./header.asp"-->
<!--#include file="./global.asp"-->
<!--#include file="./includes/styles.asp"-->
<!--#include file="./includes/BuildMultiDimensionalArray.asp"-->
<!--#include file="./includes/FC_Colors.asp"-->
<!--#include file="./includes/FusionCharts.asp" -->
<table class="Main">
<thead></thead>
<tbody>

<tr>
	<td valign="top">
	<!--#include file="./IndexEnterpriseVaultEnvironment.asp"-->
	</td>

	<td valign="top">
	<!--#include file="./IndexExchangeEnvironment.asp"-->
	</td>
</tr>

<tr>
	<td valign="top">
	<!--#include file="./IndexServiceAlerts.asp"-->
	</td>


	<td valign="top">
	<!--#include file="./IndexUserIndexes.asp"-->
	</td>
</tr>

<tr>
	<td valign="top">
	<!--#include file="./IndexMSMQMessageCountAlerts.asp"-->
	</td>

	<td valign="top">
	<!--#include file="./IndexWatchFileCounts.asp"-->
	</td>
</tr>

<tr>
	<td valign="top">
	<!--#include file="./IndexBackupTriggerFile.asp"-->
	</td>
	<td valign="top">
	<!--#include file="./IndexJournalArchiveCounts.asp"-->
	</td>
</tr>

<tr>
	<!--#include file="./IndexBuildXML.asp"-->
	<td valign="top">
	<span>
	<br><br>
	<h1>EV Mailbox Archiving Users (Total <%=TotalUsers%>)</h1><br>
	<table class="Chart">
		<tr> 
			<td valign="top" class="text" align="left">
				<%
				Call renderChartHTML("./includes/FusionCharts/FCF_Column3D.swf", "", strXML, "myNext", 625, 450)
				%>
			</td>
		</tr>
	</table><br>
	</span>
	</td>
	<td valign="top">
	<!--#include file="./IndexDiskAlerts.asp"-->
	</td>
</tr>


<br><br>
</tbody>
</table>
<table class="Main">
<tr>
	<td>
	<!--#include file="./IndexEVServerEventLogs.asp"-->
	</td>
</tr>
</table>

<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {   
		col_3: "select",    
		display_all_text: " [ Show all ] ",  
		sort: true,
		sort_config: {
			sort_types:['String','String','String','String','String','String']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
	//<![CDATA[    
	var table2_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','Number']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table2",table2_Props );  
	//]]>  
	//<![CDATA[    
	var table3_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table3",table3_Props );  
	//]]>  
		//<![CDATA[    
	var table4_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table4",table4_Props );  
	//]]>  
		//<![CDATA[    
	var table5_Props =  {
		col_1: "select",    
		display_all_text: " [ Show all ] ",  
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table5",table5_Props );  
	//]]>  
			//<![CDATA[    
	var table6_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table6",table6_Props );  
	//]]>  
				//<![CDATA[    
	var table7_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','Number','Number','Number']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table7",table7_Props );  
	//]]>  
		var table8_Props =  {
		col_1: "select",    
		col_2: "select",    
		col_3: "select",    
		display_all_text: " [ Show all ] ",  
		sort: true,
		sort_config: {
			sort_types:['String','String','Number','String','String']
		},
		alternate_rows: true
		};  
		setFilterGrid( "table8",table8_Props );  
	//]]>  
	var table9_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table9",table9_Props );  
	//]]>
		var table10_Props =  {
		sort: true,
		sort_config: {
			sort_types:['String','String','String','Number']
		},
		alternate_rows: true,
		sort_select: true
		};  
		setFilterGrid( "table10",table10_Props );  
	//]]>
</script>
<!--#include file="./footer.asp"-->