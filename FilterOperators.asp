<!--#include file="./header.asp"-->
<!--#include file="./config.asp"-->
<!--#include file="./includes/styles.asp"-->
<h1>Filter Operators</h1>
<br>
<table id="table1" class="mytable">
<br>
<thead>
<tr>
	<th nowrap align="left">Operator</th>
	<th nowrap align="left">Description</th>
	<th nowrap align="left">Type</th>
	<th nowrap align="left">Example</th>
</tr>
</thead>
<tbody>
<tr>
<td><</td><td>Values lower than search term are matched</td><td>number</td><td><1412</td>
</tr>
<tr>
<td><=</td><td>Values lower than or equal to search term are matched</td><td>number</td><td><=1412</td>
</tr>
<tr>
<td>></td><td>Values greater than search term are matched</td><td>number</td><td>>1412</td>
</tr>
<tr>
<td>>=</td><td>Values greater than or equal to search term are matched</td><td>number</td><td>>=1412</td>
</tr>
<tr>
<td>=</td><td>Exact match search: only the whole search term(s) is matched</td><td>string / number</td><td>=Sydney</td>
</tr>
<tr>
<td>*</td><td>Data containing search term(s) is matched (default operator)</td><td>string / number</td><td>*Syd</td>
</tr>
<tr>
<td>!</td><td>Data that doesn't contain search term(s) is matched</td><td>string / number</td><td>!Sydney</td>
</tr>
<tr>
<td>{</td><td>Data starting with search term is matched</td><td>string / number</td><td>{S</td>
</tr>
<tr>
<td>}</td><td>Data ending with search term is matched</td><td>string / number</td><td>}y</td>
</tr>
<tr>
<td>||</td><td>Data containing at least one of the search terms is matched</td><td>string / number</td><td>Sydney || Adelaide</td>
</tr>
<tr>
<td>&&</td><td>Data containing search terms is matched</td><td>string / number</td><td>>4.3 && <25.3</td>
</tr>
</table>
<br>
<h3>Taken straight from <a href="http://tablefilter.free.fr/" target="_blank">http://tablefilter.free.fr/</a>.</h3>
<br>
<script language="javascript" type="text/javascript">  
	//<![CDATA[    
	var table1_Props =  {
		alternate_rows: true
		};  
		setFilterGrid( "table1",table1_Props );  
	//]]>  
</script> 
<!--#include file="footer.asp"-->