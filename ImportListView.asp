<%@ Language=VBScript %>
<%Server.ScriptTimeout=600%>
<!--#include file="connect.inc" -->

<html>
<head>

<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
	<script src="jquery-latest.js" type="text/javascript"></script>
	<script src="jquery.tablesorter.js" type="text/javascript"></script>
	<script src="jquery.tablesorter.pager.js" type="text/javascript"></script>

	<link rel="stylesheet" type="text/css" href="style-tablesorter.css" />
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1250" />
	<script type="text/javascript">
	$(function() {
		$("#myTable").tablesorter({headers: { 0:{sorter: false}}, widgets: ['zebra']});
	});
	</script>
</HEAD>

<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Import New Bill</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<body>
<Center>
<%

Dim objFSO
Dim Conn, objExec
dim rs, rs2, updateSQL
dim bg, sort


response.write "</table>"
Set Rs = Server.CreateObject("ADODB.Recordset")
'rs.maxrecords=100
Rs.open ("select * from ListTEMP  order by Field2;"), BillingCon,1,3

if rs.recordcount=0 then

else
%>


<table border="0" cellpadding="1" cellspacing="1" class="tablesorter"> 
<tr><td>
<a href="ImportSpec.asp">Continue to import specification file</a>
</td></tr>
</table>

<table border="1" bordercolor="#EEEEEE" cellpadding="1" cellspacing="1" class="tablesorter" id="myTable" class="tablesorter"> 
 <thead>
  <tr>
   <th>#</th>
   <th>Month</th>
   <th>Year</th>
   <th>Phone Number</th>
   <th>3</th>
   <th>4</th>
   <th>5</th>
   <th>6</th>
   <th>7</th>
   <th>8</th>
   <th>9</th>
   <th>10</th>
   <th>11</th>
   <th>12</th>
   <th>13</th>
   <th>14</th>
   <th>15</th>
   <th>16</th>
   <th>17</th>
   <th>18</th>
   <th>19</th>
  </tr>
 </thead>

 <tfoot>
  <tr>
   <th>#</th>
   <th>Month</th>
   <th>Year</th>
   <th>Phone Number</th>
   <th>3</th>
   <th>4</th>
   <th>5</th>
   <th>6</th>
   <th>7</th>
   <th>8</th>
   <th>9</th>
   <th>10</th>
   <th>11</th>
   <th>12</th>
   <th>13</th>
   <th>14</th>
   <th>15</th>
   <th>16</th>
   <th>17</th>
   <th>18</th>
   <th>19</th>
  </tr>
 </tfoot>

 <tbody>
<%
dim i
i=1
 rs.movefirst
  do until rs.eof
    if bg="#FFCC99" then bg="ffffff" else bg="#FFCC99"%>
    <tr bgcolor="<%=bg%>"><%
    response.write"<td>" & i & "</td>"
    response.write"<td>" & rs("MonthP") & "</td>"
    response.write"<td>" & rs("YearP") & "</td>"
    response.write"<td>" & rs("Field2") & "</td>"
    response.write"<td>" & rs("Field3") & "</td>"
    response.write"<td>" & rs("Field4") & "</td>"
    response.write"<td>" & rs("Field5") & "</td>"
    response.write"<td>" & rs("Field6") & "</td>"
    response.write"<td>" & rs("Field7") & "</td>"
    response.write"<td>" & rs("Field8") & "</td>"
    response.write"<td>" & rs("Field9") & "</td>"
    response.write"<td>" & rs("Field10") & "</td>"    
    response.write"<td>" & rs("Field11") & "</td>"
    response.write"<td>" & rs("Field12") & "</td>"
    response.write"<td>" & rs("Field13") & "</td>"
    response.write"<td>" & rs("Field14") & "</td>"
    response.write"<td>" & rs("Field15") & "</td>"
    response.write"<td>" & rs("Field16") & "</td>"
    response.write"<td>" & rs("Field17") & "</td>"
    response.write"<td>" & rs("Field18") & "</td>"
    response.write"<td>" & rs("Field19") & "</td>"
    response.write"</tr>"
    i=i+1
   rs.movenext
  loop
end if
%>
</tbody>

</table>

<br>

<%
rs.close()
 Billingcon.close()
set rs = nothing

%>
</body>
</html>
