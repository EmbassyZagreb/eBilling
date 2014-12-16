<%@ Language=VBScript %>
<%Server.ScriptTimeout=600%>
<!--#include file="connect.inc" -->

<html>
<head>

<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
	<script src="jquery-latest.js" type="text/javascript"></script>
	<script src="jquery.tablesorter.js" type="text/javascript"></script>
	<link rel="stylesheet" type="text/css" href="style-tablesorter.css" />
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1250" />
	<script type="text/javascript">
	$(function() {
		$("#myTable").tablesorter({headers: { 0:{sorter: false}}, widgets: ['zebra']});
	});
	</script>
</head>
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
dim bg, sort, translation, rows

rows=request.querystring ("rows")
if isempty(rows)=true then rows="100"
translation=request.querystring ("translation")
if isempty(translation)=true then translation="english"
sort="PhoneNumber"

%>
<table>
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs2 = Server.CreateObject("ADODB.Recordset")
Rs.open ("qCallTypeMissing"), BillingCon,1,3
Rs2.open ("qMissingCallDescription"), BillingCon,1,3

if rs.recordcount<>0 then
   response.write "New call types updated."
   Set objExec = BillingCon.Execute("spAddNewCallType")
%>
<button type="submit" onclick="window.location='TranslationTable.asp';return false;">Update call desctriptions</button>
<%

else if rs2.recordcount<>0 then%>
     <button type="submit" onclick="window.location='TranslationTable.asp';return false;">Update call desctriptions</button>  
     <%rs2.close() 
   else
%>No new call types. <button type="submit" onclick="window.location='ImportSaveFinal.asp';return false;">Final Import</button><%
   end if
end if
rs.close()
response.write "</table>"

Set Rs = Server.CreateObject("ADODB.Recordset")
rs.maxrecords=rows

Rs.open ("select * from vwImportView order by " & sort), BillingCon,1,3

if rs.recordcount=0 then
else
%>

<form action="" method="GET" name="Submit">

<table border="0" cellpadding="1" cellspacing="1" class="tablesorter"> 
  <tr>
    <td>Import Main</td>
    <td>Number of rows to display: <select name="rows" onchange="this.form.submit();">
<%
Response.Write "<option value='100'"
       if rows = 100 then
         Response.Write " selected"
        end if
Response.Write ">100</option>"
Response.Write "<option value='1000'"
       if rows = 1000 then
         Response.Write " selected"
       end if
Response.Write ">1000</option>"
Response.Write "<option value='10000'"
       if rows = 10000 then
         Response.Write " selected"
       end if
Response.Write ">10000</option>"
%>
</select>
    <td>Change call descriptions to: <select name="translation" onchange="this.form.submit();">
<%
Response.Write "<option value='english'"
       if translation = "english" then
         Response.Write " selected"
         translation = "English"
       end if
Response.Write ">English</option>"
Response.Write "<option value='croatian'"
       if translation = "croatian" then
         Response.Write " selected"
         translation = "Calltype"
       end if
Response.Write ">Croatian</option>"
%>
</select>
  </tr>
</table>

<table border="1" bordercolor="#EEEEEE" cellpadding="1" cellspacing="1" class="tablesorter" id="myTable" class="tablesorter"> 
 <thead>
  <tr>
   <th>#</th>
   <th>Month</th>
   <th>Year</th>
   <th>Phone Number</th>
   <th>Dialed Date & Time</th>
   <th>Call Description</th>
   <th>Dialed Number</th>
   <th>Duration</th>
   <th>Cost</th>
  </tr>
 </thead>

 <tfoot>
  <tr>
   <th>#</th>
   <th>Month</th>
   <th>Year</th>
   <th>Phone Number</th>
   <th>Dialed Date & Time</th>
   <th>Call Description</th>
   <th>Dialed Number</th>
   <th>Duration</th>
   <th>Cost</th>
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
    response.write"<td>" & rs("PhoneNumber") & "</td>"
    response.write"<td>" & rs("DialedDateTime") & "</td>"
    response.write"<td>" & rs(translation) & "</td>"
    response.write"<td>" & rs("DialedNumber") & "</td>"
    response.write"<td>" & rs("CallDuration") & "</td>"
    response.write"<td>" & rs("Cost") & "</td>"
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
</form>
</body>
</html>
