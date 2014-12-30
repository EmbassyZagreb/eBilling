<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<html>
<head>

<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
	<script src="jquery-latest.js" type="text/javascript"></script>
	<script src="jquery.tablesorter.js" type="text/javascript"></script>
	<script src="jquery.tablesorter.pager.js" type="text/javascript"></script>
	<link href="style.css" rel="stylesheet" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1250" />
	<script type="text/javascript">
	$(function() {
		$("#myTable").tablesorter({headers: { 5:{sorter: false}}, widgets: ['zebra']});
	});
	</script>
</HEAD>
<body>

<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Import New Bill</TD>
   </TR>
	<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
	</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<%
dim transID
dim rs, bg

%>
<div>

<%
transID = request.querystring("ID")

if isempty(transID)=false then
%>

<form action="TranslationSave.asp" id="updateform" method="GET" name="updateform">
 <input type="hidden" name="ID" value="<%=transID%>">

<%
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open "SELECT * from CallTypeTranslation  where TransID = " & transID & ";", BillingCon, 1,3
  if bg="#FFCC99" then bg="ffffff" else bg="#FFCC99"
%>
<table border="1" bordercolor="#EEEEEE" cellpadding="1" cellspacing="1" class="tablesorter" id="myTable" >
  </tr>
  <TR bgcolor="<%=bg%>" width="60%">
   <td><b>Original text</b></td>
   <td><%=rs("Croatian")%></td>
  </tr>
  <tr>
   <td><b>English text</b></td>
   <td><input type="text" size="100" name="translation" value="<%=rs("english")%>"></td>
  </tr>
 </table>
<br>
 <table>
  <input type="submit" name="Submit" value="Submit">

  <button type="cancel" onclick="window.location='TranslationTable.asp?sort=<%=sorting%>';return false;">Cancel</button>
 </table>
</form>
<%
 rs.close
 end if

%>
</div>
  <tr>
    <button type="submit" name="back" onclick="window.location='ImportSpecView.asp';return false;">Back</button>
  </tr>
<div>
<form action="" id="form1" name="form1" method="GET">
<%
 Set rs = Server.CreateObject("ADODB.Recordset")
 rs.Open "SELECT * from CallTypeTranslation order by English;", BillingCon, 1,3
%>




<table border="1" bordercolor="#EEEEEE" cellpadding="1" cellspacing="1" class="tablesorter" id="myTable">
 <thead>
  <tr>
   <th>Original text</th>
   <th></th>
   <th>English Translation</th>
  </tr>
</thead>
<tbody>
<%
 rs.movefirst
 do  until rs.eof
 if bg="#FFCC99" then bg="ffffff" else bg="#FFCC99" 
%>

  <TR bgcolor="<%=bg%>">
   <TD><%=rs("Croatian")%></TD>
   <td><A HREF="TranslationTable.asp?ID=<%=rs("TransID")%>">EDIT</A></TD>
   <TD><%=rs("english")%></TD>
  </tr><% 
 rs.movenext
 loop
rs.close()
%>
</tbody>
</table>
</form>
</div>
<br>
<br>


</BODY>
</html>