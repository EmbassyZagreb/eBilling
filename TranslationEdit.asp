<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
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
<form action="TranslationSave.asp" method="GET">

<%
dim transID
transID = request.querystring("ID")
%>
<input type="hidden" name="ID" value="<%=transID%>"

<br>
<table>
<b>Change or update translation here. Only english translation will be visible in this application.</b>
</table>
<br>
<%

dim rs, bg
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * from CallTypeTranslation  where TransID = " & transID & ";", BillingCon, 1,3
if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4"
%>
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0">
<TR bgcolor="<%=bg%>" width="60%">
<td><b>Croatian text</b></td>
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
<button type="cancel" onclick="window.location='TranslationTable.asp';return false;">Cancel</button>
</table>
<%
rs.close
%>

</BODY>
</html>