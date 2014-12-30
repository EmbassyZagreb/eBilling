<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<html>
   <head>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<%
PrevEmpID_ = request.form("cmbEmpFrom")
CurEmpID_ = request.form("cmbEmpTo")

user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>

<meta http-equiv="refresh" content="1;url=Default.asp">

<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">UPDATE SUPERVISOR UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<table border=0 width=100%>
<%
       strsql = "Exec spUpdateSupervisor '" & PrevEmpID_ & "','" & CurEmpID_ & "','" & user1_ & "'"
       'response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td align=center>Supervisor name has been updated.</td></tr>
<tr><td>&nbsp;</td>
<!--
<tr><td align=center> 
<input type="button" value="Close" id="btnclose">
</td></tr>
<tr>
	<td align="center"><br><a href="ExchangeRateList.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
-->
</table>

   </body>
</html>