<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<script language="JavaScript" src="calendar.js"></script>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">ADMIN PAGE(S)</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

if not RS_Query.eof then
	UserRole_ = RS_Query("RoleID")
end if
%>  
<table cellspacing="0" cellpadding="2">  
<%if (UserRole_= "Admin") then %>
<tr>
	<td>
		<UL TYPE="square">
		       <LI><A HREF="UserList.asp"><B>Manage User(s)</B></A>
		</UL>
	</td>
</tr>
<%else %>

<tr>
	<td>You do not have permission to access this site.</td>
</tr>
<tr>
	<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
</tr>
<%end if %>
</table>
</BODY>
</html>