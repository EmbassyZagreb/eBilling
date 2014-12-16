<%@ Language=VBScript %>
<!--#include file="connect.inc" -->


<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<style>
table, td, th {border: 0px solid black;font-family:'Tahoma', Georgia, Serif;text-align:center}
td {font-size:'30px'}
table {width: 500px;}
th {height: 50px;}
</style>
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Import New Bill</TD>
   </TR>
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
user1_ = user_  'user1_ = right(user_,len(user_)-4)
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

if not RS_Query.eof then
	UserRole_ = RS_Query("RoleID")
end if
%>
<BODY>

<div>
  <%if (UserRole_= "Admin") then %>
   <table>
    <tr>
      <th>Please select vendor you want to import data:</th>
    </tr>
   </table>
  <br>
   <table>
    <tr>
      <td><input type="button" name="B1" value="VIPNET" onclick="window.location.href='ImportList.asp'"></td>
      <td><button name="B2">T-COM</button></td>
    </tr>
  </table>
</div>
<%else %>
  <div>
  	You do not have permission to access this site.
  </div>
  <div>
  	<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
<  /div>
<%end if %>
</BODY>
</HTML>