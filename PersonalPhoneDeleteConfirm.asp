<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% 
   ID_ = request("ID")
   State_ =  trim(request("State"))
'response.write ID_

        strsql = " select * from MsPersonalPhone where ID = " & ID_
  	set rsData = server.createobject("adodb.recordset") 
	'response.write strsql 
       	set rsData = BillingCon.execute(strsql)
       	if not rsData.eof then 
      		 PhoneNumber_ = rsData("PhoneNumber") 
		 Remark_ = rsData("Remark")
  	end if
%>

<html>
   <head>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Delete Personal Phone</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" action="PersonalPhoneDelete.asp?ID=<%=ID_%>"> 
<table cellpadding="1" cellspacing="0" width="50%" align="center">  
<tr>
	<td colspan="3"><b>Are you sure delete this data ?</b></td>
</tr>   
<tr>
	<td colspan="3"><br></td>
</tr>
<tr>
	<td width="100px">Phone Number </td>
	<td width="1px">:</td>
	<td><font color=blue><strong><%=PhoneNumber_ %></strong></font></td>
</tr>   
<tr>
	<td>Remark </td>
	<td>:</td>
	<td><font color=blue><strong><%=Remark_%></strong></font></td>
</tr>   
<tr>
	<td colspan="3" >&nbsp;</td>
</tr>
<tr>
	<td colspan="3" align=center>
		<input type="Submit" value="Yes" id="btnDelete"> 
		<input type="button" value="Cancel" id="btnCancel" onclick="self.history.back()"> 
	</td>
</tr>
<tr>
	<td colspan="3">
		<input type="hidden" name="txtID" value=<%=ID_ %>>
	</td>
</tr>
</table>
</form>
</body>
</html>