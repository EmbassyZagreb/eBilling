<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% dim ID_
   ID_ =  trim(request("ID"))
   State_ =  trim(request("State"))
'response.write ID_

	strsql = "select PhoneNumber from MsPhoneList Where ID=" & ID_
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	set rs = BillingCon.execute(strsql) 
	if not rs.eof then 
		PhoneNumber_ = rs("PhoneNumber")
	end if
%>

<html>
   <head>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
   </head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PHONE LIST</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<BODY bgcolor=white background=images/bg07.gif alink=blue link=blue vlink=blue >
<form method="post" action="PhoneDelete.asp"> 
<table>
<tr>
	<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%=session("Month")%> - <%=session("Year")%></lable></td>
</tr>
<tr>
	<td colspan="2"><br></td>
</tr>
<tr>
	<td colspan="2" align=center>Phone number : <font color=blue><strong><%=PhoneNumber_  %></strong></font> will be deleted, Continue ?</td>
</tr>   
<tr>
	<td colspan="2" align=center>
		<input type="Submit" value="Yes" id="btnDelete"> 
		<input type="button" value="Cancel" id="btnCancel" onclick="self.history.back()"> 
	</td>
</tr>
<tr>
	<td colspan="2">
		<INPUT TYPE="HIDDEN" NAME="txtID" value='<%=ID_%>'>
		<INPUT TYPE="HIDDEN" NAME="txtState" value='<%=State_%>'>
	</td>
</tr>
</table>
</form>
</body>
</html>