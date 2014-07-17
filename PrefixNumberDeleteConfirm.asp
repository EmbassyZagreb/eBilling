<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% dim PrefixID_
   PrefixID_ =  trim(request("PrefixID"))
   State_ =  trim(request("State"))
'response.write PrefixID_

	strsql = "select PrefixID, Code, Prefix, Type, Description from MsPrefixNumber Where PrefixID =" & PrefixID_
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	set rs = BillingCon.execute(strsql) 
	if not rs.eof then 
		Code_ = rs("Code") 
		Prefix_ = rs("Prefix")  
		Type_ = rs("Type")
		Description_ = rs("Description")
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
  	<TD COLSPAN="4" ALIGN="center" Class="SubTitle">PREFIX NUMBER DELETED</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" action="PrefixNumberDelete.asp"> 
<table cellpadding="1" cellspacing="0" width="100%" >  
<tr>
	<td align="right" width="50%">Code :</td>
	<td><label class="Label">&nbsp;<%=Code_ %></label></td>
</tr>
<tr>
	<td align="right">Prefix :</td>
	<td><label class="Label">&nbsp;<%=Prefix_%></label></td>
</tr>
<tr>
	<td align="right">Type :</td>
	<td><label class="Label">&nbsp;<%=Type_%></label></td>
</tr>
<tr>
	<td align="right">Description :</td>
	<td><label class="Label">&nbsp;<%=Description_%></label></td>
</tr>
<tr>
	<td colspan="2" align=center><br></td>
</tr>   
<tr>
	<td colspan="2" align=center>Are you sure will delete this data, Continue ?</td>
</tr>   
<tr>
	<td colspan="2" align=center>
		<input type="Submit" value="Yes" id="btnDelete"> 
		<input type="button" value="Cancel" id="btnCancel" onclick="self.history.back()"> 
	</td>
</tr>
<tr>
	<td colspan="2">
		<INPUT TYPE="HIDDEN" NAME="txtPrefixID" value='<%=PrefixID_%>'>
		<INPUT TYPE="HIDDEN" NAME="txtState" value='<%=State_%>'>
	</td>
</tr>
</table>
</form>
</body>
</html>