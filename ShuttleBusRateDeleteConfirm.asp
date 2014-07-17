<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% dim ShuttleBusRateID_
   ShuttleBusRateID_ =  trim(request("ShuttleBusRateID"))
   State_ =  trim(request("State"))
'response.write ShuttleBusRateID_ 
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
  	<TD COLSPAN="4" ALIGN="center" Class="title">SHUTTLE BUS RATE DELETE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" action="ShuttleBusRateDelete.asp"> 
<table cellpadding="1" cellspacing="0" width="100%">  
<tr>
	<td colspan="2" align=center>This record will be deleted, Continue ?</td>
</tr>   
<tr>
	<td colspan="2" align=center>
		<input type="Submit" value="Yes" id="btnDelete"> 
		<input type="button" value="Cancel" id="btnCancel" onclick="self.history.back()"> 
	</td>
</tr>
<tr>
	<td colspan="2">
		<INPUT TYPE="HIDDEN" NAME="txtShuttleBusRateID" value='<%=ShuttleBusRateID_%>'>
		<INPUT TYPE="HIDDEN" NAME="txtState" value='<%=State_%>'>
	</td>
</tr>
</table>
</form>
</body>
</html>