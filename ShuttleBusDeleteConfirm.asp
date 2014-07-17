<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% 
   ShuttleID_ =  trim(request("ShuttleID"))
   State_ =  trim(request("State"))
'response.write ID_

	strsql = "select * from vwShuttleBus Where ShuttleID =" & ShuttleID_
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	set rs = BillingCon.execute(strsql) 
	if not rs.eof then 
		EmpName_ = rs("EmpName")
		Agency_ = rs("Agency")
		Office_ = rs("Office")
		TransportDate_ = rs("TransportDate")
		EventType_ = rs("EventType")
		QtyPerson_ = rs("QtyPerson")
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
  	<TD COLSPAN="4" ALIGN="center" Class="title">Shuttle Bus Bill</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" action="ShuttleBusDelete.asp"> 
<table cellpadding="1" cellspacing="0" width="100%">  
<tr>
	<td colspan="2" align=center>Are you sure delete this data ?</td>
</tr>   
<tr>
	<td colspan="2" align="center">
	<table width="40%" cellspadding="0" cellspacing="0">
	<tr>
		<td width="35%">Employee Name </td>
		<td>:</td>
		<td><font color=blue><strong><%=EmpName_  %></strong></font></td>
	</tr>   
	<tr>
		<td>Agency / Office </td>
		<td>:</td>
		<td><font color=blue><strong><%=Agency_%> / <%=Office_%></strong></font></td>
	</tr>   
	<tr>
		<td>Transport Date</td>
		<td>:</td>
		<td><font color=blue><strong><%=TransportDate_ %></strong></font></td>
	</tr>   
	<tr>
		<td>Time</td>
		<td>:</td>
		<td><font color=blue><strong><%=EventType_ %></strong></font></td>
	</tr>   
	<tr>
		<td>Qty. Person</td>
		<td>:</td>
		<td><font color=blue><strong><%=QtyPerson_%></strong></font></td>
	</tr>   
	</table>
	</td>
</tr>
<tr>
	<td colspan="2" >&nbsp;</td>
</tr>
<tr>
	<td colspan="2" align=center>
		<input type="Submit" value="Yes" id="btnDelete"> 
		<input type="button" value="Cancel" id="btnCancel" onclick="self.history.back()"> 
	</td>
</tr>
<tr>
	<td colspan="2">
		<INPUT TYPE="HIDDEN" NAME="txtShuttleID" value='<%=ShuttleID_%>'>
		<INPUT TYPE="HIDDEN" NAME="txtState" value='<%=State_%>'>
	</td>
</tr>
</table>
</form>
</body>
</html>