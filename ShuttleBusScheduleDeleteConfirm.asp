<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% 
   EmpID_ = request("EmpID")
   ShuttleDate_ = request("ShuttleDate")
   State_ =  trim(request("State"))
'response.write ID_

	strsql = " select EmpName, Agency, Office, AM, PM from vwShuttleSchedule where EmpID = " & EmpID_ & " and ShuttleDate = '" & ShuttleDate_ & "'"
  	set rsData = server.createobject("adodb.recordset") 
	'response.write strsql 
       	set rsData = BillingCon.execute(strsql)
       	if not rsData.eof then 
      		EmpName_ = rsData("EmpName")
       		Agency_ = rsData("Agency")
       		Office_ = rsData("Office")
       		AM_ = rsData("AM")
		PM_ = rsData("PM")
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
  	<TD COLSPAN="4" ALIGN="center" Class="title">Shuttle Bus Schedule Deleted</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" action="ShuttleBusScheduleDelete.asp"> 
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
		<td>AM</td>
		<td>:</td>
		<td><font color=blue><strong><%=AM_ %></strong></font></td>
	</tr>   
	<tr>
		<td>PM</td>
		<td>:</td>
		<td><font color=blue><strong><%=PM_%></strong></font></td>
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
		<input type="hidden" name="txtEmpID" value=<%=EmpID_ %>>
		<input type="hidden" name="txtShuttleDate" value=<%=ShuttleDate_ %>>
	</td>
</tr>
</table>
</form>
</body>
</html>