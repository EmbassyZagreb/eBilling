<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="calendar.js"></script>
<script language="vbscript">
<!--
Sub btnCancel_onclick
          history.back
End Sub
--> 

<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmData.txtAM.value == "" )
	{
		msg = msg + "Please fill in AM, minimum value is 0  !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmData.txtAM.value) == false)
		{
			msg = msg + "Invalid data type for AM !!!\n"
			valid = false;
		}
	}

	if (document.frmData.txtPM.value == "" )
	{
		msg = msg + "Please fill in PM, minimum value is 0  !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmData.txtPM.value) == false)
		{
			msg = msg + "Invalid data type for PM !!!\n"
			valid = false;
		}
	}


	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
</script>
<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = right(user_,len(user_)-4)
'response.write user1_ & "<br>"

%> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Shuttle Bus Schedule Update</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form name="frmData" align="center" method="post" action="ShuttleBusScheduleSave.asp" onSubmit="return validate_form()"> 
<%  
 dim rsUser
 dim strsql
 dim rsData

 EmpID_ = request("EmpID")
 ShuttleDate_ = request("ShuttleDate")	
 State_ = request("State")
 If State_ = "" then
	 State_ = "I"
 End If
  strsql = "select * from Users where LoginID='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set UserRS = server.createobject("adodb.recordset") 
  set UserRS = BillingCon.execute(strsql)
  if not UserRS.eof then 
	if (trim(UserRS("RoleID")) = "Admin") or (trim(UserRS("RoleID")) = "Trs") then  
		'response.write State_ & "<br>"
		if State_ ="E" Then
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
		end if
		'response.write EmpID_ 
%>             

	<table align="center" cellpadding="1" cellspacing="0" width="80%">
	<tr>
		<td align="right">Employee Name</td>
		<td width="1%">:</td>
		<td class="FontContent"><%=EmpName_%></td>
	</tr>
	<tr>
		<td align="right">Agency / Office</td>
		<td width="1%">:</td>
		<td class="FontContent"><%=Agency_ %> / <%=Office_ %></td>
	</tr>
	<tr>
		<td align="right">Shuttle Date</td>
		<td width="1%">:</td>
		<td class="FontContent"><%=formatdatetime(ShuttleDate_,1) %></td>
	</tr>
	<tr>
		<td align="right">AM</td>
		<td width="1%">:</td>
		<td><input name="txtAM" size="2" value='<%=AM_%>' /></td>
	</tr>
	<tr>
		<td align="right">PM</td>
		<td width="1%">:</td>
		<td><input name="txtPM" size="2" value='<%=PM_%>' /></td>
	</tr>
	<tr>
	  	<td colspan="3"><br></td>
	</tr>
	<tr>
	  	<td>&nbsp;</td>
	  	<td colspan="2"><input type="submit" name="btnSubmit" value="Submit">
		      <input type="hidden" name="txtEmpID" value=<%=EmpID_ %>>
		      <input type="hidden" name="txtShuttleDate" value=<%=ShuttleDate_ %>>
		      &nbsp;<input type="button" value="Cancel" name="btnCancel">
		 </td>
	</tr>  
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	</table>
	<%else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>

		</tr>
	</table>
<% end if
else %>
	<table align="center">
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>

		</tr>
	</table>
<% end if %>
</form>

</BODY>
</html>