<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmPersonalPhone.txtPhoneNumber.value == "" )
	{
		msg = msg + "Please fill in Phone Number!!!\n"
		valid = false;
	}

	if (valid == false)
	{
		alert(msg);
	}
	return valid;
}
</script>
<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>
<form method="post" name="frmPersonalPhone" action="PersonalPhoneSaveAddOn.asp" onsubmit="return validate_form();"> 
<%  
 dim rst 
 dim strsql

ID_ = request("ID")
State_ = request("State")

strsql = "select RoleID from Users where loginId ='" & user1_ & "'"
set UserRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set UserRS = BillingCon.execute(strsql)
if not UserRS.eof then
	UserRole_ = UserRS("RoleID")
Else
	UserRole_ = ""
end if

  if not UserRS.eof then 
'     if (trim(rst("RoleID")) = "Admin") or (trim(rst("RoleID")) = "IM") then  
      If (UserRole_ <> "") Then
	if State_ = "E" then
	        strsql = " select * from MsPersonalPhone where ID = " & ID_
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
        	   PhoneNumber_ = rst1("PhoneNumber") 
		   Remark_ = rst1("Remark")
        	end if
       	end if
	'response.write State_ & "<br>"
%>             
	<table align=center>  
	<tr>
	  <td>Phone Number</td>
	  <td width="1px">:</td>
	  <td><input type="input" name="txtPhoneNumber" value='<%=PhoneNumber_ %>' size="30" maxlength="50" /></td>
	</tr> 
	<tr>
	  <td>Remark</td>
	  <td>:</td>
	  <td><input type="input" name="txtRemark" value='<%=Remark_ %>' size="60" maxlength="100" /></td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td width="1%">&nbsp;</td>
	  <td><input type="submit" name="btnSubmit" value="Submit">
		<%if State_= "E" then %>
		      <input type="hidden" name="txtID" value='<%=ID_ %>'>
		<%End If%>
	      <input type="hidden" name="State" value=<%=State_ %> >
	      &nbsp;<input type="button" value="Cancel" name="btnCancel" onClick="Javascript:history.go(-1)">
	 </td>
	</tr>  
	<tr>
		<td colspan=3>&nbsp;</td></tr>
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
<%   end if 
else %>
	<table align="center">
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<%end if %>
</form>
</BODY>
</html>