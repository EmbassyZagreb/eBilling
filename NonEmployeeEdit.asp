<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmNonEmployee.txtPhoneNumber.value == "" )
	{
		msg = msg + "Please fill in Phone Number!!!\n"
		valid = false;
	}


	if (document.frmNonEmployee.txtAlternateEmail.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmNonEmployee.txtAlternateEmail.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for alternative email address !!!\n"
			valid = false;
		}
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
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Non Employee Update</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" name="frmNonEmployee" action="NonEmployeeSave.asp" onsubmit="return validate_form();"> 
<%  
 dim rst 
 dim strsql

NonEmpID_ = request("NonEmpID")
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
	        strsql = " select * from MsNonEmployee where NonEmpID = '" & NonEmpID_ & "'"
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
        	   NonEmpName_ = rst1("NonEmpName") 
	  	   AgencyID_ = rst1("AgencyId")
		   Email_ = rst1("Email") 
		   Remark_ = rst1("Remark")
		   Status_ = rst1("Status")
        	end if
       	end if
	'response.write State_ & "<br>"
%>             
	<table align=center>  
	<tr>
	  <td>Name</td>
	  <td width="1%">:</td>
	  <td><input type="input" name="txtName" value='<%=NonEmpName_ %>' size="50" maxlength="50" /></td>
	</tr> 
	<tr>
	  <td>Funding Agency</td>
	  <td>:</td>
	  <td>
		<select id="cmbFundingAgency" name="cmbFundingAgency">
			<option value="">-- Select --</option>
<%
			Dim AgencyRS
			strsql = "Select AgencyId, AgencyFundingCode, AgencyDesc from AgencyFunding Where Disabled='N' Order by AgencyDesc"
			'response.write strsql & "<br>"
			set AgencyRS = server.createobject("adodb.recordset")
			set AgencyRS =BillingCon.execute(strsql)				        
			do while not AgencyRS.eof
				AgencyFundingCode_ = AgencyRS("AgencyId")
				AgencyFunding_ = AgencyRS("AgencyDesc")
%>	
		        <OPTION value='<%=AgencyFundingCode_%>' <%if (trim(AgencyID_ ) = trim(AgencyFundingCode_ )) then%>selected<%end if%>><%= AgencyFunding_   %>
<%
  	               AgencyRS.movenext
		        loop
%>  
		</select>	
	  </td>
	</tr>
	<tr>
	  <td>Email</td>
	  <td width="1%">:</td>
	  <td><input type="input" name="txtEmail" value='<%=Email_ %>' size="50" maxlength="50" /></td>
	</tr>
	<tr>
	  <td>Remark</td>
	  <td>:</td>
	  <td><input type="input" name="txtRemark" value='<%=Remark_ %>' size="100" maxlength="100" /></td>
	</tr>
	<tr>
	  <td>Status</td>
	  <td>:</td>
	  <td>
		  <select name="cmbStatus">
			<option value="C" <%if Status_ ="C" Then %>Selected<%End If%>>Current</option>
			<option value="D" <%if Status_ ="D" Then %>Selected<%End If%>>Departed</option>
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td width="1%">&nbsp;</td>
	  <td><input type="submit" name="btnSubmit" value="Submit">
		<%if State_= "E" then %>
		      <input type="hidden" name="txtNonEmpID" value='<%=NonEmpID_ %>'>
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