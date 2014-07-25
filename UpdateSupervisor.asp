
<!--#include file="connect.inc" -->
<html>
<head>
<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmUpdateSupervisor.cmbEmpFrom.value == "" )
	{
		msg = msg + "Please select From Supervisor."
		valid = false;
	}

	if (document.frmUpdateSupervisor.cmbEmpTo.value == "" )
	{
		msg = msg + "Please select To Supervisor. "
		valid = false;
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
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">UPDATE SUPERVISOR</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" name="frmUpdateSupervisor" action="UpdateSupervisorSave.asp" onsubmit="return validate_form();">	
<%  
' State_ = "I"
  strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
     if trim(rst("RoleID")) = "Admin" or trim(rst("RoleID")) = "FMC" or trim(rst("RoleID")) = "Voucher" then 
'response.write State_ & "<br>"
%>             
<table align=center>  
<tr>
	<td>From Supervisor</td>
	<td>:</td>
	<td>
<%
 				strsql ="select EmpID, EmpName from vwPhoneCustomerList Where EmpType = 'AMER' order by EmpName"
				set EmpRS = server.createobject("adodb.recordset")
				set EmpRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbEmpFrom">
					<Option value=''>--Select--</Option>
<%				Do While not EmpRS.eof 
%>
					<Option value='<%=EmpRS("EmpID")%>' <%if trim(EmpID_) = trim(EmpRS("EmpID")) then %>Selected<%End If%> ><%=EmpRS("EmpName") %></Option>
					
<%					EmpRS.MoveNext
				Loop%>
				</select>

			</td>	
</tr>
<tr>
	<td>To Supervisor</td>
	<td>:</td>
	<td>
<%
 				strsql ="select EmpID, EmpName from vwPhoneCustomerList Where EmpType = 'AMER' order by EmpName"
				set EmpRS = server.createobject("adodb.recordset")
				set EmpRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbEmpTo">
					<Option value=''>--Select--</Option>
<%				Do While not EmpRS.eof 
%>
					<Option value='<%=EmpRS("EmpID")%>' <%if trim(EmpID_) = trim(EmpRS("EmpID")) then %>Selected<%End If%> ><%=EmpRS("EmpName") %></Option>
					
<%					EmpRS.MoveNext
				Loop%>
				</select>

	</td>	
</tr>
<tr>
  <td colspan="2"></td>
  <td><input type="submit" name="btnSubmit" value="Update"> </td>
</tr>  
<tr><td colspan=2>&nbsp;</td></tr>
</table>
<%
   else 
%>
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

<%
end if 
%>
</form>
</BODY>
</html>