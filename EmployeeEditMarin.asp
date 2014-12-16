<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>


<script language="JavaScript">
function validate_form()
{
	valid = true;
	msg="";
	if (document.frmEmployee.cmbReportTo.value == "" )
	{
		msg = msg + "Please select Supervisor !!!\n"
		valid = false;
	}

	if (document.frmEmployee.cmbFundingAgency.value == "" )
	{
		msg = msg + "Please select Funding Agency !!!\n"
		valid = false;
	}

	if (document.frmEmployee.txtEmailAddress.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmEmployee.txtEmailAddress.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for email address !!!\n"
			valid = false;
		}
	}

	if (document.frmEmployee.txtAlternateEmail.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmEmployee.txtAlternateEmail.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for alternative email address !!!\n"
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
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

 EmpID_ = Request("EmpID")
 EmpType_ = Request("Type")

%> 
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">EMPLOYEE LIST</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<form method="post" action="EmployeeSaveMarin.asp" name="frmEmployee" onsubmit="return validate_form()"> 
<%  
 dim rst 
 dim strsql

 strsql = "Select EmpID, EmpName, PostName, Agency, Office, Type, ReportToID, ReportToName, EmailAddress, LoginID, AlternateEmail, AgencyId, AgencyFundingCode, AgencyFunding From vwDirectReport where EmpID ='" & EmpID_ & "' And Type='" & EmpType_ & "'"
 'response.write strsql 
  set rsData = server.createobject("adodb.recordset") 
  set rsData = BillingCon.execute(strsql)
  if not rsData.eof then 
	EmpID_ = rsData("EmpID")
	EmpName_ = rsData("EmpName")
	Post_ = rsData("PostName")
	Agency_ = rsData("Agency")
	Office_ = rsData("Office")
	Type_ = rsData("Type")
	ReportToID_ = rsData("ReportToID")
	EmailAddress_ = rsData("EmailAddress")
	LoginID_ = rsData("LoginID")
	AlternateEmail_ = rsData("AlternateEmail")
	AgencyFundingCodeEmp_ = rsData("AgencyId")
  end if
 'response.write ReportToID_ 
%>             
<table align="center">
<tr>
  <td>Employee Name :</td> 
  <td><input type="input" name="txtName" value='<%=EmpName_ %>' size="50" maxlength="50" /></td>
</tr>
<tr>
  <td>Post :</td>
  <td><input type="input" name="txtPost" value='<%=Post_ %>' size="50" maxlength="50" /></td>
</tr>
<tr>
  <td>Agency :</td>
  <td><input type="input" name="txtAgency" value='<%=Agency_ %>' size="50" maxlength="50" /></td>
</tr>
<tr>
  <td>Office :</td>
  <td><input type="input" name="txtOffice" value='<%=Office_ %>' size="50" maxlength="50" /></td>
</tr>
<tr>
  <td>Type :</td>
  <td><input type="input" name="txtType" value='<%=Type_ %>' size="50" maxlength="50" /></td>
</tr>
<tr>
  <td>Email Address :</td>
  <td><input type="input" name="txtEmailAddress" size="50" value="<%=EmailAddress_%>"/>
  </td>
</tr>
<tr>
  <td>Login ID :</td>
  <td><input type="input" name="txtLoginID" size="50" value="<%=LoginID_%>"/>
  </td>
</tr>
<tr>
  <td>Supervisor :</td>
  <td>
	<select id="cmbReportTo" name="cmbReportTo">
		<option value="">-- Select --</option>
<%
		Dim UserRS
		strsql = "Select EmpID, ISNULL(EmpName,'')+' - '+ISNULL(Office,'') As EmpName from vwPhoneCustomerList Where LEN(ISNULL(EmpName,''))<>''  Order by EmpName"
		response.write strsql & "<br>"
		set UserRS = server.createobject("adodb.recordset")
		set UserRS =BillingCon.execute(strsql)				        
		do while not UserRS.eof
			EmpIDX_ = UserRS("EmpID")
			Ename_ = UserRS("EmpName")
%>
	        <OPTION value='<%=EmpIDX_%>' <%if (trim(EmpIDX_) = trim(ReportToID_ )) then%>selected<%end if%>><%= EName_  %>
<%
                 UserRS.movenext
	        loop
%>  
	</select>	
  </td>
</tr>
<tr>
  <td>Alternate Email  :</td>
  <td><input id="txtAlternateEmail" name="txtAlternateEmail" value='<%=AlternateEmail_ %>' size="50" </input>
  </td>
</tr>

<tr>
  <td>Funding Agency :</td>
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
	        <OPTION value='<%=AgencyFundingCode_%>' <%if (trim(AgencyFundingCodeEmp_) = trim(AgencyFundingCode_ )) then%>selected<%end if%>><%= AgencyFunding_   %>
<%
                 AgencyRS.movenext
	        loop
%>  
	</select>	
  </td>
</tr>
<tr>
  <td colspan="2"><br></td>
</tr>
<tr>
  <td></td>
  <td><input type="submit" name="btnSubmit" value="Update">
      <input type="hidden" name="txtEmpID" value=<%=EmpID_  %>>
      <input type="hidden" name="txtEmpType" value=<%=Type_ %>>
      &nbsp;<input type="button" value="Cancel" name="btnCancel" onClick="javascript:location.href='EmployeeListMarin.asp'">
 </td>
</tr>  
<tr><td colspan=2>&nbsp;</td></tr>
</table>
</form>
</BODY>
</html>